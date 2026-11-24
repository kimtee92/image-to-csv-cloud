import base64
import csv
import io
import json
import logging
import os
import re
from datetime import date, timedelta
from html.parser import HTMLParser
from pathlib import Path
from typing import Optional, Any

from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.staticfiles import StaticFiles
from fastapi.responses import Response
from pydantic import BaseModel
import httpx
from PIL import Image, ImageOps, UnidentifiedImageError
import pypdfium2 as pdfium
import google.auth
from google.auth.transport.requests import Request as AuthRequest

import clients as client_registry
import holidays as holidays_cal
import timekeeping
import xlsx_fill

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = FastAPI(title="Image to CSV")

# Initialize Google Cloud credentials (ADC)
try:
    credentials, project_id = google.auth.default(scopes=["https://www.googleapis.com/auth/cloud-platform"])
except Exception as e:
    logger.warning(f"Failed to load Google credentials: {e}")
    credentials = None

GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "")
GEMINI_MODEL = os.getenv("GEMINI_MODEL", "google/gemini-3.1-flash-lite")
GEMINI_BASE_URL = os.getenv(
    "GEMINI_BASE_URL", "https://generativelanguage.googleapis.com/v1beta/openai"
)

def _gemini_headers() -> dict[str, str]:
    """Always use Google Cloud Service Account credentials (ADC) if available, fallback to API key."""
    if credentials:
        if not credentials.valid:
            credentials.refresh(AuthRequest())
        return {"Authorization": f"Bearer {credentials.token}"}
    
    if GEMINI_API_KEY:
        return {"Authorization": f"Bearer {GEMINI_API_KEY}"}
        
    raise ValueError("No valid Google Cloud credentials (ADC) or GEMINI_API_KEY found.")

MAX_FILE_SIZE = 20 * 1024 * 1024  # 20 MB
# Keep uploaded images reasonably sized for fast, cheap Gemini vision calls.
MAX_IMAGE_PIXELS = 3_600_000

# Uploads get orientation + optional upscale only, sent as PNG. Nothing is
# denoised, greyscaled or JPEG-compressed, so small digits keep their exact pixels
# (6 vs 8 survives) and coloured cells keep their contrast. The old greyscale /
# denoise prepass measured materially WORSE on these forms (83.3% vs 93.1%) and
# has been removed rather than left as a switch someone might flip.
# Upscale small scans so thin strokes span more pixels; capped to keep payloads sane.
LOSSLESS_MIN_SIDE = int(os.getenv("LOSSLESS_MIN_SIDE", "2000"))
LOSSLESS_MAX_PIXELS = int(os.getenv("LOSSLESS_MAX_PIXELS", "12000000"))
# Second-pass arithmetic verification of in/out/total columns (see _verify_time_math).
TIME_VERIFY = os.getenv("TIME_VERIFY", "1").strip().lower() not in {"0", "false", "no"}
# Second opinion on every DTR row from a magnified band crop, with disagreements
# flagged for review. Off by default because it roughly doubles the per-file cost.
# Measured on the 10 reference2 scans: the two readings agree on 249/252 cells, and
# of the 3 disagreements the crop was right once, the full page once, and both were
# wrong once — so the crop is NOT a reliable arbiter and nothing is overwritten on
# its say-so. What it buys is detection: 2 of the 3 residual errors that survive the
# chronology check are chronologically valid, so this is the only thing that spots
# them, for 3 flags per 252 cells.
DTR_CROSSCHECK = os.getenv("DTR_CROSSCHECK", "0").strip().lower() in {"1", "true", "yes"}
# How much the model is allowed to deliberate. "minimal" measured best on the
# flash-lite models — raising it to "low" made extraction WORSE there (95.6% vs
# 98.8%, from dropped rows). It is configurable because the level is not portable:
# gemini-3.8-flash rejects "minimal" outright ("Thinking level MINIMAL is not
# supported"), which 400s every call, so switching model means setting this too.
# Empty string omits the field entirely.
REASONING_EFFORT = os.getenv("REASONING_EFFORT", "minimal").strip().lower()
# Escalation: what to do with a row the cheap first pass could not settle.
#
# Deliberately NOT a model upgrade by default. Measured on the hand-read benchmark
# (benchmark/ground_truth.json), the model tier buys nothing here — 3.1-lite,
# 3.5-lite and 3.8-flash score 96.9 / 98.4 / 96.9%, identical on two of three
# pages and one cell apart on the third. Raising reasoning_effort measured worse
# still. The one lever that demonstrably works is MAGNIFICATION: a tight crop
# recovered a stable misread 6/6 where the full page failed 6/6.
#
# So escalation means "look closer", and only optionally "ask a better model".
# ESCALATION_MODEL="" keeps the same model and just zooms further in.
ESCALATION_MODEL = os.getenv("ESCALATION_MODEL", "").strip()
DTR_ESCALATE = os.getenv("DTR_ESCALATE", "1").strip().lower() not in {"0", "false", "no"}
# Cloud Run kills the request at its own timeout (300s by default), so the OCR call
# plus the verification re-read have to fit inside that with room to spare — otherwise
# a slow pair of calls is dropped with no result instead of returning the table.
OCR_TIMEOUT = float(os.getenv("OCR_TIMEOUT", "170"))
VERIFY_TIMEOUT = float(os.getenv("VERIFY_TIMEOUT", "90"))



def _orientation_candidates(img: Image.Image) -> tuple[int, int]:
    """The two rotations worth considering, from the image's shape.

    A portrait scan of a landscape page is either 90 or 270 out; a landscape one is
    either correct or upside down. Narrowing to two turns orientation into a single
    binary question.
    """
    w, h = img.size
    return (90, 270) if h > w else (0, 180)


def _rotation_thumb_b64(img: Image.Image, angle: int, size: int = 1100) -> str:
    rotated = _rotate_image(img, angle)
    rotated.thumbnail((size, size), getattr(getattr(Image, "Resampling", Image), "LANCZOS"))
    buf = io.BytesIO()
    rotated.save(buf, format="JPEG", quality=88)
    return base64.b64encode(buf.getvalue()).decode()


async def _detect_orientation_via_gemini(img: Image.Image) -> dict:
    """Decide which way up a scan belongs, by comparing its two candidate rotations.

    Asking the model which edge the top faces does not work — measured at 35-80%
    across every input size tried, i.e. barely better than a coin flip on the
    90-vs-270 axis, which would auto-rotate a large share of files upside down.
    Showing it both candidates and asking which one reads correctly is a far easier
    discrimination and measured 28/28 on the reference scans.

    Returns: {"needs_rotation": bool, "rotation_angle": int, "reasoning": str}
    """
    first, second = _orientation_candidates(img)
    instruction = (
        "You are shown the SAME document twice, rotated differently: image A then image B.\n"
        "Exactly one of them is the right way up, so that its text reads normally "
        "left-to-right and its title is at the top.\n\n"
        'Answer ONLY: {"upright": "A"|"B"}'
    )
    content = [
        {"type": "text", "text": "Image A:"},
        {"type": "image_url",
         "image_url": {"url": f"data:image/jpeg;base64,{_rotation_thumb_b64(img, first)}"}},
        {"type": "text", "text": "Image B:"},
        {"type": "image_url",
         "image_url": {"url": f"data:image/jpeg;base64,{_rotation_thumb_b64(img, second)}"}},
        {"type": "text", "text": instruction},
    ]
    payload = {
        "model": await get_model_name(),
        "messages": [{"role": "user", "content": content}],
        "max_tokens": 100,
        "temperature": 0.0,
        "response_format": {"type": "json_object"},
    }
    
    try:
        async with httpx.AsyncClient(timeout=30.0) as client:
            resp = await client.post(
                f"{GEMINI_BASE_URL}/chat/completions", json=payload, headers=_gemini_headers()
            )
            resp.raise_for_status()
            raw = resp.json()["choices"][0]["message"]["content"]
        
        result = _extract_json_object(raw)
        if result:
            pick = str(result.get("upright", "")).strip().upper()[:1]
            if pick not in {"A", "B"}:
                logger.warning("Unexpected orientation choice %r; assuming upright", pick)
                return {"needs_rotation": False, "rotation_angle": 0,
                        "reasoning": "Unclear response, left as-is"}
            angle = first if pick == "A" else second
            return {
                "needs_rotation": angle != 0,
                "rotation_angle": angle,
                "reasoning": f"Chose image {pick} ({angle} deg clockwise) as upright",
            }
    except Exception as e:
        logger.warning("Orientation detection failed (%s): %s", type(e).__name__, e)
    
    # Default: no rotation needed
    return {"needs_rotation": False, "rotation_angle": 0, "reasoning": "Detection failed, assuming correct orientation"}


def _rotate_image(img: Image.Image, angle: int) -> Image.Image:
    """Rotate an image by the specified angle (90, 180, or 270 degrees clockwise)."""
    if angle == 0:
        return img
    if angle == 90:
        return img.rotate(270, expand=True)  # PIL rotates counter-clockwise
    if angle == 180:
        return img.rotate(180, expand=True)
    if angle == 270:
        return img.rotate(90, expand=True)
    return img


if not GEMINI_API_KEY:
    logger.warning("GEMINI_API_KEY is not set — /api/ocr and /api/amend will fail")


class AmendRequest(BaseModel):
    instruction: str
    headers: list[Any] = []
    rows: list[Any] = []
    header_fields: list[Any] = []
    footer_fields: list[Any] = []
    ocr_context: str = ""  # raw JSON output of the original OCR extraction


async def get_model_name() -> str:
    return GEMINI_MODEL



def _prepare_lossless(img: Image.Image) -> Image.Image:
    """Orientation-correct and upscale, with no filtering or colour loss.

    Deliberately does nothing else. The old enhancement prepass measurably hurt
    dense forms: a 3x3 median filter closes the gap in small digits (8 reads as 6)
    and greyscaling collapses black-on-red highlighted cells to near-zero contrast.
    """
    img = ImageOps.exif_transpose(img)
    if img.mode != "RGB":
        img = img.convert("RGB")

    resampling = getattr(getattr(Image, "Resampling", Image), "LANCZOS")
    w, h = img.size
    min_side = min(w, h)
    if min_side < LOSSLESS_MIN_SIDE:
        scale = LOSSLESS_MIN_SIDE / float(min_side)
        img = img.resize((int(round(w * scale)), int(round(h * scale))), resampling)

    w, h = img.size
    if w * h > LOSSLESS_MAX_PIXELS:
        scale = (LOSSLESS_MAX_PIXELS / float(w * h)) ** 0.5
        img = img.resize((int(w * scale), int(h * scale)), resampling)
    return img


def _encode_for_model(img: Image.Image) -> tuple[bytes, str]:
    """Unfiltered PNG — see the note on preprocessing above."""
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue(), "image/png"


def _preprocess_image(raw_bytes: bytes) -> tuple[bytes, str]:
    return _encode_for_model(_prepare_lossless(Image.open(io.BytesIO(raw_bytes))))


def _exif_rotated(raw_bytes: bytes) -> bool:
    """True when an EXIF tag rotates this image, i.e. it is already upright for us.

    Lets the frontend tell apart uploads that the browser has already straightened
    from ones whose stored pixels are sideways, without guessing from the picture.
    """
    try:
        tag = Image.open(io.BytesIO(raw_bytes)).getexif().get(274)
    except Exception:
        return False
    return tag in {5, 6, 7, 8}  # the four orientations that transpose actually turns


def _is_pdf_upload(content_type: str | None, filename: str | None) -> bool:
    ct = (content_type or "").lower()
    name = (filename or "").lower()
    return ct in {"application/pdf", "application/x-pdf"} or name.endswith(".pdf")


def _is_image_upload(content_type: str | None, filename: str | None) -> bool:
    ct = (content_type or "").lower()
    if ct.startswith("image/"):
        return True
    ext = Path(filename or "").suffix.lower()
    return ext in {".png", ".jpg", ".jpeg", ".webp", ".bmp", ".tif", ".tiff", ".gif"}


def _pdf_page_count(raw_bytes: bytes) -> int:
    pdf = pdfium.PdfDocument(io.BytesIO(raw_bytes))
    try:
        return len(pdf)
    finally:
        pdf.close()


def _preprocess_pdf(raw_bytes: bytes, page_index: int = 0) -> tuple[bytes, str]:
    """Render one PDF page and prepare it for the model.

    `page_index` matters: several clients submit their whole period as a single
    multi-page PDF (Troner's is 7 pages, one employee each). Rendering only page 0
    silently delivered one employee and dropped the rest.
    """
    pdf = None
    page = None
    try:
        pdf = pdfium.PdfDocument(io.BytesIO(raw_bytes))
        if len(pdf) == 0:
            raise ValueError("PDF has no pages")
        if not 0 <= page_index < len(pdf):
            raise ValueError(f"PDF has {len(pdf)} page(s); page {page_index + 1} requested")
        page = pdf[page_index]
        pil_image = page.render(scale=3).to_pil()
        return _encode_for_model(_prepare_lossless(pil_image))
    finally:
        if page is not None:
            page.close()
        if pdf is not None:
            pdf.close()


def _prepare_upload_for_model(
    raw_bytes: bytes, content_type: str | None, filename: str | None,
    page_index: int = 0
) -> tuple[bytes, str]:
    """Normalize image or PDF uploads into a model-ready image payload."""
    if _is_pdf_upload(content_type, filename):
        return _preprocess_pdf(raw_bytes, page_index)
    return _preprocess_image(raw_bytes)


def _normalise_rows(headers: list[str], rows_in: list[Any]) -> list[list[str]]:
    width = max(1, len(headers))
    rows: list[list[str]] = []
    for row in rows_in or []:
        if isinstance(row, list):
            cells = [str(c) if c is not None else "" for c in row]
        else:
            cells = [str(row) if row is not None else ""]
        while len(cells) < width:
            cells.append("")
        cells = cells[:width]
        if any(c.strip() for c in cells):
            rows.append(cells)
    return rows


def _normalise_fields(fields_in: list[Any]) -> list[dict[str, str]]:
    out: list[dict[str, str]] = []
    for item in fields_in or []:
        label = ""
        value = ""
        if isinstance(item, dict):
            label = str(item.get("label", "")).strip()
            value = str(item.get("value", "")).strip()
        elif isinstance(item, list) and item:
            label = str(item[0]).strip() if len(item) > 0 else ""
            value = str(item[1]).strip() if len(item) > 1 else ""
        if label or value:
            out.append({"label": label, "value": value})
    return out


# ── HTML table parser ──────────────────────────────────────────────────────
class _HTMLTableParser(HTMLParser):
    """Extract rows from an HTML <table>, tracking thead/tbody and span attributes."""

    def __init__(self):
        super().__init__()
        self.thead: list[list[tuple[str, int, int]]] = []  # rows of (text, rowspan, colspan)
        self.tbody: list[list[tuple[str, int, int]]] = []
        self._section = "tbody"
        self._cur_row: list[tuple[str, int, int]] | None = None
        self._cur_cell: list[str] | None = None
        self._in_cell = False
        self._cur_rs = 1
        self._cur_cs = 1

    def handle_starttag(self, tag, attrs):
        d = dict(attrs)
        if tag == "thead":
            self._section = "thead"
        elif tag == "tbody":
            self._section = "tbody"
        elif tag == "tr":
            self._cur_row = []
        elif tag in ("td", "th"):
            self._cur_cell = []
            self._in_cell = True
            try:
                self._cur_rs = int(d.get("rowspan", 1))
            except (ValueError, TypeError):
                self._cur_rs = 1
            try:
                self._cur_cs = int(d.get("colspan", 1))
            except (ValueError, TypeError):
                self._cur_cs = 1

    def handle_endtag(self, tag):
        if tag in ("td", "th") and self._cur_row is not None:
            text = "".join(self._cur_cell or []).strip()
            self._cur_row.append((text, self._cur_rs, self._cur_cs))
            self._cur_cell = None
            self._in_cell = False
        elif tag == "tr" and self._cur_row is not None:
            dest = self.thead if self._section == "thead" else self.tbody
            dest.append(self._cur_row)
            self._cur_row = None

    def handle_data(self, data):
        if self._in_cell and self._cur_cell is not None:
            self._cur_cell.append(data)


def _expand_rows(raw_rows: list[list[tuple[str, int, int]]]) -> list[list[str]]:
    """Expand (text, rowspan, colspan) rows into a flat 2D list, honouring span attributes."""
    pending: dict[int, tuple[int, str]] = {}  # col -> (remaining_rows, text)
    result: list[list[str]] = []

    for raw_row in raw_rows:
        row_dict: dict[int, str] = {}
        col = 0
        for text, rs, cs in raw_row:
            # Drain pending rowspan cells that occupy columns at or before current col
            while col in pending:
                rem, val = pending[col]
                row_dict[col] = val
                if rem > 1:
                    pending[col] = (rem - 1, val)
                else:
                    del pending[col]
                col += 1
            # Fill colspan columns with this cell's text
            for c in range(cs):
                row_dict[col + c] = text
                if rs > 1:
                    pending[col + c] = (rs - 1, text)
            col += cs
        # Drain any trailing pending rowspan columns not reached by actual cells
        for pc in sorted(k for k in pending if k not in row_dict):
            rem, val = pending[pc]
            row_dict[pc] = val
            if rem > 1:
                pending[pc] = (rem - 1, val)
            else:
                del pending[pc]
        if row_dict:
            max_col = max(row_dict.keys()) + 1
            result.append([row_dict.get(i, "") for i in range(max_col)])

    return result


def _combine_header_rows(head_rows: list[list[str]]) -> list[str]:
    """Combine multi-row headers into single column names.

    Rowspan repeats are deduplicated; colspan groups are combined with their children.
    e.g. row0=["MORNING","MORNING"] + row1=["IN","OUT"] → ["MORNING IN","MORNING OUT"]
    """
    if not head_rows:
        return []
    if len(head_rows) == 1:
        return head_rows[0]
    num_cols = max(len(r) for r in head_rows)
    headers = []
    for col in range(num_cols):
        parts: list[str] = []
        seen: set[str] = set()
        for row in head_rows:
            val = row[col] if col < len(row) else ""
            if val and val not in seen:
                parts.append(val)
                seen.add(val)
        headers.append(" ".join(parts) if parts else f"Column {col + 1}")
    return headers


def _extract_non_table_text(raw: str) -> dict:
    """Split raw OCR output into text_above, table_html, text_below."""
    result = {"text_above": [], "text_below": []}

    # Find the <table...>...</table> block
    match = re.search(r"<table[\s>].*?</table>", raw, re.DOTALL | re.IGNORECASE)
    if not match:
        # Try markdown table
        lines = raw.strip().split("\n")
        above, below = [], []
        in_table = False
        past_table = False
        for line in lines:
            stripped = line.strip()
            if "|" in stripped and not past_table:
                in_table = True
            elif in_table and "|" not in stripped:
                in_table = False
                past_table = True
            if not in_table and not past_table and stripped:
                above.append(stripped)
            elif past_table and stripped:
                below.append(stripped)
        result["text_above"] = above
        result["text_below"] = below
        return result

    before = raw[: match.start()].strip()
    after = raw[match.end() :].strip()

    if before:
        result["text_above"] = [
            line.strip()
            for line in before.split("\n")
            if line.strip()
        ]
    if after:
        result["text_below"] = [
            line.strip()
            for line in after.split("\n")
            if line.strip()
        ]
    return result


def parse_html_table(html: str) -> dict | None:
    """Parse an HTML table string into headers + rows, handling colspan/rowspan correctly."""
    if "<table" not in html.lower():
        return None

    parser = _HTMLTableParser()
    parser.feed(html)

    # Expand thead rows into flat grid
    head_rows = _expand_rows(parser.thead)
    if head_rows:
        headers = _combine_header_rows(head_rows)
        data_rows = _expand_rows(parser.tbody)
    else:
        # No explicit thead — first tbody row is the header
        all_rows = _expand_rows(parser.tbody)
        if not all_rows:
            return None
        headers = all_rows[0]
        data_rows = all_rows[1:]

    if not headers:
        return None

    # Normalise column count
    n = max(len(headers), *(len(r) for r in data_rows) if data_rows else [0])
    while len(headers) < n:
        headers.append(f"Column {len(headers) + 1}")

    padded_rows = []
    for row in data_rows:
        if not any(c != "" for c in row):
            continue  # skip fully empty rows
        padded = (row + [""] * n)[:n]
        padded_rows.append(padded)

    return {"headers": headers, "rows": padded_rows}


def parse_markdown_table(md_text: str) -> dict:
    """Parse a markdown pipe table into headers + rows."""
    lines = [line.strip() for line in md_text.strip().split("\n") if line.strip()]

    table_lines = [line for line in lines if "|" in line]
    if not table_lines:
        return {"headers": ["Content"], "rows": [[line] for line in lines if line]}

    parsed = []
    for line in table_lines:
        cells = [c.strip() for c in line.split("|")]
        if cells and cells[0] == "":
            cells = cells[1:]
        if cells and cells[-1] == "":
            cells = cells[:-1]
        parsed.append(cells)

    if not parsed:
        return {"headers": ["Content"], "rows": [[md_text.strip()]]}

    headers = parsed[0]
    rows = []
    for row in parsed[1:]:
        if all(re.match(r"^[-:\s]+$", cell) for cell in row):
            continue
        while len(row) < len(headers):
            row.append("")
        rows.append(row[: len(headers)])

    return {"headers": headers, "rows": rows}


def parse_table(text: str) -> dict:
    """Try HTML first, then markdown, then plain-text fallback."""
    result = parse_html_table(text)
    if result and result["headers"]:
        return result
    return parse_markdown_table(text)



# ── Time arithmetic self-validation ───────────────────────────────────────
# Timesheet/DTR forms print their own totals: TOTAL = (time out - time in).
# That makes every in/out/total group a free correctness check on the OCR, which
# is how single-digit misreads (18:45 read as 16:45) get caught and repaired.
# A trailing AM/PM is tolerated: the prompt asks for 24-hour HH:MM, but the model
# does not always comply, and an unparsed time silently skips the chronology check
# instead of being validated.
# Seconds are optional but must be accepted: Upson's DTRs are machine-printed to
# the second ("7:48:49"), and a parser that rejects them makes every one of their
# time cells look like a non-time — which is how their TIME IN / TIME OUT columns
# came out blank while the CSV showed the times fine. Kept in step with
# xlsx_fill._HHMM, which has always accepted them.
_HHMM_RE = re.compile(r"^(\d{1,2}):(\d{2})(?::(\d{2}))?\s*([AaPp])?\.?[Mm]?\.?$")

# These forms are spreadsheets that compute totals from second-precision stamps but
# print everything rounded to the minute, so time in, time out and the total can each
# be off by up to a minute — 2 minutes of disagreement is honest rounding, not a
# misread. Verified against the source scans: 2-minute gaps are present on the printed
# forms themselves, while every real digit error measured was 3 minutes or more.
TIME_TOLERANCE_MIN = 2


def _parse_hhmm(value: Any) -> int | None:
    """Parse "H:MM"/"HH:MM" into minutes, or None when it isn't a clock value."""
    m = _HHMM_RE.match(str(value or "").strip())
    if not m:
        return None
    hours, minutes = int(m.group(1)), int(m.group(2))
    ampm = (m.group(4) or "").lower()
    if ampm and hours <= 12:
        if hours < 1 or minutes > 59:
            return None
        hours = (hours % 12) + (12 if ampm == "p" else 0)
    # Past 12 the meridiem adds nothing and is often contradictory ("20:00AM"):
    # keep the 24-hour reading and drop the suffix, so the cell is a real time
    # rather than text the client's formulas choke on.
    if hours > 47 or minutes > 59:
        return None
    return hours * 60 + minutes


def _seconds_of(value: Any) -> str:
    """The ":SS" a time was written with, or "" — so we can put it back."""
    m = _HHMM_RE.match(str(value or "").strip())
    return f":{m.group(3)}" if m and m.group(3) else ""


def _fmt_hhmm(total_minutes: int, seconds: str = "") -> str:
    return f"{total_minutes // 60}:{total_minutes % 60:02d}{seconds}"


def _normalise_time_text(value: Any) -> str:
    """Render a clock value as 24-hour HH:MM, leaving anything else untouched.

    The prompt asks for 24-hour, but compliance is not guaranteed — gemini-3.5
    frequently appends "AM"/"PM". Normalising here rather than trusting the model
    means the CSV, the chronology check and the Excel time serial all see one
    format, whatever the model felt like emitting. Non-times (REST DAY, "-", a
    letter code) pass through unchanged.
    """
    text = str(value or "").strip()
    minutes = _parse_hhmm(text)
    if minutes is None:
        return text
    # Seconds are carried through, not rounded away: they are what the machine
    # actually stamped, and the team reconciles against the printed DTR.
    # No % 24 here. A shift that ends at 27:00 is a shift that ended at 3am the
    # next morning, and wrapping it back to "3:00" puts the end of the day before
    # its start — both for the reader and for every duration formula downstream.
    return _fmt_hhmm(minutes, _seconds_of(text))


# Rolling a backwards time forward by 12 hours before trying 24 matters: the two
# fix different mistakes. +12 corrects an AM/PM misread ("12:30 PM" on a night
# shift that ended at 12:30 AM); +24 carries a genuine overnight shift past
# midnight ("3:00" out against an 18:00 in becomes 27:00). Taking the smallest
# roll that moves the row forward picks the right one without having to know
# which mistake it was.
_ROLLS = (12 * 60, 24 * 60, 36 * 60)


def _roll_overnight(values: list[Any]) -> tuple[list[Any], list[int]]:
    """Push times that run backwards onto the next half-day or day.

    Only fires on a value strictly earlier than the one before it, so an ordinary
    day shift is never touched. Blank and non-clock cells pass through and do not
    break the sequence.

    Returns the rolled values and the indices pushed a whole day or more. Those
    are worth surfacing: a +12 roll just fixes a missing AM/PM, but carrying a
    time past midnight is a real inference about when the shift ended, and it is
    only as good as the reading it starts from. One DTR whose 4:36 was misread as
    14:34 turned a 2:00 PM out into "26:00" — self-consistent, and wrong, and it
    passed the chronology check precisely because the roll had made it pass.
    """
    out = list(values)
    crossed: list[int] = []
    previous: int | None = None
    for i, value in enumerate(out):
        current = _parse_hhmm(value)
        if current is None:
            continue
        if previous is not None and current <= previous:
            rolled = next(
                (current + r for r in _ROLLS
                 if current + r > previous and current + r <= 47 * 60 + 59),
                None,
            )
            if rolled is None:
                previous = current
                continue        # can't be reconciled; leave it for the flagger
            if rolled - current >= 24 * 60:
                crossed.append(i)
            current = rolled
            out[i] = _fmt_hhmm(current, _seconds_of(value))
        previous = current
    return out, crossed


def _find_time_groups(headers: list[str]) -> list[dict[str, Any]]:
    """Locate (time in, time out, total) column triples by header name.

    Matches the "<section> - Time in / Time out / TOTAL" shape that timesheets
    use, requiring the three columns to sit next to each other so an unrelated
    grand-total column can't be pulled into a group.
    """
    groups: list[dict[str, Any]] = []
    current: dict[str, Any] | None = None
    for i, header in enumerate(headers):
        text = str(header)
        lowered = text.lower()
        if re.search(r"\b(?:time|break)\s*in\b", lowered):
            label = re.split(r"\s*[-–—]\s*", text)[0].strip() or text.strip()
            current = {"label": label, "in": i}
        elif current is not None and "out" not in current and re.search(
            r"\b(?:time|break)\s*out\b", lowered
        ):
            current["out"] = i
        elif current is not None and "out" in current and "total" in lowered:
            if i - current["out"] <= 2:
                current["total"] = i
                groups.append(current)
            current = None
    return groups


def _time_group_span(row: list[str], group: dict[str, Any]) -> tuple[int, int] | None:
    """Return (computed span, printed total) in minutes, or None if unparseable."""
    values = []
    for key in ("in", "out", "total"):
        col = group[key]
        values.append(_parse_hhmm(row[col]) if col < len(row) else None)
    start, end, printed = values
    if start is None or end is None or printed is None:
        return None
    span = end - start
    if span < 0:
        span += 24 * 60  # shift crossing midnight
    return span, printed


def _failing_time_groups(
    headers: list[str], rows: list[list[str]]
) -> list[dict[str, Any]]:
    """Every in/out/total group whose arithmetic doesn't add up."""
    groups = _find_time_groups(headers)
    if not groups:
        return []
    failures: list[dict[str, Any]] = []
    for row_idx, row in enumerate(rows):
        for group in groups:
            span = _time_group_span(row, group)
            if span and abs(span[0] - span[1]) > TIME_TOLERANCE_MIN:
                failures.append(
                    {
                        "row": row_idx,
                        "group": group,
                        "computed": _fmt_hhmm(span[0]),
                        "printed": _fmt_hhmm(span[1]),
                    }
                )
    return failures


async def _reread_columns(
    b64_image: str,
    mime_type: str,
    column_names: list[str],
    row_labels: list[str],
) -> list[dict[str, Any]]:
    """Re-transcribe specific columns from the image, blind to the first reading.

    Deliberately does NOT show the model its previous values — priming it with a
    misread anchors it to that misread, so it re-reads the columns from scratch.
    """
    instruction = "\n".join(
        [
            "Read this table from the image and transcribe ONLY these columns, exactly as printed:",
            json.dumps(column_names, ensure_ascii=False),
            "",
            "The table's rows, in order, are identified by these labels from its first column:",
            json.dumps(row_labels, ensure_ascii=False),
            "",
            "Read every cell digit by digit. These are printed clock times (H:MM or HH:MM).",
            "Take special care to distinguish 6 from 8, 0 from 8, 1 from 7, 5 from 6, and 9 from 4.",
            'Use "" for cells that are blank.',
            "",
            'Return JSON: {"rows": [{"row_label": "...", "values": {"<column name>": "<value>"}}]}',
            "One entry per row label, in the same order. Output ONLY valid JSON.",
        ]
    )
    payload = {
        "model": await get_model_name(),
        "messages": [
            {
                "role": "user",
                "content": [
                    {
                        "type": "image_url",
                        "image_url": {"url": f"data:{mime_type};base64,{b64_image}"},
                    },
                    {"type": "text", "text": instruction},
                ],
            }
        ],
        "max_tokens": 8192,
        "temperature": 0.0,
        "response_format": {"type": "json_object"},
    }
    async with httpx.AsyncClient(timeout=VERIFY_TIMEOUT) as client:
        resp = await client.post(
            f"{GEMINI_BASE_URL}/chat/completions", json=payload, headers=_gemini_headers()
        )
        resp.raise_for_status()
        raw = resp.json()["choices"][0]["message"]["content"]
    parsed = _extract_json_object(re.sub(r"<think>[\s\S]*?</think>", "", raw))
    if not parsed or not isinstance(parsed.get("rows"), list):
        return []
    return [r for r in parsed["rows"] if isinstance(r, dict)]


# ── DTR (Daily Time Record) support ───────────────────────────────────────
# The printed forms in reference/ carry a per-section TOTAL column, which makes
# them arithmetically self-validating (see _verify_time_math). These handwritten
# DTRs have no TOTAL column at all, so _find_time_groups returns nothing on them
# and that whole safety net sits inert. What they do have is a different free
# invariant: the six stamps in a row run forward through a single day, so a
# reading that goes backwards is provably a misread.
DTR_HEADERS = [
    "DAILY TIME SCHEDULE", "DATE",
    "MORNING IN", "MORNING OUT",
    "AFTERNOON IN", "AFTERNOON OUT",
    "EVENING IN", "EVENING OUT",
    "NO. OF DAYS/HRS", "O.T. HRS", "REASON FOR O.T.",
]
# Canonical slot order, parallel to DTR_HEADERS.
DTR_SLOTS = ["schedule", "date", "m_in", "m_out", "a_in", "a_out",
             "e_in", "e_out", "days_hrs", "ot_hrs", "reason"]
DTR_TIME_SLOTS = ["m_in", "m_out", "a_in", "a_out", "e_in", "e_out"]

# Cropping is the only lever that helps a cell the model keeps misreading. The
# model tokenises an image to a fixed budget, so re-sending the whole page at
# 3000 or 4000px changes nothing (measured — it returns the identical misread),
# while a full-width band crop at 2.5x resolves it. Two overlapping bands cover
# where the data rows sit; full page width is kept so no column is ever cut off
# on a single-form scan with no punch card beside it.
DTR_BANDS = [(0.15, 0.55, 2.5), (0.38, 0.80, 2.5)]
# Narrower slices at higher zoom, tried only for rows the wider bands could not
# settle. Fewer rows in frame and more pixels per digit is exactly the trade that
# fixed the stable misreads, so the escalation tier pushes further along it.
DTR_ESCALATION_BANDS = [(0.15, 0.40, 4.0), (0.33, 0.58, 4.0),
                        (0.50, 0.75, 4.0), (0.66, 0.90, 4.0)]
DTR_BAND_MAX_PIXELS = 8_000_000
# A start this far from the rostered shift start means a misread hour digit.
DTR_SCHEDULE_TOLERANCE_MIN = 180


def _dtr_map_columns(headers: list[str]) -> dict[str, int]:
    """Map arbitrary header spellings onto canonical DTR slots.

    Named sections win, but the fallback is positional: the 1st, 2nd and 3rd
    in/out pairs are morning, afternoon and evening whatever they are called.
    Needed because on scans whose printed header row has degraded the model
    borrows "OVERTIME IN/OUT" from the payroll card beside the grid — the times
    underneath are right, only the label is wrong.
    """
    out: dict[str, int] = {}
    pairs: list[tuple[int, int]] = []
    pending: int | None = None
    for i, header in enumerate(headers):
        text = str(header).lower()
        squashed = re.sub(r"[^a-z]", "", text)
        if "schedule" in squashed:
            out.setdefault("schedule", i)
            continue
        if squashed == "date" or squashed.endswith("date"):
            out.setdefault("date", i)
            continue
        if "reason" in squashed:
            out.setdefault("reason", i)
            continue
        if "days" in squashed and "hrs" in squashed:
            out.setdefault("days_hrs", i)
            continue
        if squashed in {"ot", "othrs", "otohrs"} or ("ot" in squashed and "hrs" in squashed):
            out.setdefault("ot_hrs", i)
            continue
        section = next(
            (s for s in ("morning", "afternoon", "evening") if s in squashed), None
        )
        is_out = "out" in squashed
        is_in = not is_out and (bool(re.search(r"\bin\b", text)) or squashed.endswith("in"))
        if section:
            key = {"morning": "m_", "afternoon": "a_", "evening": "e_"}[section]
            if is_out:
                out.setdefault(key + "out", i)
            elif is_in:
                out.setdefault(key + "in", i)
        if is_in:
            pending = i
        elif is_out and pending is not None:
            pairs.append((pending, i))
            pending = None

    for in_slot, out_slot, pair in zip(
        ("m_in", "a_in", "e_in"), ("m_out", "a_out", "e_out"), pairs
    ):
        out.setdefault(in_slot, pair[0])
        out.setdefault(out_slot, pair[1])
    return out


def _is_dtr_table(headers: list[str]) -> bool:
    """True when these headers describe a DTR-style three-session time grid.

    Deliberately strict. The printed timesheets this tool already handles also
    have three in/out pairs (regular hours, lunch, coffee, overtime) plus a date
    column, so the positional fallback alone would mistake them for a DTR and
    flatten their 14 columns onto these 11 — throwing away the TOTAL columns that
    make them arithmetically checkable. Two guards keep them out: a real DTR
    names its morning and afternoon sessions, and it has no per-section TOTAL.
    """
    mapping = _dtr_map_columns(headers)
    # Morning and afternoon are required; the evening pair is NOT. On a bad read
    # the model sometimes omits the evening columns entirely, and refusing to lock
    # in that case is backwards — the page whose extraction went wrong is the one
    # that most needs to keep the batch's shape. It comes back with the evening
    # cells empty, which is visible, rather than an 8-column page that silently
    # fails to line up with the rest.
    required = ["date", "m_in", "m_out", "a_in", "a_out"]
    if not all(s in mapping for s in required):
        return False
    text = " ".join(str(h).lower() for h in headers)
    if "morning" not in text or "afternoon" not in text:
        return False
    # An in/out/TOTAL triple means _verify_time_math owns this form, not the DTR path.
    return not _find_time_groups(headers)


def _dtr_canonicalise(table_data: dict) -> bool:
    """Rewrite a recognised DTR table into the fixed 11-column schema.

    The same form otherwise comes back with 9, 10 or 11 columns and a different
    header spelling per page (measured: 9 distinct header tuples across 20 runs
    of these 10 scans), because the empty-column stripper drops whichever of
    NO. OF DAYS/HRS or O.T. HRS happens to be blank. Pages of one batch then
    refuse to line up when they are pasted together for payroll.
    """
    headers = table_data.get("headers") or []
    if not _is_dtr_table(headers):
        return False
    mapping = _dtr_map_columns(headers)
    rows = table_data.get("rows") or []
    canonical = [
        [
            _normalise_time_text(row[mapping[slot]])
            if slot in mapping and mapping[slot] < len(row) else ""
            for slot in DTR_SLOTS
        ]
        for row in rows
    ]
    # The six time columns sit in chronological order, so a cell that reads
    # earlier than the one to its left has either crossed midnight or lost its
    # AM/PM. Both are repaired by rolling forward; day shifts are left alone.
    first = DTR_SLOTS.index(DTR_TIME_SLOTS[0])
    last = DTR_SLOTS.index(DTR_TIME_SLOTS[-1]) + 1
    midnight_flags: list[dict[str, Any]] = []
    for row_idx, row in enumerate(canonical):
        row[first:last] = _clear_smeared_word(row[first:last])
        rolled, crossed = _roll_overnight(row[first:last])
        row[first:last] = rolled
        for offset in crossed:
            midnight_flags.append({
                "row": row_idx, "col": first + offset,
                "reason": "read as running past midnight — confirm this against "
                          "the form, and check the time-in on this row",
            })
    midnight_flags += _half_day_flags(
        canonical,
        [DTR_SLOTS.index(k) for k in ("m_in", "a_in", "e_in")],
        [DTR_SLOTS.index(k) for k in ("m_out", "a_out", "e_out")])
    if midnight_flags:
        table_data["flagged_cells"] = (
            table_data.get("flagged_cells") or []) + midnight_flags
    table_data["rows"] = canonical
    table_data["headers"] = list(DTR_HEADERS)
    return True


def _dtr_chronology_ok(values: list[Any]) -> bool:
    """True when these stamps can be one day's forward sequence.

    Tolerant of 12-hour input: a value before 12:00 may be the afternoon of the
    same day, so each cell is allowed to read as either itself or itself plus 12
    hours, and the earliest reading that still moves forward is taken. Without
    that, a page transcribed "1:00, 4:00, 4:15, 5:30" — correct, just not
    normalised — would light up every row as an error.

    Blank cells are skipped rather than treated as breaks: staff routinely leave
    a whole session empty on a short day.
    """
    previous: int | None = None
    for value in values:
        current = _parse_hhmm(value)
        if current is None:
            continue
        options = [current] + ([current + 12 * 60] if current < 12 * 60 else [])
        nxt = next((o for o in options if previous is None or o > previous), None)
        if nxt is None:
            return False
        previous = nxt
    return True


def _dtr_schedule_window(schedule: Any) -> tuple[int, int] | None:
    """(start, end) minutes from a "12:15-9:15" / "10:00 to 7:00" schedule cell.

    The end is read as an evening time: these are retail shifts that finish in
    the afternoon or evening, so a bare "5:30" end means 17:30.
    """
    found = re.findall(r"(\d{1,2}):(\d{2})", str(schedule or ""))
    if len(found) < 2:
        return None
    start = (int(found[0][0]) % 24) * 60 + int(found[0][1])
    end_h, end_m = int(found[1][0]) % 12, int(found[1][1])
    return start, (end_h + 12) * 60 + end_m


def _dtr_schedule_disagrees(values: list[Any], schedule: Any) -> bool:
    """True when the row's first stamp is nowhere near the rostered start.

    Kept separate from the chronology check on purpose: a schedule cell clipped
    by the page edge is misread while all six times are perfect, and re-reading
    the times in that case is wasted work.
    """
    window = _dtr_schedule_window(schedule)
    first = next((_parse_hhmm(v) for v in values if _parse_hhmm(v) is not None), None)
    if not window or first is None:
        return False
    return abs(first - window[0]) > DTR_SCHEDULE_TOLERANCE_MIN


def _dtr_accept_reread(before: list[str], after: list[str]) -> list[int] | None:
    """Indices to change, or None when the re-read isn't a credible repair.

    Guards against a re-read that latched onto the wrong row and would swap in a
    wrong-but-ordered set of times: every change has to look like a misread digit
    or two in a value of the same length, and the result has to fix the ordering.
    """
    changed = [
        i for i in range(len(before))
        if str(before[i]).strip() != str(after[i]).strip()
    ]
    if not changed:
        return None
    total = 0
    for i in changed:
        a, b = str(before[i]).strip(), str(after[i]).strip()
        if not a or not b or len(a) != len(b):
            return None
        distance = _edit_distance(a, b)
        if distance > MAX_CELL_EDITS:
            return None
        total += distance
    if total > MAX_GROUP_EDITS or not _dtr_chronology_ok(after):
        return None
    return changed


def _source_image(
    raw_bytes: bytes, content_type: str | None, filename: str | None,
    page_index: int = 0
) -> Image.Image:
    """The upload as an upright RGB image, for cropping bands out of.

    Cropped from the source rather than the model copy: the model copy has already
    been upscaled, so magnifying it again would only interpolate what is there.
    """
    if _is_pdf_upload(content_type, filename):
        pdf = pdfium.PdfDocument(io.BytesIO(raw_bytes))
        try:
            page = pdf[min(max(0, page_index), len(pdf) - 1)]
            try:
                return page.render(scale=3).to_pil().convert("RGB")
            finally:
                page.close()
        finally:
            pdf.close()
    return ImageOps.exif_transpose(Image.open(io.BytesIO(raw_bytes))).convert("RGB")


def _dtr_band_image(img: Image.Image, band: tuple[float, float, float]) -> Image.Image:
    """Crop a full-width horizontal band and magnify it, within a pixel budget."""
    y0, y1, zoom = band
    w, h = img.size
    crop = img.crop((0, int(y0 * h), w, int(y1 * h)))
    cw, ch = int(crop.width * zoom), int(crop.height * zoom)
    if cw * ch > DTR_BAND_MAX_PIXELS:
        shrink = (DTR_BAND_MAX_PIXELS / float(cw * ch)) ** 0.5
        cw, ch = int(cw * shrink), int(ch * shrink)
    resampling = getattr(getattr(Image, "Resampling", Image), "LANCZOS")
    return crop.resize((max(cw, 1), max(ch, 1)), resampling)


# ── Break-detail timesheet support ────────────────────────────────────────
# The other client's forms are a different shape again: a date column plus
# REGULAR HOURS / LUNCH BREAK / COFFEE BREAK / OVERTIME, each with its own
# in / out / TOTAL triple. Measured across 20 of them the printed layout itself
# varies (12, 13, 14, 15 and 17 columns — some drop DUTY LOCATION, some drop the
# lunch in/out and keep only its total), so the extractor was returning a
# different schema per page even before anything went wrong. This canonical
# superset is the union of every variant seen.
TIMESHEET_HEADERS = [
    "DATE", "PLOTTED SCHEDULE", "DUTY LOCATION",
    "REGULAR HOURS - Time in", "REGULAR HOURS - Time out", "REGULAR HOURS - TOTAL",
    "LUNCH BREAK - Break out", "LUNCH BREAK - Break in", "LUNCH BREAK - TOTAL",
    "COFFEE BREAK - Break out", "COFFEE BREAK - Break in", "COFFEE BREAK - TOTAL",
    "OVERTIME - Time in", "OVERTIME - Time out", "OVERTIME - TOTAL",
    "TARDINESS", "REASON/REMARKS",
]
TIMESHEET_SLOTS = [
    "date", "schedule", "location",
    "reg_in", "reg_out", "reg_total",
    "lunch_out", "lunch_in", "lunch_total",
    "coffee_out", "coffee_in", "coffee_total",
    "ot_in", "ot_out", "ot_total",
    "tardiness", "reason",
]
_TS_SECTIONS = {"regular": "reg", "lunch": "lunch", "coffee": "coffee",
                "overtime": "ot", "o.t": "ot"}


def _timesheet_map_columns(headers: list[str]) -> dict[str, int]:
    """Map arbitrary header spellings onto the canonical timesheet slots."""
    out: dict[str, int] = {}
    for i, header in enumerate(headers):
        text = str(header).lower()
        squashed = re.sub(r"[^a-z]", "", text)
        # Checked before the section scan: "TARDINESS" contains "in", and
        # "DUTY LOCATION" would otherwise fight the location/section tests.
        if "tardiness" in squashed or squashed == "late":
            out.setdefault("tardiness", i)
            continue
        if "reason" in squashed or "remark" in squashed:
            out.setdefault("reason", i)
            continue
        if "location" in squashed:
            out.setdefault("location", i)
            continue
        if "schedule" in squashed:
            out.setdefault("schedule", i)
            continue
        if squashed == "date" or squashed.endswith("date"):
            out.setdefault("date", i)
            continue
        prefix = next((p for key, p in _TS_SECTIONS.items() if key in text), None)
        if not prefix:
            continue
        if "total" in squashed:
            out.setdefault(f"{prefix}_total", i)
        elif "out" in squashed:
            out.setdefault(f"{prefix}_out", i)
        elif re.search(r"\bin\b", text) or squashed.endswith("in"):
            out.setdefault(f"{prefix}_in", i)

    if "reg_in" not in out:
        _timesheet_positional_fallback(headers, out)
    return out


def _timesheet_positional_fallback(headers: list[str], out: dict[str, int]) -> None:
    """Assign unprefixed in/out/TOTAL triples to sections by position.

    The model does not always keep the section prefixes — one run returned
    "Date, Time In, Time out, TOTAL, BREAK OUT, BREAK IN, ..." with no "REGULAR
    HOURS" anywhere, and the named pass then found nothing. The section order on
    these forms is fixed, so the 1st triple is regular hours, the 2nd lunch, the
    3rd coffee and the 4th overtime, whatever they happen to be called.

    Only engages when at least three triples are present. A lone in/out/TOTAL is
    far too weak a signal to claim a document on.
    """
    kinds: list[tuple[int, str]] = []
    for i, header in enumerate(headers):
        text = str(header).lower()
        squashed = re.sub(r"[^a-z]", "", text)
        if "tardiness" in squashed or "reason" in squashed or "remark" in squashed:
            continue
        if "total" in squashed:
            kinds.append((i, "total"))
        elif "out" in squashed:
            kinds.append((i, "out"))
        elif re.search(r"\bin\b", text) or squashed.endswith("in"):
            kinds.append((i, "in"))

    groups: list[list[tuple[int, str]]] = []
    current: list[tuple[int, str]] = []
    for i, kind in kinds:
        current.append((i, kind))
        if kind == "total":          # a TOTAL closes its section
            groups.append(current)
            current = []
    if current:
        groups.append(current)

    if len(groups) < 3:
        return
    for prefix, group in zip(("reg", "lunch", "coffee", "ot"), groups):
        for i, kind in group:
            out.setdefault(f"{prefix}_{kind}", i)


def _is_timesheet_table(headers: list[str]) -> bool:
    """True for the break-detail timesheet shape.

    Keyed on a REGULAR HOURS in/out pair, which is distinctive enough on its own.
    A date column is deliberately NOT required: on a bad read the model sometimes
    drops the date column altogether, and refusing to lock the schema in that case
    is precisely backwards — the page whose extraction went wrong is the one that
    most needs to keep the same shape as the rest of the batch. It comes back with
    an empty DATE instead, which is visible, rather than a 13-column page that
    silently fails to line up.

    Defers to the DTR path when that matches, so the two can never both claim a form.
    """
    if _is_dtr_table(headers):
        return False
    mapping = _timesheet_map_columns(headers)
    return "reg_in" in mapping and "reg_out" in mapping


# Each (start, end) session of a break-detail timesheet, in slot names.
_TIMESHEET_PAIRS = (("reg_in", "reg_out"), ("lunch_out", "lunch_in"),
                    ("coffee_out", "coffee_in"), ("ot_in", "ot_out"))


def _timesheet_canonicalise(table_data: dict) -> bool:
    """Rewrite a recognised timesheet into the fixed 17-column schema."""
    headers = table_data.get("headers") or []
    if not _is_timesheet_table(headers):
        return False
    mapping = _timesheet_map_columns(headers)
    rows = table_data.get("rows") or []
    canonical = [
        [
            _normalise_time_text(row[mapping[slot]])
            if slot in mapping and mapping[slot] < len(row) else ""
            for slot in TIMESHEET_SLOTS
        ]
        for row in rows
    ]
    # Roll each session's end past its start, pair by pair and never the TOTAL
    # columns (those are durations). The client's own Alberto sheets carry
    # time-outs in 12-hour form — in 8:54, out 5:44 — and the workbook computes
    # out - in with no wrap, so those came out negative. The time-math check
    # already reads a backwards out as the next day; this makes the cells agree.
    midnight_flags: list[dict[str, Any]] = []
    for row_idx, row in enumerate(canonical):
        for start, end in _TIMESHEET_PAIRS:
            a, b = TIMESHEET_SLOTS.index(start), TIMESHEET_SLOTS.index(end)
            (row[a], row[b]), crossed = _roll_overnight([row[a], row[b]])
            if crossed:
                midnight_flags.append({
                    "row": row_idx, "col": b,
                    "reason": "read as running past midnight — confirm this against "
                              "the form, and check the time-in on this row",
                })
    midnight_flags += _half_day_flags(
        canonical, [TIMESHEET_SLOTS.index("reg_in")], [TIMESHEET_SLOTS.index("reg_out")])
    if midnight_flags:
        table_data["flagged_cells"] = (
            table_data.get("flagged_cells") or []) + midnight_flags
    table_data["rows"] = canonical
    table_data["headers"] = list(TIMESHEET_HEADERS)
    return True


# ── Simple in/out payroll sheet (24 Chicken) ──────────────────────────────
# One punch pair per day plus the client's own computed columns. Distinct from
# both the DTR and the break-detail timesheet: there is no second or third
# session, so the only structural signal is a single in/out pair sitting beside
# a break/net-total column set.
# Holiday columns are in the canonical set even though the workbook template does
# not read them: some of these sheets print them, and dropping a column that the
# client filled in loses data that only exists on the paper.
# SCHEDULE and REMARKS: 24 Chicken's sheets print the day's shift ("12-10", "RD")
# between DAY and TIME IN, and the checker writes WDO / SPL / OT beside a row.
# Both feed the timekeeping engine (schedule, Remarks tag); sheets without them
# just leave the columns blank.
INOUT_HEADERS = [
    "DATE", "DAY", "SCHEDULE", "TIME IN", "TIME OUT", "BREAK (HR)", "NORMAL", "OT HRS",
    "HOLIDAY PAY", "HOLIDAY+OT", "TOTAL", "DEDUCTIONS", "NET TOTAL", "REMARKS",
]
INOUT_SLOTS = [
    "date", "day", "schedule", "time_in", "time_out", "break_hrs", "normal", "ot_hrs",
    "holiday_pay", "holiday_ot", "total", "deductions", "net_total", "remarks",
]


def _inout_map_columns(headers: list[str]) -> dict[str, int]:
    out: dict[str, int] = {}
    for i, header in enumerate(headers):
        text = str(header).lower()
        squashed = re.sub(r"[^a-z0-9]", "", text)
        if "sched" in squashed or squashed == "shift":
            out.setdefault("schedule", i)
        elif "remark" in squashed or squashed in ("notes", "note", "marks"):
            out.setdefault("remarks", i)
        elif squashed == "date" or squashed.endswith("date"):
            out.setdefault("date", i)
        elif squashed == "day":
            out.setdefault("day", i)
        elif "holiday" in squashed and ("ot" in squashed or "+" in text):
            out.setdefault("holiday_ot", i)
        elif "holiday" in squashed:
            out.setdefault("holiday_pay", i)
        elif "nettotal" in squashed:
            out.setdefault("net_total", i)
        elif "deduction" in squashed:
            out.setdefault("deductions", i)
        elif "break" in squashed:
            out.setdefault("break_hrs", i)
        elif "normal" in squashed:
            out.setdefault("normal", i)
        elif squashed.startswith("ot") or "othrs" in squashed or "overtime" in squashed:
            out.setdefault("ot_hrs", i)
        elif "total" in squashed:
            out.setdefault("total", i)
        elif "out" in squashed:
            out.setdefault("time_out", i)
        elif re.search(r"\bin\b", text) or squashed.endswith("in"):
            out.setdefault("time_in", i)
    return out


def _is_inout_table(headers: list[str]) -> bool:
    """True for the single in/out payroll sheet, and only that.

    Defers to both richer layouts first: a DTR or a break-detail timesheet also
    contains a time in and a time out, so without those guards this would claim
    either of them and flatten away their extra sessions.
    """
    if _is_dtr_table(headers) or _is_timesheet_table(headers):
        return False
    mapping = _inout_map_columns(headers)
    if not all(k in mapping for k in ("date", "time_in", "time_out")):
        return False
    # Needs at least one of the client's computed columns, so an arbitrary two
    # column in/out table isn't mistaken for this form.
    return any(k in mapping for k in ("break_hrs", "net_total", "normal", "deductions"))


def _inout_unshift(rows: list[list[str]]) -> list[dict[str, Any]]:
    """Pull back a row read one column to the right.

    A handwritten mark at the start of a row (an "SPL" bracket beside the
    schedule) can make the model emit an extra blank cell: TIME IN comes back
    empty, the time in lands in TIME OUT and the time out in BREAK (Oppa Foods,
    2 Nov). BREAK holds hours, never a clock time, so that pattern is proof.
    """
    t_in, t_out = INOUT_SLOTS.index("time_in"), INOUT_SLOTS.index("time_out")
    brk, rem = INOUT_SLOTS.index("break_hrs"), INOUT_SLOTS.index("remarks")
    def clock(v) -> int | None:
        m = re.fullmatch(r"\s*(\d{1,2}):(\d{2})\s*", str(v or ""))
        return int(m.group(1)) * 60 + int(m.group(2)) if m else None

    flags = []
    for i, row in enumerate(rows):
        start, end = clock(row[t_out]), clock(row[brk])
        # Only when the pair reads as a day's in and out ("9:00", "19:00"), so a
        # form that prints BREAK as "1:00" is never mistaken for a shifted row.
        if not str(row[t_in]).strip() and start is not None and end is not None and start < end:
            row[t_in:rem] = row[t_in + 1:rem] + [""]
            flags.append({"row": i, "col": t_in,
                          "reason": "this row was read one column to the right and has been "
                                    "moved back; check the time in and time out"})
    return flags


def _inout_canonicalise(table_data: dict) -> bool:
    """Rewrite a recognised in/out sheet into the fixed in/out schema."""
    headers = table_data.get("headers") or []
    if not _is_inout_table(headers):
        return False
    mapping = _inout_map_columns(headers)
    rows = table_data.get("rows") or []
    canonical = [
        [
            _normalise_time_text(row[mapping[slot]])
            if slot in mapping and mapping[slot] < len(row) else ""
            for slot in INOUT_SLOTS
        ]
        for row in rows
    ]
    inout_flags: list[dict[str, Any]] = _inout_unshift(canonical)
    # time_in/time_out are adjacent and in order; a night shift's out belongs on
    # the next day. Everything after them is hours and money, so stop there.
    lo = INOUT_SLOTS.index("time_in")
    for row_idx, row in enumerate(canonical):
        rolled, crossed = _roll_overnight(row[lo:lo + 2])
        row[lo:lo + 2] = rolled
        for offset in crossed:
            inout_flags.append({
                "row": row_idx, "col": lo + offset,
                "reason": "read as running past midnight — confirm this against "
                          "the form, and check the time-in on this row",
            })
    inout_flags += _half_day_flags(
        canonical, [INOUT_SLOTS.index("time_in")], [INOUT_SLOTS.index("time_out")])
    if inout_flags:
        table_data["flagged_cells"] = (
            table_data.get("flagged_cells") or []) + inout_flags
    table_data["rows"] = canonical
    table_data["headers"] = list(INOUT_HEADERS)
    return True


# ── Missing-row detection ─────────────────────────────────────────────────
# A timesheet covers a contiguous run of days, and prints the range it covers.
# That makes a dropped row provable rather than merely suspicious: either a day is
# missing from the middle of the sequence, or the sequence stops short of the
# printed period. Silently losing a day costs someone a day's pay.
_MONTHS = {
    m: i + 1 for i, m in enumerate(
        ["january", "february", "march", "april", "may", "june", "july",
         "august", "september", "october", "november", "december"])
}
_ROW_DATE_RE = re.compile(
    r"([A-Za-z]+)\s+(\d{1,2})\s*,?\s*(\d{4})")
_PERIOD_RE = re.compile(
    r"([A-Za-z]+)\s+(\d{1,2})\s*[-–—]\s*(\d{1,2})\s*,?\s*(\d{4})")


def _month_num(name: str) -> int | None:
    key = name.strip().lower()
    for full, num in _MONTHS.items():
        if full.startswith(key[:3]):
            return num
    return None


def _parse_row_date(text: Any) -> "date | None":
    m = _ROW_DATE_RE.search(str(text or ""))
    if not m:
        return None
    month = _month_num(m.group(1))
    if not month:
        return None
    try:
        return date(int(m.group(3)), month, int(m.group(2)))
    except ValueError:
        return None


def _parse_period(lines: list[str]) -> "tuple[date, date] | None":
    """Read the printed 'Period Covered: JULY 16-31 2026' range, if present."""
    for line in lines or []:
        if "period" not in str(line).lower():
            continue
        m = _PERIOD_RE.search(str(line))
        if not m:
            continue
        month = _month_num(m.group(1))
        if not month:
            continue
        try:
            year = int(m.group(4))
            return (date(year, month, int(m.group(2))),
                    date(year, month, int(m.group(3))))
        except ValueError:
            continue
    return None


def _find_missing_rows(table_data: dict) -> list[str]:
    """Days the table should contain but doesn't, as human-readable dates."""
    rows = table_data.get("rows") or []
    dates = [_parse_row_date(r[0]) for r in rows if r]
    present = sorted({d for d in dates if d})
    # Only meaningful for a date-indexed table where most rows parsed.
    if len(present) < 3 or len(present) < len(rows) * 0.6:
        return []

    expected_start, expected_end = present[0], present[-1]
    period = _parse_period(table_data.get("ocr_text_lines") or [])
    if period:
        # Trust the printed period only when it plausibly matches these rows,
        # so a misread period can't invent dozens of missing days.
        p_start, p_end = period
        if 0 <= (present[0] - p_start).days <= 3 and 0 <= (p_end - present[-1]).days <= 3:
            expected_start, expected_end = p_start, p_end

    missing = []
    day = expected_start
    while day <= expected_end:
        if day not in present:
            missing.append(day.strftime("%A, %B %-d, %Y"))
        day += timedelta(days=1)
    return missing


# Words that only ever appear in a column heading. A form with stacked headings —
# "REGULAR HOURS" over "Time in / Time out / TOTAL" — often comes back with the
# second heading row transcribed as the first data row, which then lands in the
# workbook as the literal text "Time Out" where a time belongs.
_HEADER_WORD = re.compile(
    r"^(?:time\s*)?(?:in|out)$|^total$|^date$|^day$|^a\.?m\.?$|^p\.?m\.?$"
    r"|^morning$|^afternoon$|^evening$|^break(?:\s*(?:in|out))?$|^remarks?$"
    r"|^reason$|^schedule$|^hrs?$|^hours$|^o\.?t\.?$|^no\.?$|^duty\s*location$"
    r"|^tardiness$|^undertime$|^time$",
    re.I,
)


def _looks_like_header_row(cells: list[Any]) -> bool:
    """True for a row that is a repeat of the column headings, not a day's data.

    Deliberately strict: a single clock value or date anywhere in the row is
    enough to rule it out, so a real day is never dropped on a form that happens
    to label a column "OUT".
    """
    values = [str(c).strip() for c in cells if str(c).strip()]
    if len(values) < 2:
        return False
    if any(_parse_hhmm(v) is not None for v in values):
        return False
    if any(_parse_row_date(v) is not None for v in values):
        return False
    hits = sum(1 for v in values if _HEADER_WORD.match(v))
    return hits >= 2 and hits * 2 >= len(values)


# Which canonical order each client layout is re-mapped onto before it is written
# into the workbook by position. dtr_ot shares the DTR order — its evening pair is
# the OT pair — so both use the same mapper.
_LAYOUT_MAPPERS = {
    "dtr": (_dtr_map_columns, DTR_SLOTS, ["date", *DTR_TIME_SLOTS]),
    "dtr_ot": (_dtr_map_columns, DTR_SLOTS, ["date", *DTR_TIME_SLOTS]),
    "timesheet": (_timesheet_map_columns, TIMESHEET_SLOTS, ["date", "reg_in", "reg_out"]),
    "inout": (_inout_map_columns, INOUT_SLOTS, ["date", "time_in", "time_out"]),
}


# Some client workbooks carry columns that no input layout supplies, because they
# summarise the row rather than transcribe a box on the form. Nothing was filling
# them: SM Asinan's and Trinity's FINAL OUT, Bikezilla's and Upson's too, were
# written blank on every row while the individual CSV showed the times correctly —
# which is exactly the mismatch the client reported.
_OUT_SLOTS_LEFT_TO_RIGHT = ("s1_out", "s2_out", "s3_out", "ot_out")
_IN_SLOTS_LEFT_TO_RIGHT = ("s1_in", "s2_in", "s3_in", "ot_in")

# Row dates arrive either as prose ("Thursday, July 16, 2026") or numerically
# ("06/13/2026"), depending on the form. _parse_row_date only reads the prose
# form, and it is used for row indexing where widening the match would change
# behaviour, so the numeric case is handled separately here.
_NUMERIC_DATE_RE = re.compile(r"\b(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})\b")


def _row_weekday(value: Any) -> str:
    """Abbreviated weekday for a row date, or "" when it cannot be read."""
    parsed = _parse_row_date(value)
    if parsed is None:
        m = _NUMERIC_DATE_RE.search(str(value or ""))
        if not m:
            return ""
        month, day, year = (int(g) for g in m.groups())
        if year < 100:
            year += 2000
        try:
            parsed = date(year, month, day)
        except ValueError:
            return ""
    return parsed.strftime("%a")


# A word written across the row — "REST DAY", "SUNDAY", "NO WORK" — comes back
# with one letter per column, because the model reads the grid cell by cell. Those
# letters land in the workbook's time columns as text, and the client's formulas
# return #VALUE! on every one. Require several of them before blanking, so a form
# that legitimately puts a single code in one column is left alone.
# Digits included: the model routinely reads the O in "NO WORK" as a zero, and a
# bare "0" in a time column is no more a time than a bare "W" is.
_SMEARED_LETTER = re.compile(r"^[A-Za-z0-9]\s?[A-Za-z0-9]?$")


def _clear_smeared_word(values: list[Any]) -> list[Any]:
    """Blank a word that has been split one letter per time column."""
    hits = [i for i, v in enumerate(values) if _SMEARED_LETTER.match(str(v or "").strip())]
    if len(hits) < 3:
        return values
    out = list(values)
    for i in hits:
        out[i] = ""
    return out


# The clock columns of a row, in the order they are printed, across every layout.
# Used to sweep a row's time cells as a group.
_ROW_TIME_SLOTS = ("s1_in", "s1_out", "s2_in", "s2_out", "s3_in", "s3_out",
                   "ot_in", "ot_out", "time_in", "time_out")


_HALF_DAY_REASON = ("a time-in with no time-out (or the reverse) — this day counts "
                    "no hours until the missing punch is filled in")


def _half_day_flags(rows: list[list[Any]], ins: list[int], outs: list[int]) -> list[dict[str, Any]]:
    """Flag rows with clock-ins but no clock-out, or the reverse."""
    flags: list[dict[str, Any]] = []
    for row_idx, row in enumerate(rows):
        got_in = [c for c in ins if c < len(row) and _parse_hhmm(row[c]) is not None]
        got_out = [c for c in outs if c < len(row) and _parse_hhmm(row[c]) is not None]
        if bool(got_in) != bool(got_out):
            for col in got_in or got_out:
                flags.append({"row": row_idx, "col": col, "reason": _HALF_DAY_REASON})
    return flags


def _derive_row_slots(values: dict) -> dict:
    """Fill template columns the layout has no column of its own for.

    Several workbooks give the day a single IN and a single OUT box, and feed a
    FINAL OUT formula off the OUT one. Writing the *afternoon* out into it leaves
    it blank whenever the shift ran into the evening, which is why those columns
    came out empty in the workbook while the individual CSV showed the times.
    The day's first in and last out are what those columns actually mean.
    """
    # Second line of defence, matching the header-row strip: /api/ocr clears a
    # smeared word during canonicalisation, but the batch held in the browser is
    # what actually reaches the workbook, and can be re-sent after edits. Letters
    # left in a clock column are written as text and the client's formulas return
    # #VALUE! on every row built from them.
    present = [k for k in _ROW_TIME_SLOTS if k in values]
    if present:
        swept = _clear_smeared_word([values[k] for k in present])
        for k, v in zip(present, swept):
            values[k] = v

    # A day with no clock punches but a word in one of its time boxes (REST DAY,
    # ABSENT, OFF) is a day marker. The columns the TOTAL formula reads take only
    # times now, so the word would otherwise vanish from the sheet; the writer
    # puts it back somewhere no formula reads it.
    if not str(values.get("day_note") or "").strip():
        clocks = [k for k in present if _parse_hhmm(values.get(k)) is not None]
        words = [str(values.get(k)).strip() for k in present
                 if re.search(r"[A-Za-z]{2}", str(values.get(k) or ""))
                 and _parse_hhmm(values.get(k)) is None]
        values["day_note"] = words[0] if (words and not clocks) else ""

    derived_in = not str(values.get("first_in") or "").strip()
    derived_out = not str(values.get("final_out") or "").strip()

    if not str(values.get("final_out") or "").strip():
        last = ""
        for slot in _OUT_SLOTS_LEFT_TO_RIGHT:
            text = str(values.get(slot) or "").strip()
            if text and _parse_hhmm(text) is not None:
                last = text
        values["final_out"] = last
    if not str(values.get("first_in") or "").strip():
        first = ""
        for slot in _IN_SLOTS_LEFT_TO_RIGHT:
            text = str(values.get(slot) or "").strip()
            if text and _parse_hhmm(text) is not None:
                first = text
                break
        values["first_in"] = first
    # A day's hours need both ends. TOTAL is IF(AND(IN<>"",OUT<>""),OUT-IN,0), but
    # where OUT is itself a formula (FINAL OUT) it reads 0 rather than blank when
    # nothing was written — so a time-in whose time-out the model missed, or wrote
    # as "20:5PM", came out as 0 - IN: -11:00. Feed the pair whole or not at all;
    # the review screen flags the half-read day so the missing punch gets filled.
    if derived_in and derived_out and bool(values["first_in"]) != bool(values["final_out"]):
        values["first_in"] = values["final_out"] = ""
    if not str(values.get("day") or "").strip():
        values["day"] = _row_weekday(values.get("date"))
    return values


def _strip_leading_header_rows(table_data: dict) -> int:
    """Drop heading rows that were transcribed as data at the top of the table."""
    rows = table_data.get("rows") or []
    dropped = 0
    while rows and _looks_like_header_row(rows[0]):
        rows.pop(0)
        dropped += 1
    if dropped:
        table_data["rows"] = rows
        logger.info("Dropped %d repeated header row(s) from the top of the table", dropped)
    return dropped


def _strip_trailing_summary_row(table_data: dict) -> bool:
    """Drop a trailing 'TOTAL:' row from a date-indexed table.

    The form's grand-total line is a summary, not a day. It gets picked up only
    sometimes, which makes the row count wobble between 16 and 17 and knocks the
    rows out of alignment when pasted into a fixed template.
    """
    rows = table_data.get("rows") or []
    if len(rows) < 3:
        return False
    dated = sum(1 for r in rows[:-1] if r and _parse_row_date(r[0]))
    if dated < (len(rows) - 1) * 0.8:
        return False  # not a date-indexed table; leave it alone
    last = rows[-1]
    if not last or _parse_row_date(last[0]):
        return False
    if not re.match(r"^\s*total\b", str(last[0] or ""), re.I):
        return False
    rows.pop()
    logger.info("Dropped trailing summary row from date-indexed table")
    return True


def _norm_hhmm_text(value: Any) -> str:
    """Canonical H:MM text so "09:43" and "9:43" compare as the same reading."""
    parsed = _parse_hhmm(value)
    return _fmt_hhmm(parsed) if parsed is not None else str(value or "").strip()


def _edit_distance(a: str, b: str) -> int:
    """Levenshtein distance — small for a digit misread, large for a different value."""
    if a == b:
        return 0
    prev = list(range(len(b) + 1))
    for i, ca in enumerate(a, 1):
        cur = [i]
        for j, cb in enumerate(b, 1):
            cur.append(min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + (ca != cb)))
        prev = cur
    return prev[-1]


# A re-read may resolve the arithmetic while actually having latched onto a
# different cell or row, which would swap in a wrong-but-self-consistent triple.
# Only accept changes that look like what they claim to be: a misread digit or two
# in a value of the same length.
MAX_CELL_EDITS = 2
MAX_GROUP_EDITS = 3


def _is_plausible_reread(before: Any, after: Any) -> bool:
    """True when `after` reads like a digit-level correction of `before`."""
    a, b = _norm_hhmm_text(before), _norm_hhmm_text(after)
    if a == b:
        return True
    if not a or not b or len(a) != len(b):
        return False
    return _edit_distance(a, b) <= MAX_CELL_EDITS


async def _verify_time_math(
    table_data: dict, b64_image: str, mime_type: str
) -> None:
    """Detect impossible in/out/total rows, then repair them via a blind re-read.

    A correction is only accepted when the independently re-read values make the
    group's arithmetic work out — otherwise the cells are flagged for the user to
    check rather than overwritten with another guess.
    """
    headers: list[str] = table_data.get("headers") or []
    rows: list[list[str]] = table_data.get("rows") or []
    failures = _failing_time_groups(headers, rows)
    table_data["time_checked"] = bool(_find_time_groups(headers))
    if not failures:
        return

    logger.info("Time-math check: %d inconsistent group(s), re-reading", len(failures))
    col_idxs = sorted(
        {failure["group"][key] for failure in failures for key in ("in", "out", "total")}
    )
    column_names = [headers[i] for i in col_idxs if i < len(headers)]
    row_labels = [str(row[0]) if row else "" for row in rows]

    try:
        fresh = await _reread_columns(b64_image, mime_type, column_names, row_labels)
    except Exception as e:
        logger.warning("Time-math re-read failed (%s): %s", type(e).__name__, e)
        fresh = []

    def fresh_value(row_idx: int, column: str) -> str | None:
        if row_idx >= len(fresh):
            return None
        values = fresh[row_idx].get("values")
        if not isinstance(values, dict):
            return None
        value = values.get(column)
        if value is None:
            # tolerate minor header rewording from the model
            target = re.sub(r"[^a-z0-9]", "", column.lower())
            for k, v in values.items():
                if re.sub(r"[^a-z0-9]", "", str(k).lower()) == target:
                    value = v
                    break
        text = str(value).strip() if value is not None else ""
        return text or None

    corrections: list[dict[str, Any]] = []
    flagged: list[dict[str, Any]] = []

    for failure in failures:
        row_idx, group = failure["row"], failure["group"]
        candidate = list(rows[row_idx])
        for key in ("in", "out", "total"):
            col = group[key]
            if col >= len(headers):
                continue
            value = fresh_value(row_idx, headers[col])
            if value and _parse_hhmm(value) is not None and col < len(candidate):
                candidate[col] = value

        changed = [
            group[key]
            for key in ("in", "out", "total")
            if group[key] < len(candidate) and candidate[group[key]] != rows[row_idx][group[key]]
        ]
        edits = sum(
            _edit_distance(_norm_hhmm_text(rows[row_idx][col]), _norm_hhmm_text(candidate[col]))
            for col in changed
        )
        plausible = edits <= MAX_GROUP_EDITS and all(
            _is_plausible_reread(rows[row_idx][col], candidate[col]) for col in changed
        )

        span = _time_group_span(candidate, group)
        if span and abs(span[0] - span[1]) <= TIME_TOLERANCE_MIN and plausible:
            for col in changed:
                corrections.append(
                    {
                        "row": row_idx,
                        "col": col,
                        "column": headers[col] if col < len(headers) else "",
                        "from": rows[row_idx][col],
                        "to": candidate[col],
                    }
                )
            rows[row_idx] = candidate
        else:
            reason = (
                f"{group['label']}: printed total {failure['printed']} but "
                f"time out - time in = {failure['computed']}"
            )
            for key in ("in", "out", "total"):
                flagged.append({"row": row_idx, "col": group[key], "reason": reason})

    if corrections:
        logger.info("Time-math check: auto-corrected %d cell(s)", len(corrections))
    if flagged:
        logger.info("Time-math check: flagged %d cell(s) for review", len(flagged))
    table_data["time_corrections"] = corrections
    # Append: canonicalisation may already have flagged a midnight crossing.
    table_data["flagged_cells"] = (table_data.get("flagged_cells") or []) + flagged


_DTR_REREAD_PROMPT = """This is a magnified horizontal band of a Daily Time Record grid.

Find the row whose DATE cell reads exactly one of these, and only these:
%ROWS%

For each one you can find, read its six clock-time cells in this left-to-right
order, digit by digit:
  MORNING IN, MORNING OUT, AFTERNOON IN, AFTERNOON OUT, EVENING IN, EVENING OUT

Take special care with 6 vs 8, 0 vs 8, 1 vs 7, 2 vs 7, 5 vs 6, 9 vs 4.
Write each time as 24-hour HH:MM: the afternoon and evening are written in
12-hour form on the paper, so "2:46" in AFTERNOON OUT is 14:46.
Use "" for a cell that is blank.
If a date above is not visible in this band, omit it entirely -- do not guess.

Return JSON:
{"rows": [{"date": "...", "values": ["HH:MM","HH:MM","HH:MM","HH:MM","HH:MM","HH:MM"]}]}
Output ONLY valid JSON.
"""


async def _reread_dtr_band(img: Image.Image, band, dates: list[str],
                           model: str | None = None) -> dict[str, list[str]]:
    """Re-read the six time cells for `dates` from one magnified band crop.

    Deliberately does not show the model its previous values — priming it with a
    misread anchors it to that misread. The magnification is what does the work.
    """
    buf = io.BytesIO()
    _dtr_band_image(img, band).save(buf, format="PNG")
    b64 = base64.b64encode(buf.getvalue()).decode()
    instruction = _DTR_REREAD_PROMPT.replace(
        "%ROWS%", "\n".join("  " + d for d in dates)
    )
    payload = {
        "model": model or await get_model_name(),
        "messages": [{"role": "user", "content": [
            {"type": "image_url", "image_url": {"url": f"data:image/png;base64,{b64}"}},
            {"type": "text", "text": instruction},
        ]}],
        "max_tokens": 2048,
        "temperature": 0.0,
        "response_format": {"type": "json_object"},
    }
    async with httpx.AsyncClient(timeout=VERIFY_TIMEOUT) as client:
        resp = await client.post(
            f"{GEMINI_BASE_URL}/chat/completions", json=payload, headers=_gemini_headers()
        )
        resp.raise_for_status()
        raw = resp.json()["choices"][0]["message"]["content"]
    parsed = _extract_json_object(re.sub(r"<think>[\s\S]*?</think>", "", raw)) or {}
    out: dict[str, list[str]] = {}
    for entry in parsed.get("rows") or []:
        if not isinstance(entry, dict):
            continue
        values = entry.get("values")
        key = _dtr_date_key(entry.get("date"))
        if key and isinstance(values, list) and len(values) == len(DTR_TIME_SLOTS):
            out[key] = [str(v or "").strip() for v in values]
    return out


def _dtr_date_key(value: Any) -> str | None:
    """Month+day digits, for matching a re-read row back to the row it came from."""
    nums = re.findall(r"\d+", str(value or ""))
    if len(nums) < 2:
        return None
    return f"{int(nums[0])}-{int(nums[1])}"


async def _verify_dtr_times(table_data: dict, source: Image.Image) -> None:
    """Detect rows whose six stamps cannot be one day, then repair by zooming in.

    A correction is only accepted when an independent magnified re-read produces
    a digit-level change that makes the row chronological. Anything it cannot
    resolve is flagged for a human rather than overwritten with another guess —
    this is payroll data, and a silent wrong value is worse than a marked one.
    """
    headers: list[str] = table_data.get("headers") or []
    rows: list[list[str]] = table_data.get("rows") or []
    mapping = _dtr_map_columns(headers)
    if "date" not in mapping or not all(s in mapping for s in DTR_TIME_SLOTS):
        return
    seq = [mapping[s] for s in DTR_TIME_SLOTS]
    table_data["dtr_checked"] = True

    def row_times(row: list[str]) -> list[str]:
        return [row[i] if i < len(row) else "" for i in seq]

    suspect: list[int] = []
    schedule_flags: list[dict[str, Any]] = []
    for idx, row in enumerate(rows):
        schedule = (
            row[mapping["schedule"]]
            if "schedule" in mapping and mapping["schedule"] < len(row) else ""
        )
        if not _dtr_chronology_ok(row_times(row)):
            suspect.append(idx)
        elif _dtr_schedule_disagrees(row_times(row), schedule):
            schedule_flags.append({
                "row": idx, "col": mapping["schedule"],
                "reason": f"shift start {schedule!r} doesn't match the first punch",
            })

    corrections: list[dict[str, Any]] = []
    pending = list(suspect)
    if pending:
        logger.info("DTR check: %d row(s) out of chronological order", len(pending))

    # Rows needing a magnified look: the out-of-order ones always, every row when
    # the cross-check is on. Both share the same band reads, so turning the
    # cross-check on costs no more calls than it already takes to repair a row.
    wanted = set(pending) | (set(range(len(rows))) if DTR_CROSSCHECK else set())
    second_read: dict[int, list[str]] = {}
    for band in DTR_BANDS:
        outstanding = [i for i in sorted(wanted) if i not in second_read]
        if not outstanding:
            break
        dates = [str(rows[i][mapping["date"]]) for i in outstanding]
        try:
            fresh = await _reread_dtr_band(source, band, dates)
        except Exception as e:
            logger.warning("DTR band re-read failed (%s): %s", type(e).__name__, e)
            break
        for idx in outstanding:
            values = fresh.get(_dtr_date_key(rows[idx][mapping["date"]]) or "")
            if values:
                second_read[idx] = values

    still: list[int] = []
    for idx in pending:
        after = second_read.get(idx)
        before = row_times(rows[idx])
        changed = _dtr_accept_reread(before, after) if after else None
        if changed is None:
            still.append(idx)
            continue
        for c in changed:
            corrections.append({
                "row": idx, "col": seq[c],
                "column": headers[seq[c]] if seq[c] < len(headers) else "",
                "from": before[c], "to": after[c],
            })
            rows[idx][seq[c]] = after[c]
    pending = still

    # ── escalation tier ───────────────────────────────────────────────────
    # Only the rows the cheap pass could not settle get here, so the extra spend
    # is proportional to the difficulty of the page rather than its size. On the
    # client batches that is roughly 1-3% of rows.
    escalations: list[dict[str, Any]] = []
    if pending and DTR_ESCALATE:
        logger.info("DTR escalation: %d row(s) unresolved, looking closer", len(pending))
        for band in DTR_ESCALATION_BANDS:
            if not pending:
                break
            dates = [str(rows[i][mapping["date"]]) for i in pending]
            try:
                fresh = await _reread_dtr_band(source, band, dates,
                                               model=ESCALATION_MODEL or None)
            except Exception as e:
                logger.warning("DTR escalation re-read failed (%s): %s",
                               type(e).__name__, e)
                break
            remaining: list[int] = []
            for idx in pending:
                after = fresh.get(_dtr_date_key(rows[idx][mapping["date"]]) or "")
                before = row_times(rows[idx])
                changed = _dtr_accept_reread(before, after) if after else None
                if changed is None:
                    remaining.append(idx)
                    continue
                for c in changed:
                    corrections.append({
                        "row": idx, "col": seq[c],
                        "column": headers[seq[c]] if seq[c] < len(headers) else "",
                        "from": before[c], "to": after[c], "escalated": True,
                    })
                    rows[idx][seq[c]] = after[c]
                escalations.append({"row": idx, "band": band[:2],
                                    "model": ESCALATION_MODEL or "same"})
            pending = remaining
        if escalations:
            logger.info("DTR escalation: settled %d row(s)", len(escalations))
    table_data["dtr_escalations"] = escalations

    # Cross-check: a cell the two readings disagree on is not corrected — measured,
    # the crop is right about as often as the full page — but it is worth a human's
    # eye, because a wrong-but-chronological time is invisible to every other check.
    disagreements: list[dict[str, Any]] = []
    if DTR_CROSSCHECK:
        for idx, after in second_read.items():
            if idx in suspect:
                continue
            for c, value in enumerate(after):
                current = rows[idx][seq[c]] if seq[c] < len(rows[idx]) else ""
                a, b = _parse_hhmm(current), _parse_hhmm(value)
                if a is None or b is None or a == b:
                    continue
                disagreements.append({
                    "row": idx, "col": seq[c],
                    "reason": f"two readings disagree: {current} on the full page, "
                              f"{value} when zoomed in — please check",
                })
        if disagreements:
            logger.info("DTR cross-check: %d cell(s) disagreed", len(disagreements))

    flagged = [
        {"row": idx, "col": col,
         "reason": "times are not in chronological order — please check this row"}
        for idx in pending for col in seq
    ] + schedule_flags + disagreements

    if corrections:
        logger.info("DTR check: auto-corrected %d cell(s) by zooming in", len(corrections))
    if flagged:
        logger.info("DTR check: flagged %d cell(s) for review", len(flagged))
    table_data["time_corrections"] = table_data.get("time_corrections", []) + corrections
    table_data["flagged_cells"] = table_data.get("flagged_cells", []) + flagged


def _extract_json_object(text: str) -> dict | None:
    """Parse the first complete JSON object in text, tolerating trailing garbage.

    Gemini occasionally emits stray tokens after the closing brace (e.g. an
    extra "}\n]\n}"), which breaks a greedy first-{...last-} regex approach.
    raw_decode stops at the end of the first valid object instead.
    """
    decoder = json.JSONDecoder()
    idx = text.find("{")
    while idx != -1:
        try:
            obj, _ = decoder.raw_decode(text, idx)
            if isinstance(obj, dict):
                return obj
        except json.JSONDecodeError:
            pass
        idx = text.find("{", idx + 1)
    return None


def _try_partial_json(text: str) -> dict | None:
    """Extract 'headers' and 'full_text' from a truncated/malformed JSON string.

    Used when json.loads fails (e.g. full_text grew so large it hit the token
    limit and truncated the JSON).  Recovers headers reliably (they are short and
    appear early) and grabs as much of full_text as possible for OCR field detection.
    """
    result: dict = {}

    # headers array is short — almost always intact even in truncated responses
    m = re.search(r'"headers"\s*:\s*(\[)', text)
    if m:
        start = m.start(1)
        depth = 0
        for i, ch in enumerate(text[start:], start):
            if ch == "[":
                depth += 1
            elif ch == "]":
                depth -= 1
                if depth == 0:
                    try:
                        result["headers"] = json.loads(text[start : i + 1])
                    except Exception:
                        pass
                    break

    # full_text: grab the raw string value, stopping at an unescaped quote or 4000 chars
    m = re.search(r'"full_text"\s*:\s*"', text)
    if m:
        fragment = text[m.end():]
        ft_chars: list[str] = []
        idx = 0
        while idx < len(fragment) and len(ft_chars) < 4000:
            ch = fragment[idx]
            if ch == "\\" and idx + 1 < len(fragment):
                nch = fragment[idx + 1]
                ft_chars.append("\n" if nch == "n" else ("\t" if nch == "t" else nch))
                idx += 2
            elif ch == '"':
                break
            else:
                ft_chars.append(ch)
                idx += 1
        if ft_chars:
            result["full_text"] = "".join(ft_chars)

    return result if result else None


def _dtr_instructions(user_prompt: str = "") -> str:
    """Prompt for the DTR timesheets this tool is mainly pointed at.

    Measured against the generic extractor on the 10 reference2 scans (2 reps):
    cell accuracy 97.8% -> 98.8%, 12h/24h resolution 71% -> 100%, and the number
    of distinct header tuples the batch comes back with 9 -> 1.
    """
    parts = [
        "You are transcribing a Philippine Daily Time Record (DTR) timesheet.",
        "Return a single valid JSON object with the keys below IN THIS ORDER.",
        "",
        '"reasoning": 2-3 sentences — confirm this is a DTR, and note any cell whose',
        "handwriting you could not read confidently.",
        "",
        '"full_text": Document-level metadata ONLY (name, company assignment, day-off,',
        "outlet, the TOTAL NO. OF DAYS/HRS line, NO. OF REG. O.T. HRS, signatures).",
        'Write each as "LABEL: value" on its OWN line. Do not re-transcribe data rows.',
        "Keep under 200 words.",
        "",
        '"headers": Use EXACTLY this list, verbatim, always all 11, in this order:',
    ]
    parts += ["  " + h for h in DTR_HEADERS]
    parts += [
        "",
        '"rows": One array per used row of the grid, EXACTLY 11 values each.',
        "",
        "HOW TO READ THE GRID",
        "The main grid is on the LEFT. Each used row is one working day and holds six",
        "clock times in this fixed left-to-right order:",
        "  MORNING IN -> MORNING OUT -> AFTERNOON IN -> AFTERNOON OUT -> EVENING IN -> EVENING OUT",
        "1. CHRONOLOGY. Those six times run forward through a single day and must strictly",
        "   increase in that order. If your reading breaks the order, you have misread a",
        "   digit -- look again at that cell before answering.",
        "2. 24-HOUR OUTPUT. Write every time as 24-hour HH:MM. Staff write the afternoon",
        '   and evening in 12-hour form, so "2:46" in the AFTERNOON OUT column is 14:46',
        '   and "9:16" in EVENING OUT is 21:16. Use the DAILY TIME SCHEDULE cell of that',
        '   same row to decide: a "12:15-9:15" schedule means the shift ends 21:15, so an',
        '   evening-out near "9:15" is 21:15, never 09:15 and never 19:15.',
        "   Only the MORNING IN column is normally genuinely AM.",
        "3. SECOND COPY. Most of these scans also carry a punch card or payroll card on",
        "   the RIGHT side, listing the same six machine-printed stamps against a day",
        "   number. When a day appears on both sides they are the SAME shift. The machine",
        "   print is more reliable than the handwriting -- use it to confirm or correct",
        "   every digit you take from the left grid.",
        "4. DIGITS. Take special care with 6 vs 8, 0 vs 8, 1 vs 7, 2 vs 7, 5 vs 6, 9 vs 4.",
        "5. DATES. Dates are numeric (6-26-26, 6/26/26, 06-26). All rows on one form fall",
        "   inside one pay period, so the month and year agree across rows.",
        "6. BLANKS. A cell that is empty, dashed or struck through -> \"\". Never shift a",
        "   value left or right to fill a gap; keep every column in its own position.",
        '7. ANNOTATIONS. A row reading "REST DAY" (or similar) written across the grid:',
        '   put the words in MORNING IN and leave the other time cells "".',
        "8. Red pen marks are a checker's ticks, not data. Ignore them unless they are the",
        "   only value in the cell.",
        '9. "NO. OF DAYS/HRS", "O.T. HRS" and "REASON FOR O.T." are often blank -> "".',
        "   Still emit all 11 values per row.",
        "",
        'In "full_text", note that a total written like "3 D / 24 HRS" means 3 Days --',
        "that character is a letter D, not a zero.",
    ]
    if user_prompt.strip():
        parts += ["", "Additional instructions for the rows:", user_prompt.strip()]
    parts += ["", "Output ONLY valid JSON. No markdown fences, no explanation."]
    return "\n".join(parts)


def _inout_instructions(user_prompt: str = "") -> str:
    """Prompt for the single in/out payroll sheet (24 Chicken and similar)."""
    parts = [
        "You are transcribing a payroll time sheet with ONE punch pair per day.",
        "Return a single valid JSON object with the keys below IN THIS ORDER.",
        "",
        '"reasoning": 2-3 sentences confirming the form and noting any unclear cell.',
        "",
        '"full_text": Document-level metadata ONLY (employee name, position, period,',
        'store/outlet, any totals printed outside the grid). "LABEL: value" per line.',
        "",
        f'"headers": Use EXACTLY this list, verbatim, always all {len(INOUT_HEADERS)}, in this order:',
    ]
    parts += ["  " + h for h in INOUT_HEADERS]
    parts += [
        "",
        f'"rows": One array per row of the grid, EXACTLY {len(INOUT_HEADERS)} values each.',
        "",
        "HOW TO READ IT",
        "1. Each row is one calendar day: a DATE, the weekday, the SCHEDULE, one TIME IN",
        "   and one TIME OUT, then the sheet's own computed columns.",
        '   SCHEDULE: copy the schedule cell exactly as written ("12-10", "9-7", "RD",',
        '   "-"). Do not convert it to clock times. "" when the sheet has no SCHEDULE column.',
        "2. TIME IN must be earlier in the day than TIME OUT. Write both as 24-hour",
        "   HH:MM. A dash, blank or '-' means no punch that day -> \"\".",
        "3. BREAK (HR), NORMAL, OT HRS, TOTAL, DEDUCTIONS and NET TOTAL are numbers",
        "   the sheet already prints. Copy them exactly, including 0 and 0.00.",
        "   Do not compute or correct them.",
        "4. Keep every column in its own position. If a column is absent from this",
        '   sheet, emit "" for it rather than shifting the others left.',
        "5. Take care with 6 vs 8, 0 vs 8, 1 vs 7, 2 vs 7, 5 vs 6, 9 vs 4.",
        "6. REMARKS: words or codes handwritten on that row on top of the printed",
        '   sheet, e.g. WDO, SPL, SH, LH, OT, "1 OT", "1 NDF", REST DAY. Copy them as',
        '   written. "" when there are none. Ticks, check marks and initials are not remarks.',
        "   Handwriting is often written over the OT HRS or HOLIDAY PAY cells: put it in",
        "   REMARKS, and keep OT HRS and HOLIDAY PAY to the numbers the sheet PRINTS.",
        '7. HOLIDAY PAY and HOLIDAY+OT: copy what is printed, blank stays "".',
        "8. NORMAL (9am-10PM) is usually printed as TWO sub-columns, hours then pay",
        '   (e.g. "8 | 600"). Put only the hours in NORMAL; the pay has no column here.',
        "   Keep counting columns from the printed header, so OT HRS stays OT HRS.",
        "9. A number printed to the RIGHT of NET TOTAL, outside the grid (e.g. 50 or 60,",
        "   an allowance), belongs to no column. Ignore it; it is NOT a remark.",
        "10. Return one row for EVERY date printed in the DATE column, in order, even when",
        "   handwriting covers the row. Never merge, skip or drop a printed date.",
    ]
    if user_prompt.strip():
        parts += ["", "Additional instructions for the rows:", user_prompt.strip()]
    parts += ["", "Output ONLY valid JSON. No markdown fences, no explanation."]
    return "\n".join(parts)


def _timesheet_instructions(user_prompt: str = "") -> str:
    """Prompt for the break-detail timesheets (regular / lunch / coffee / overtime).

    Pinning the header list is the whole point: the printed forms themselves come
    in 12-, 13-, 14-, 15- and 17-column variants, so without this the same batch
    returns a different schema per page and the pages refuse to line up.
    """
    parts = [
        "You are transcribing a printed employee timesheet.",
        "Return a single valid JSON object with the keys below IN THIS ORDER.",
        "",
        '"reasoning": 2-3 sentences — confirm the form type and note any cell you',
        "could not read confidently.",
        "",
        '"full_text": Document-level metadata ONLY (Employee #, Name, Position,',
        "Store, Period Covered, Prepared by, and any TOTAL lines printed outside",
        'the grid). Write each as "LABEL: value" on its OWN line. Do not',
        "re-transcribe data rows. Keep under 200 words.",
        "",
        '"headers": Use EXACTLY this list, verbatim, always all 17, in this order:',
    ]
    parts += ["  " + h for h in TIMESHEET_HEADERS]
    parts += [
        "",
        '"rows": One array per printed day, EXACTLY 17 values each.',
        "",
        "HOW TO READ THE GRID",
        "Each row is one day. The grid groups its columns into sections, each with",
        "its own in / out / TOTAL:",
        "  REGULAR HOURS · LUNCH BREAK · COFFEE BREAK · OVERTIME",
        "1. NOT EVERY FORM HAS EVERY COLUMN. Some print only a LUNCH BREAK total and",
        "   no lunch in/out; some have no DUTY LOCATION or TARDINESS column. Map what",
        '   the form does have onto the list above and put "" in every slot the form',
        "   does not print. Never shift a value into a neighbouring column to close a",
        "   gap, and never drop a column from the list because this page lacks it.",
        "2. TOTALS ARE PRINTED, NOT COMPUTED BY YOU. Each TOTAL is on the form —",
        "   transcribe it as printed, even when it disagrees with out minus in.",
        "3. 24-HOUR TIMES. Transcribe times as printed; these forms already use",
        "   24-hour (8:50, 18:00, 21:00).",
        "4. DIGITS. Take special care with 6 vs 8, 0 vs 8, 1 vs 7, 2 vs 7, 5 vs 6.",
        '5. NON-TIME ENTRIES. A day marked RD, UA, ABSENT, LEAVE or similar: put that',
        '   text in REGULAR HOURS - Time in and leave the other time cells "".',
        "6. DATES. Transcribe the date exactly as printed, including the weekday if",
        '   the form prints one (e.g. "Thursday, July 16, 2026").',
        "7. Red ink marks rest days and holidays — transcribe the value, ignore the",
        "   colour.",
    ]
    if user_prompt.strip():
        parts += ["", "Additional instructions for the rows:", user_prompt.strip()]
    parts += ["", "Output ONLY valid JSON. No markdown fences, no explanation."]
    return "\n".join(parts)


@app.post("/api/ocr")
async def ocr_image(
    file: UploadFile = File(...),
    prompt: str = Form(""),
    document_type: str = Form("auto"),
    client: str = Form(""),
    page: int = Form(0),
):
    """Extract a timesheet/DTR scan into the client's fixed column schema."""
    # "dtr" and "timesheet" each use a form-specific prompt with a pinned header
    # list. "auto" runs the generic extractor but still canonicalises whichever of
    # the two the result turns out to be — and it is the default, because forcing
    # one form's header list onto the other's scans mangles them.
    # A chosen client pins the layout; document_type stays as the low-level knob
    # underneath it, so an explicit document_type still wins for one-off files.
    client_key = (client or "").strip().lower()
    if client_key and document_type in ("", "auto"):
        layout = client_registry.get_client(client_key).get("layout")
        document_type = {"dtr": "dtr", "dtr_ot": "dtr",
                         "timesheet": "timesheet",
                         "inout": "inout"}.get(layout or "", "auto")

    document_type = (document_type or "auto").strip().lower()
    if document_type not in {"auto", "dtr", "timesheet", "inout", "generic"}:
        raise HTTPException(
            400, "Invalid document_type. Use 'auto', 'dtr', 'timesheet', 'inout' or 'generic'."
        )

    if not (_is_image_upload(file.content_type, file.filename) or _is_pdf_upload(file.content_type, file.filename)):
        raise HTTPException(400, "Only image or PDF files are supported")

    contents = await file.read()
    if len(contents) > MAX_FILE_SIZE:
        raise HTTPException(400, "File too large (max 20 MB)")

    # Read this off the original upload: preprocessing re-encodes to PNG, which
    # carries no EXIF, so the tag is gone by the time the response is assembled.
    upload_exif_rotated = _exif_rotated(contents)
    # Kept for the DTR band re-read, which crops from the source pixels: magnifying
    # the already-upscaled model copy would only interpolate what is there.
    original_bytes = contents

    try:
        contents, mime_type = _prepare_upload_for_model(
            contents, file.content_type, file.filename, max(0, int(page or 0))
        )
    except UnidentifiedImageError:
        raise HTTPException(400, "Unsupported or corrupted image file")
    except Exception as e:
        if _is_pdf_upload(file.content_type, file.filename):
            raise HTTPException(400, f"Unsupported or corrupted PDF file: {e}")
        raise HTTPException(400, f"Unable to process uploaded file: {e}")

    b64_image = base64.b64encode(contents).decode("utf-8")

    has_user_prompt = bool(prompt.strip())

    if document_type in {"dtr", "timesheet", "inout"}:
        instr_parts = [
            _dtr_instructions(prompt) if document_type == "dtr"
            else _inout_instructions(prompt) if document_type == "inout"
            else _timesheet_instructions(prompt)
        ]
    else:
        # Universal VLM call — analyse then extract (image-agnostic)
        instr_parts = [
            "You are a universal document-to-table extractor that works on ANY type of image containing tabular data.",
            "Analyse the image in two phases, then return a single valid JSON object with the keys below IN THIS ORDER.",
            "",
            "PHASE 1 — ANALYSE (answer inside \'reasoning\'):",
            "  (a) Document category: lab sheet, timesheet, invoice, schedule, receipt, financial report, form, etc.",
            "  (b) Table orientation:",
            "      STANDARD  — column headers across the top row, data rows below.",
            "      ROW-LABEL — row-label stubs down the left edge (e.g. \'Mass of beaker\', \'Monday\') PLUS column headers across the top.",
            "  (c) Document-level metadata: any fields printed OUTSIDE the main table (name, date, ID, title, experiment number, etc.).",
            "",
            "PHASE 2 — EXTRACT (populate the JSON keys):",
            "",
            '"reasoning": 2-4 sentences from Phase 1.',
        ]

        # Include fields schema only when template metadata fields exist

        instr_parts.append('"full_text": Concise transcription of document-level metadata only — form fields, titles, signature lines, and summary totals. For each labeled form field always write "LABEL: value" with a colon (e.g. "DAY-OFF: Tuesday", "OUTLET: Landmark"). Put each field on its OWN line separated by \\n — NEVER join multiple fields on one line with periods or commas. For the main data table include only the column header row — do NOT re-transcribe every data row (they are captured in "rows" below). Keep full_text under 300 words.')

        if True:
            instr_parts += [
                '"headers": Column header strings of the MAIN DATA TABLE only.',
                '  STANDARD table  → text labels in the top header row (e.g. ["Employee","Date","Hours"]).',
                '  ROW-LABEL table → build headers by combining two groups IN ORDER:',
                '    GROUP 1 — left-edge stub column(s): infer a descriptive name for each from context.',
                '      Single stub  (e.g. lab data):  ["Measurement", ...]',
                '      Two stubs    (e.g. timesheet):  ["Daily Time Schedule","Date", ...]',
                '    GROUP 2 — the column labels at the TOP of the data grid:',
                '      CRITICAL: labels like "Trial 1","Trial 2","Morning In" are column headers, NOT row values.',
                '      They MUST appear in "headers" — never inside a row array.',
                '      Full examples:',
                '        Single stub: ["Measurement","Trial 1","Trial 2"]',
                '        Two stubs:   ["Daily Time Schedule","Date","Morning In","Morning Out","Afternoon In","Afternoon Out","Evening In","Evening Out"]',
                '  NEVER use document-level metadata fields (Exp. No., Name, Date, Lab Partner, ID, etc.) as column headers.',
                '  NEVER use single letters (A, B, C...) or bare numbers as column headers.',
                '  Combine multi-tier / nested headers into one string per column (e.g. "Category - Subcategory").',
            ]
        instr_parts += [
            '"rows": Array of row arrays — each row has EXACTLY len(headers) values.',
            '  ROW-LABEL table → include values for ALL left-edge label columns first, then the data values.',
            '    MERGED / ROW-SPANNING cells: write the value ONLY in the first row it covers; use "" for every subsequent row it spans.',
            '    Example: if "10 x 7" spans dates 18 and 19, row for 18 = ["10 x 7","18",...], row for 19 = ["","19",...].',
            '  STANDARD table  → data values only, left to right.',
            '  Empty / blank / dashed / crossed-out cells → "" — ALWAYS include the cell position, never skip or collapse it.',
            '  Do NOT include document-level metadata (experiment info, name, date, etc.) as data rows.',
            '  Do NOT place the header row text inside rows.',
            "",
            "RULES:",
            '1. FIDELITY  — Transcribe values exactly as written. Keep units, formulas, symbols (e.g. "(0.2027/18.02) = 0.01125").',
            '2. OCR FIX   — In numeric/date/time columns ONLY, correct obvious OCR artefacts (O→0, l→1). Leave text/name columns unchanged.',
            '3. ALIGNMENT — Every row must have exactly len(headers) elements. Pad with "" where a cell is absent.',
            '4. ANOMALIES — Deliberate non-numeric text in a numeric column (N/A, Void, -, n.d.) → transcribe exactly as-is.',
            '5. MULTI-TABLE — If multiple tables exist, extract the most data-rich one (highest rows × columns product).',
            '6. MERGED CELLS — Row-spanning cells: value in first covered row, "" in all subsequent covered rows. Never repeat the merged value.',
        ]

        if has_user_prompt:
            instr_parts += ["", "Transformation instructions for the rows:", prompt.strip()]

        instr_parts += ["", "Output ONLY valid JSON. No markdown fences, no explanation."]

    payload = {
        "model": await get_model_name(),
        "messages": [
            {
                "role": "user",
                "content": [
                    {
                        "type": "image_url",
                        "image_url": {"url": f"data:{mime_type};base64,{b64_image}"},
                    },
                    {"type": "text", "text": "\n".join(instr_parts)},
                ],
            }
        ],
        "max_tokens": 8192,
        "temperature": 0.1,
        "response_format": {"type": "json_object"},
    }
    # Note "reasoning_effort", not "thinking_config": the latter is rejected with a
    # 400 ("Unknown name") by the OpenAI-compat surface. See REASONING_EFFORT for
    # why the level is configurable rather than pinned.
    if REASONING_EFFORT:
        payload["reasoning_effort"] = REASONING_EFFORT

    try:
        async with httpx.AsyncClient(timeout=OCR_TIMEOUT) as client:
            resp = await client.post(
                f"{GEMINI_BASE_URL}/chat/completions", json=payload, headers=_gemini_headers()
            )
            resp.raise_for_status()
            raw_result = resp.json()
    except httpx.HTTPStatusError as e:
        # Pass a quota 429 through as a 429 rather than flattening it into a 502.
        # The two need completely different handling: a 502 is worth retrying in a
        # second, a quota exhaustion needs backing off for tens of seconds, and the
        # client can only tell them apart if the status survives.
        if e.response.status_code == 429:
            raise HTTPException(
                429,
                "Gemini rate limit reached — the batch is going faster than the "
                "project's quota allows. Retrying shortly.",
                headers={"Retry-After": e.response.headers.get("Retry-After", "20")},
            )
        raise HTTPException(502, f"OCR model error: {e.response.text}")
    except (httpx.ConnectError, httpx.ConnectTimeout):
        raise HTTPException(503, "Cannot connect to Gemini API")
    except httpx.HTTPError as e:
        # Read/write/timeout errors mid-request: transient, so surface a retryable
        # status instead of a 500 — batches otherwise lose a file to a blip.
        raise HTTPException(503, f"Gemini API request failed ({type(e).__name__}) — please retry")

    raw_content = raw_result["choices"][0]["message"]["content"]
    raw_content = re.sub(r"<think>[\s\S]*?</think>", "", raw_content).strip()
    logger.info("VLM response (first 2000): %s", raw_content[:2000])

    # Initialise table_data with safe defaults
    table_data: dict = {
        "mode": "table",
        "headers": [], "rows": [], "text_above": [], "text_below": [],
        "raw": raw_content, "llm_applied": False, "llm_raw": raw_content,
    }
    full_ocr_text = raw_content  # fallback if JSON parse fails

    try:
        result_json = _extract_json_object(raw_content)
        if result_json is None:
            raise ValueError("No JSON object found in VLM response")

        full_ocr_text = result_json.get("full_text", raw_content)

        # Headers
        raw_headers = result_json.get("headers", [])
        if isinstance(raw_headers, list):
            table_data["headers"] = [str(h) for h in raw_headers]

        # Rows — normalise to n columns, drop empty rows
        n = len(table_data["headers"])
        raw_rows = result_json.get("rows", [])
        if isinstance(raw_rows, list):
            cleaned: list[list[str]] = []
            for row in raw_rows:
                if isinstance(row, list):
                    r = [str(c) for c in row]
                    while len(r) < n:
                        r.append("")
                    if n:
                        r = r[:n]
                    if any(c for c in r):
                        cleaned.append(r)
            # Filter separator rows (every cell is only dashes, colons, or whitespace)
            cleaned = [
                row for row in cleaned
                if not all(re.match(r'^[-:\s]*$', cell) for cell in row)
            ]
            table_data["rows"] = cleaned

        # A recognised DTR gets the fixed 11-column schema instead, so every page of
        # a batch lines up. Do this before the empty-column strip, which is what
        # otherwise drops whichever of NO. OF DAYS/HRS or O.T. HRS happens to be
        # blank on this particular page.
        # A break-detail timesheet gets its own fixed 17-column schema for the same
        # reason — those forms print in 12- to 17-column variants.
        table_data["is_dtr"] = False
        table_data["is_timesheet"] = False
        table_data["is_inout"] = False
        if document_type != "generic":
            # Before the schema lock, or a stray heading row becomes row 1 of the
            # employee's block and "Time Out" is written where a time belongs.
            _strip_leading_header_rows(table_data)
            table_data["is_dtr"] = _dtr_canonicalise(table_data)
            if not table_data["is_dtr"]:
                table_data["is_timesheet"] = _timesheet_canonicalise(table_data)
            if not (table_data["is_dtr"] or table_data["is_timesheet"]):
                table_data["is_inout"] = _inout_canonicalise(table_data)

        # Strip entirely-empty columns after row normalisation. Skipped for both
        # canonical schemas: dropping a column that is blank on this page is
        # exactly what made the batch fail to line up.
        _locked = (table_data["is_dtr"] or table_data["is_timesheet"]
                   or table_data.get("is_inout"))
        if table_data["headers"] and table_data["rows"] and not _locked:
            _n_h = len(table_data["headers"])
            _col_ok = [any(row[c] for row in table_data["rows"]) for c in range(_n_h)]
            if not all(_col_ok):
                _keep = [c for c, ok in enumerate(_col_ok) if ok]
                table_data["headers"] = [table_data["headers"][c] for c in _keep]
                table_data["rows"] = [[row[c] for c in _keep] for row in table_data["rows"]]

        # Heuristic: fix merged schedule cells where the model duplicated the adjacent
        # date/day value instead of leaving the cell empty.  This happens when a
        # row-spanning "Daily Time Schedule" cell covers multiple date rows — the
        # model incorrectly copies the date into col-0 for the non-first rows.
        # Only applied when headers clearly look like (schedule-ish, date-ish, ...).
        if len(table_data["headers"]) >= 2 and table_data["rows"]:
            _h0 = table_data["headers"][0].lower()
            _h1 = table_data["headers"][1].lower()
            _sched_col = any(w in _h0 for w in ("schedule", "time", "shift", "duty", "daily"))
            _date_col  = any(w in _h1 for w in ("date", "day", "dates"))
            if _sched_col and _date_col:
                for _row in table_data["rows"]:
                    # If schedule value equals the date value it is almost certainly
                    # a spurious copy — clear the schedule cell.
                    if _row[0].strip() and _row[0].strip() == _row[1].strip():
                        _row[0] = ""

        if has_user_prompt and table_data["rows"]:
            table_data["llm_applied"] = True


    except Exception as e:
        logger.warning("VLM JSON parse failed (%s): %s — trying partial extraction", type(e).__name__, e)
        # Attempt to recover structured data from a truncated / malformed JSON response.
        # This prevents the full_text value (which may contain | characters) from being
        # misinterpreted as a pipe-delimited markdown table by the parse_table fallback.
        partial = _try_partial_json(raw_content)
        if partial and isinstance(partial.get("headers"), list) and len(partial["headers"]) > 0:
            logger.info("Partial JSON recovery: %d headers", len(partial["headers"]))
            table_data["headers"] = [str(h) for h in partial["headers"]]
            if partial.get("full_text"):
                full_ocr_text = partial["full_text"]
        else:
            logger.warning("Partial extraction failed — falling back to table parser")
            fallback = parse_table(raw_content)
            table_data["headers"] = fallback.get("headers", [])
            table_data["rows"] = fallback.get("rows", [])
            non_table = _extract_non_table_text(raw_content)
            table_data["text_above"] = non_table["text_above"]
            table_data["text_below"] = non_table["text_below"]


    # Split full OCR text into header lines (before table) and footer lines (after table).
    # Uses the extracted grid data to locate the table region in the text, then keeps only
    # lines that look like metadata key:value fields.
    _time_re = re.compile(r'\b\d{1,2}:\d{2}\b')

    def _is_field_line(ln: str) -> bool:
        ln = ln.strip()
        if not ln:
            return False
        idx = ln.find(':')
        if not (0 < idx <= 60):
            return False
        label = ln[:idx].strip()
        value = ln[idx + 1:].strip()
        if not value:
            return False
        if not any(c.isalpha() for c in label):
            return False
        # The model sometimes echoes the column header row as "HEADERS: a, b, c"
        if label.lower().rstrip('s') in {"header", "column", "table header", "column header"}:
            return False
        # A line with 2+ time patterns is a table data row, not a field
        if len(_time_re.findall(ln)) >= 2:
            return False
        return True

    # Build set of grid values for cross-referencing
    grid_vals: set[str] = set()
    for row in table_data.get("rows", []):
        for cell in row:
            v = str(cell).strip().lower()
            if v and len(v) > 1:
                grid_vals.add(v)
    for h in table_data.get("headers", []):
        v = h.strip().lower()
        if v and len(v) > 1:
            grid_vals.add(v)

    # Normalised header strings for label cross-reference (prevents table row labels
    # from being misclassified as metadata header/footer fields)
    _header_norms: set[str] = set()
    for h in table_data.get("headers", []):
        norm = re.sub(r'[^a-z0-9 ]', ' ', h.lower()).strip()
        if len(norm) > 4:
            _header_norms.add(norm)

    def _is_table_line(ln: str) -> bool:
        """True if this line is likely a table row (dense grid values or 2+ time patterns)."""
        if len(_time_re.findall(ln)) >= 2:
            return True
        tokens = re.findall(r'[\w:×·]+', ln.lower())
        if not tokens:
            return False
        hits = sum(1 for t in tokens if t in grid_vals)
        return hits / len(tokens) >= 0.4

    def _is_footer_label(label_text: str) -> bool:
        """Semantic footer detector for totals/signatures that may appear near the top-right of forms."""
        label_norm = re.sub(r'[^a-z0-9 ]', ' ', (label_text or '').lower())
        label_norm = re.sub(r'\s+', ' ', label_norm).strip()
        if not label_norm:
            return False

        footer_keywords = (
            "total regular",
            "total overtime",
            "overtime",
            "total no of days",
            "no of reg",
            "reg ot hrs",
            "hours",
            "holiday",
            "signature",
            "employee s signat",
            "supervisor s signa",
            "approved by",
            "verified by",
        )
        return any(k in label_norm for k in footer_keywords)

    # The model sometimes writes the line breaks in full_text as a literal "\n",
    # which ran every header field into the employee name (Oppa Foods, 24 Nov).
    ocr_lines_all = re.sub(r"(?:\\r)?\\n", "\n", str(full_ocr_text)).splitlines()

    # Find the table region in the OCR text.
    # Exclude lines that look like metadata key:value fields — common tokens like
    # "day", "off", "date" would otherwise cause field lines (e.g. "DAY-OFF : Tuesday")
    # to be misclassified as table rows, which shifts table_start too early.
    table_start: int | None = None
    table_end: int | None = None
    for i, ln in enumerate(ocr_lines_all):
        if _is_table_line(ln) and not _is_field_line(ln):
            if table_start is None:
                table_start = i
            table_end = i

    # Collect header and footer field lines based on position relative to table
    header_ocr_lines: list[str] = []
    footer_ocr_lines: list[str] = []
    for i, ln in enumerate(ocr_lines_all):
        if not _is_field_line(ln):
            continue
        # Cross-reference: skip if the value after the colon is just a raw grid cell
        _colon_i = ln.find(':')
        val_part = ln[_colon_i + 1:].strip().lower()
        if val_part in grid_vals:
            continue
        # Also check the first value token (handles compound "val | key: val2" lines)
        _first_val = re.split(r'[\s|,;]', val_part)[0] if val_part else ''
        if _first_val and _first_val in grid_vals:
            continue
        # Skip if the label part closely matches a table column header
        # (these are table data rows mistaken for key:value metadata fields)
        _label_core = re.sub(r'[^a-z0-9 ]', ' ', ln[:_colon_i].lower()).strip()
        _label_core = re.sub(r'\s*\|?\s*trial\s*\d*\s*$', '', _label_core).strip()
        if _header_norms and any(
            _label_core.startswith(h) or h.startswith(_label_core)
            for h in _header_norms if len(h) > 4
        ):
            continue
        # Also skip if the label core matches a grid cell value (e.g. row-label column value)
        if _label_core and _label_core in grid_vals:
            continue

        # Semantic override: totals/signatures/day-count lines belong in footer
        # even when physically printed near the top-right of a form.
        if _is_footer_label(_label_core):
            footer_ocr_lines.append(ln.strip())
            continue

        if table_start is None or i < table_start:
            header_ocr_lines.append(ln.strip())
        elif table_end is not None and i > table_end:
            footer_ocr_lines.append(ln.strip())

    table_data["ocr_text_lines"] = header_ocr_lines
    table_data["ocr_footer_lines"] = footer_ocr_lines
    # Lets the preview decide which files still need turning the right way up.
    table_data["exif_rotated"] = upload_exif_rotated

    # Drop the form's own grand-total line before anything indexes the rows, then
    # report any day the table is missing so the client can ask for a re-read.
    _strip_trailing_summary_row(table_data)
    table_data["missing_rows"] = _find_missing_rows(table_data)
    if table_data["missing_rows"]:
        logger.info("Missing row(s) detected: %s", ", ".join(table_data["missing_rows"]))

    # DTRs have no TOTAL column to check against, so they get the chronology check
    # instead: a row whose six stamps run backwards is provably misread, and a
    # magnified re-read of that band usually resolves it.
    if TIME_VERIFY and table_data.get("is_dtr"):
        try:
            await _verify_dtr_times(
                table_data,
                _source_image(original_bytes, file.content_type, file.filename,
                              max(0, int(page or 0))),
            )
        except Exception as e:
            logger.warning("DTR time check skipped (%s): %s", type(e).__name__, e)

    # Timesheets print their own totals, so verify the extracted in/out/total
    # arithmetic and repair single-digit misreads before returning.
    if TIME_VERIFY:
        try:
            await _verify_time_math(table_data, b64_image, mime_type)
        except Exception as e:
            logger.warning("Time-math verification skipped (%s): %s", type(e).__name__, e)

    return table_data


@app.post("/api/amend")
async def amend_table(req: AmendRequest):
    """Apply user amendment instructions to table/field structure using existing LLM service."""
    instruction = (req.instruction or "").strip()
    if not instruction:
        raise HTTPException(400, "Instruction cannot be empty")

    headers = [str(h).strip() for h in (req.headers or []) if str(h).strip()]
    if not headers:
        headers = ["Column 1"]
    rows = _normalise_rows(headers, req.rows or [])
    header_fields = _normalise_fields(req.header_fields or [])
    footer_fields = _normalise_fields(req.footer_fields or [])

    current_data = {
        "headers": headers,
        "rows": rows,
        "header_fields": header_fields,
        "footer_fields": footer_fields,
    }

    system_prompt = (
        "You are a spreadsheet amendment assistant. "
        "Given current extracted OCR data and a user instruction, apply amendments exactly and return ONLY valid JSON. "
        "You may edit values and structure (add/remove/reorder columns, rows, or fields) when requested. "
        "When the original OCR extraction is provided, use it as reference context — e.g. to answer questions "
        "about the source document or to restore original values the user has edited away."
    )
    ocr_context = (req.ocr_context or "").strip()[:12000]
    context_block = (
        "Original OCR extraction JSON (reference context from the source document):\n"
        + ocr_context
        + "\n\n"
        if ocr_context
        else ""
    )
    user_prompt = (
        context_block
        + "Current data JSON:\n"
        + json.dumps(current_data, ensure_ascii=False)
        + "\n\n"
        + "User amendment instruction:\n"
        + instruction
        + "\n\n"
        + "Return JSON with this schema exactly:\n"
        + "{\n"
        + '  "assistant_message": "short confirmation of what changed",\n'
        + '  "headers": ["..."],\n'
        + '  "rows": [["..."], ["..."]],\n'
        + '  "header_fields": [{"label":"...","value":"..."}],\n'
        + '  "footer_fields": [{"label":"...","value":"..."}]\n'
        + "}\n"
        + "Rules: rows must have exactly len(headers) cells, use empty string for blanks, preserve existing data unless instruction changes it."
    )

    payload = {
        "model": await get_model_name(),
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_prompt},
        ],
        "max_tokens": 4096,
        "temperature": 0.1,
        "response_format": {"type": "json_object"},
    }

    try:
        async with httpx.AsyncClient(timeout=120.0) as client:
            resp = await client.post(
                f"{GEMINI_BASE_URL}/chat/completions", json=payload, headers=_gemini_headers()
            )
            resp.raise_for_status()
            raw = resp.json()["choices"][0]["message"]["content"]
    except httpx.HTTPStatusError as e:
        raise HTTPException(502, f"LLM error: {e.response.text}")
    except (httpx.ConnectError, httpx.ConnectTimeout):
        raise HTTPException(503, "Cannot connect to Gemini API")
    except httpx.HTTPError as e:
        raise HTTPException(503, f"Gemini API request failed ({type(e).__name__}) — please retry")

    try:
        result = _extract_json_object(raw)
        if result is None:
            raise ValueError("No JSON object found in LLM response")
    except Exception as e:
        raise HTTPException(502, f"Invalid LLM JSON response: {e}")

    out_headers = [str(h).strip() for h in result.get("headers", []) if str(h).strip()]
    if not out_headers:
        out_headers = headers
    out_rows = _normalise_rows(out_headers, result.get("rows", []))
    out_header_fields = _normalise_fields(result.get("header_fields", []))
    out_footer_fields = _normalise_fields(result.get("footer_fields", []))
    assistant_message = str(result.get("assistant_message", "Amendment applied.")).strip() or "Amendment applied."

    return {
        "assistant_message": assistant_message,
        "headers": out_headers,
        "rows": out_rows,
        "header_fields": out_header_fields,
        "footer_fields": out_footer_fields,
    }


def _header_safe(text: str) -> str:
    """Make text safe to put in an HTTP header.

    Starlette encodes header values as latin-1, so a single character outside it
    raises UnicodeEncodeError and turns a perfectly good workbook into a 500.
    That is not hypothetical: the warnings quote employee names, and a name like
    "Vergaño" — or the em dash the messages themselves used — is enough.
    """
    replacements = {"—": "-", "–": "-", "‘": "'", "’": "'",
                    "“": '"', "”": '"', "…": "..."}
    for bad, good in replacements.items():
        text = text.replace(bad, good)
    return text.encode("ascii", "replace").decode("ascii")


class DTRWorkbookFile(BaseModel):
    name: str = ""            # employee name, as shown on the card
    file: str = ""            # source filename, for the warning messages
    headers: list[Any] = []
    rows: list[Any] = []
    flagged_rows: list[Any] = []   # row indices the time check couldn't verify


class TimekeepingTag(BaseModel):
    employee: int             # index into files
    date: str                 # YYYY-MM-DD
    tag: str                  # "" (regular), WDO, SH, SHWDO, LH, LHWDO, ULH


class TimekeepingAdjustment(BaseModel):
    employee: int             # index into files
    date: str                 # YYYY-MM-DD
    minutes: int = 0          # added to the day's Late/UT (an overbreak the times don't show)


class TimekeepingHoliday(BaseModel):
    date: str                 # YYYY-MM-DD
    kind: str = "special"     # "special" or "regular"
    name: str = "Local holiday"


class DTRWorkbookRequest(BaseModel):
    files: list[DTRWorkbookFile] = []
    client: str = ""          # which client's workbook template to fill
    # For clients the app calculates (timekeeping.PROFILES): the review screen's
    # confirmed tags and local holidays, so the export matches what was reviewed.
    tags: list[TimekeepingTag] = []
    holidays: list[TimekeepingHoliday] = []
    adjustments: list[TimekeepingAdjustment] = []


@app.post("/api/pdf-info")
async def pdf_info(file: UploadFile = File(...)):
    """How many pages a PDF has, so the client can queue one job per page."""
    if not _is_pdf_upload(file.content_type, file.filename):
        raise HTTPException(400, "Not a PDF")
    contents = await file.read()
    if len(contents) > MAX_FILE_SIZE:
        raise HTTPException(400, "File too large (max 20 MB)")
    try:
        return {"pages": _pdf_page_count(contents)}
    except Exception as e:
        raise HTTPException(400, f"Unreadable PDF: {e}")


@app.get("/api/clients")
async def list_clients():
    """The client picker's catalogue: label, input layout and whether we hold
    a timekeeping workbook for them."""
    return {"clients": client_registry.catalogue(),
            "default": client_registry.DEFAULT_CLIENT}


@app.post("/api/dtr-workbook")
async def dtr_workbook(req: DTRWorkbookRequest):
    """Compile a whole batch of extracted DTRs into the client timekeeping workbook.

    One Sheet1 block per employee (pages of the same person merged and sorted by
    date), plus the DTR Summary tab. Formulas are the client's own, row-shifted.
    """
    layout_name = (client_registry.get_client(req.client) or {}).get("layout")
    payload: list[dict] = []
    for f in req.files:
        headers = [str(h) for h in (f.headers or [])]
        rows = [[str(c) if c is not None else "" for c in r]
                for r in (f.rows or []) if isinstance(r, list)]
        # Second line of defence, not a duplicate: /api/ocr strips a stacked
        # heading before returning rows, but the batch stored in the browser is
        # what actually gets sent here, and can be re-sent after edits. A stray
        # heading row reaching this point is written into the workbook as literal
        # text — "Time Out" where a time belongs — because nothing after this is
        # checking for one.
        while rows and _looks_like_header_row(rows[0]):
            rows.pop(0)
        # Re-map onto the canonical column order, so a file that was edited in the
        # UI (or came back with the columns renamed) still lands correctly. The
        # order to map onto is the client's OWN layout: running a 17-column
        # timesheet through the 11-column DTR order silently drops six columns and
        # shifts the rest, because the workbook then reads them by position.
        mapper, canon, required = _LAYOUT_MAPPERS.get(layout_name, (None, None, None))
        if mapper is not None:
            mapping = mapper(headers)
            if all(s in mapping for s in required):
                rows = [
                    [
                        row[mapping[slot]] if slot in mapping and mapping[slot] < len(row) else ""
                        for slot in canon
                    ]
                    for row in rows
                ]
        payload.append({"name": f.name, "file": f.file, "rows": rows,
                        "flagged_rows": f.flagged_rows})

    if not payload:
        raise HTTPException(400, "No files supplied")

    # When the client ships a timekeeping workbook, fill theirs — their formulas,
    # number formats and Shift Schedule sheet are preserved because the file is
    # edited in place rather than rebuilt. Everyone else gets the generic builder.
    tpl = client_registry.template_for(req.client)
    if tpl:
        layout = client_registry.layout_for(req.client) or {}
        slots = layout.get("slots") or []
        blocks = []
        for f in payload:
            rows = []
            for row in f["rows"]:
                rows.append(_derive_row_slots(
                    {slot: (row[i] if i < len(row) else "")
                     for i, slot in enumerate(slots)}))
            blocks.append({"name": f["name"], "rows": rows})

        # Clients the app calculates: every computed column is written as a value
        # (the template's formulas are replaced), from the same engine run and the
        # same confirmed tags as the review screen.
        engine_notes: list[str] = []
        client_key = (req.client or "").strip().lower()
        if client_key in timekeeping.PROFILES and tpl.get("computed"):
            profile, reviews, confirmed = _engine_reviews(client_key, req.files, req.tags, req.holidays,
                                                          req.adjustments)
            flagged = hints = 0
            for idx, (block, (review, days, _)) in enumerate(zip(blocks, reviews)):
                by_date = {r.date: r for r in review.employee.days}
                read = {d.date: d for d in days}
                counts: dict[int, int] = {}
                for x in days:
                    counts[x.date.year] = counts.get(x.date.year, 0) + 1
                year = max(counts, key=lambda y: (counts[y], y)) if counts else None
                for row in block["rows"]:
                    d = timekeeping.parse_date(row.get("date"), year)
                    if d is not None and d not in by_date and year is not None:
                        # a row printed with a stray year (see timekeeping.align_years)
                        try:
                            aligned = d.replace(year=year)
                        except ValueError:
                            aligned = None
                        if aligned in by_date and any(
                                "date printed with the year" in f for f in by_date[aligned].flags):
                            d = aligned
                            row["date"] = f"{d.month}/{d.day}/{d.year % 100:02d}"
                    res = by_date.get(d) if d else None
                    if res is not None:
                        row["__cells__"] = timekeeping.template_cells(
                            res, tpl["computed"], profile, tpl.get("day_minutes"))
                        # The time columns show the times the app calculated with, so a
                        # repaired day (one punch read, time out from the schedule)
                        # doesn't sit beside hours its own times don't add up to.
                        day = read.get(d)
                        for slots, value in ((("time_in", "first_in"), day and day.time_in),
                                             (("time_out", "final_out"), day and day.time_out)):
                            for s in slots:
                                if day is not None and s in row:
                                    row[s] = "" if value is None else str(value)
                flagged += sum(1 for r in review.employee.days if r.flags)
                hints += sum(1 for dt, s in review.suggestions.items()
                             if not s.sure and dt not in (confirmed.get(idx) or {})
                             and review.tags.get(dt) != s.tag)
            if flagged:
                engine_notes.append(f"{flagged} day(s) are flagged for review; check them in "
                                    "Review timekeeping before payroll.")
            if hints:
                engine_notes.append(f"{hints} tag hint(s) not yet decided in Review timekeeping.")

        path = str(Path(__file__).resolve().parent / "templates" / tpl["file"])
        try:
            content, warnings = xlsx_fill.fill(path, blocks, tpl)
        except Exception as e:
            logger.exception("client workbook fill failed")
            raise HTTPException(500, f"Could not fill the workbook: {e}")
        warnings = engine_notes + warnings
        label = client_registry.get_client(req.client)["label"]
        return Response(
            content=content,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={
                "Content-Disposition":
                    _header_safe(f'attachment; filename="DTR - Timekeeping - {label}.xlsx"'),
                "X-Export-Warnings": _header_safe(" | ".join(warnings))[:900],
            },
        )

    label = client_registry.get_client(req.client)["label"]
    raise HTTPException(
        400,
        f"No timekeeping workbook is configured for {label}. Their extracted rows "
        f"are available as CSV; add their .xlsx to app/templates/ and a TEMPLATES "
        f"entry to enable the Excel export."
    )


# ── Timekeeping engine (computed columns, tag review) ─────────────────────────

class TimekeepingRequest(BaseModel):
    client: str = ""
    files: list[DTRWorkbookFile] = []
    tags: list[TimekeepingTag] = []          # tags the team confirmed on the review screen
    holidays: list[TimekeepingHoliday] = []  # local holidays for this branch
    adjustments: list[TimekeepingAdjustment] = []  # Late/UT minutes entered on the review screen


def _engine_reviews(client_key: str, files: list, tags: list, holidays: list,
                    adjustments: list | None = None):
    """Run a batch through the client's timekeeping rules.

    Returns (profile, [(review, days, skipped_rows) per file], confirmed tags).
    Shared by the review endpoint and the Excel export, so both always calculate
    the same way from the same tags.
    """
    profile = timekeeping.PROFILES[client_key]
    layout_name = (client_registry.get_client(client_key) or {}).get("layout")
    mapper, _, _ = _LAYOUT_MAPPERS.get(layout_name, (None, None, None))
    if mapper is None:
        raise HTTPException(400, f"No input layout for {client_key}.")
    extra: dict[date, tuple[str, str]] = {}
    for h in holidays or []:
        d = timekeeping.parse_date(h.date)
        if d:
            extra[d] = (holidays_cal.REGULAR if h.kind.lower().startswith("reg") else holidays_cal.SPECIAL,
                        h.name or "Local holiday")
    confirmed: dict[int, dict[date, str]] = {}
    for t in tags or []:
        d = timekeeping.parse_date(t.date)
        if d:
            confirmed.setdefault(t.employee, {})[d] = timekeeping.normalize_tag(t.tag)
    adjusted: dict[int, dict[date, int]] = {}
    for a in adjustments or []:
        d = timekeeping.parse_date(a.date)
        if d and a.minutes > 0:
            adjusted.setdefault(a.employee, {})[d] = min(int(a.minutes), 24 * 60)

    reviews = []
    for idx, f in enumerate(files):
        headers = [str(h) for h in (f.headers or [])]
        rows = [[str(c) if c is not None else "" for c in r]
                for r in (f.rows or []) if isinstance(r, list)]
        while rows and _looks_like_header_row(rows[0]):
            rows.pop(0)
        mapping = mapper(headers)
        # Forms print a 2-digit year or none; take the batch's year from any row that has one.
        year = next((d.year for r in rows
                     for d in [timekeeping.parse_date(r[mapping["date"]])]
                     if "date" in mapping and mapping["date"] < len(r) and d), None)
        days: list[timekeeping.DayInput] = []
        skipped = 0
        for r in rows:
            values = {slot: r[i] for slot, i in mapping.items() if i < len(r)}
            day = timekeeping.day_from_row(values, year)
            if day is None:
                skipped += bool("".join(r).strip())
                continue
            days.append(day)
        days = timekeeping.align_years(days)
        review = timekeeping.review_employee(
            f.name or f.file or f"Employee {idx + 1}", days, profile, confirmed.get(idx), extra,
            adjusted.get(idx))
        reviews.append((review, days, skipped))
    return profile, reviews, confirmed


@app.get("/api/timekeeping/clients")
async def timekeeping_clients():
    """Clients whose columns the app calculates itself (the review screen shows for these)."""
    return {"clients": sorted(timekeeping.PROFILES),
            "tags": [*timekeeping.TAG_MATRIX.keys(), "ULH"],
            "columns": timekeeping.COLUMNS}


@app.post("/api/timekeeping")
async def timekeeping_compute(req: TimekeepingRequest):
    """Calculate a batch with the client's rules: per day the suggested and applied
    Remarks tag, the routed hours, Late/UT and review flags; per employee the
    DTR Summary's 26 columns. Nothing here touches Excel.
    """
    client_key = (req.client or "").strip().lower()
    if timekeeping.PROFILES.get(client_key) is None:
        label = (client_registry.get_client(client_key) or {}).get("label") or client_key or "this client"
        raise HTTPException(400, f"The app doesn't calculate {label}'s columns yet.")
    profile, reviews, confirmed = _engine_reviews(client_key, req.files, req.tags, req.holidays,
                                                  req.adjustments)

    out = []
    for idx, (f, (review, days, skipped)) in enumerate(zip(req.files, reviews)):
        emp = review.employee
        by_date = {d.date: d for d in days}
        day_out = []
        for res in emp.days:
            src = by_date.get(res.date)
            sug = review.suggestions.get(res.date)
            day_out.append({
                "date": res.date.isoformat(),
                "weekday": res.date.strftime("%a"),
                "schedule": src.schedule if src else "",
                "time_in": str(src.time_in or "") if src else "",
                "time_out": str(src.time_out or "") if src else "",
                "remarks": src.remarks if src else "",
                "tag": res.tag if res.tag in timekeeping.TAG_MATRIX or res.tag == "ULH" else "",
                "confirmed": res.date in (confirmed.get(idx) or {}),
                "suggestion": ({"tag": sug.tag, "reason": sug.reason, "sure": sug.sure} if sug else None),
                "worked": res.worked,
                "late": res.late, "undertime": res.undertime,
                "extra_ut": res.extra_ut,
                "ot": res.ot, "nd": res.nd, "nd_ot": res.nd_ot,
                "columns": {k: v for k, v in res.columns.items() if v},
                "flags": res.flags,
            })
        flags = list(emp.flags)
        if skipped:
            flags.insert(0, f"{skipped} row(s) without a readable date were left out")
        out.append({"index": idx, "name": emp.name, "file": f.file,
                    "totals": emp.totals, "flags": flags, "days": day_out})
    return {"client": client_key, "profile": profile.name,
            "columns": timekeeping.COLUMNS, "employees": out}


@app.post("/api/detect-orientation")
async def detect_orientation(file: UploadFile = File(...)):
    """Detect image orientation and return the recommended rotation angle.
    
    Uses Gemini to analyze the image and determine if it needs rotation.
    Returns JSON: {"needs_rotation": bool, "rotation_angle": int, "reasoning": str}
    """
    if not _is_image_upload(file.content_type, file.filename):
        raise HTTPException(400, "Only image files are supported for orientation detection")

    contents = await file.read()
    if len(contents) > MAX_FILE_SIZE:
        raise HTTPException(400, "File too large (max 20 MB)")

    try:
        img = Image.open(io.BytesIO(contents))
        # Judge the same pixels the OCR path will see. Without applying the EXIF tag
        # first, the model looks at raw sideways pixels and reports a rotation that
        # /api/ocr already handles — acting on that tips an upright page over.
        img = ImageOps.exif_transpose(img)
        if img.mode != "RGB":
            img = img.convert("RGB")
        return await _detect_orientation_via_gemini(img)
    except UnidentifiedImageError:
        raise HTTPException(400, "Unsupported or corrupted image file")
    except Exception as e:
        raise HTTPException(500, f"Orientation detection failed: {str(e)}")


@app.post("/api/rotate")
async def rotate_image_endpoint(
    file: UploadFile = File(...),
    angle: int = Form(0),
):
    """Rotate an image by the specified angle (0, 90, 180, or 270 degrees clockwise).
    
    Returns the rotated image as PNG.
    """
    if not _is_image_upload(file.content_type, file.filename):
        raise HTTPException(400, "Only image files are supported for rotation")
    
    if angle not in [0, 90, 180, 270]:
        raise HTTPException(400, "Angle must be 0, 90, 180, or 270 degrees")
    
    contents = await file.read()
    if len(contents) > MAX_FILE_SIZE:
        raise HTTPException(400, "File too large (max 20 MB)")
    
    try:
        img = Image.open(io.BytesIO(contents))
        # Bake in the EXIF orientation before rotating. The response is PNG, which
        # carries no orientation tag, so skipping this silently re-rotates an
        # EXIF-tagged upload — even at angle=0, which must be a true no-op.
        img = ImageOps.exif_transpose(img)
        if img.mode != "RGB":
            img = img.convert("RGB")
        rotated = _rotate_image(img, angle)

        # Return as PNG
        buf = io.BytesIO()
        rotated.save(buf, format="PNG")
        buf.seek(0)
        
        return Response(
            content=buf.getvalue(),
            media_type="image/png",
            headers={"X-Rotation-Applied": str(angle)}
        )
    except UnidentifiedImageError:
        raise HTTPException(400, "Unsupported or corrupted image file")
    except Exception as e:
        raise HTTPException(500, f"Rotation failed: {str(e)}")


@app.get("/api/health")
async def health():
    """Health check – confirms credentials are valid."""
    is_ready = False
    if credentials:
        try:
            if not credentials.valid:
                credentials.refresh(AuthRequest())
            is_ready = credentials.valid
        except Exception:
            is_ready = False
    elif GEMINI_API_KEY:
        is_ready = True
        
    return {"status": "ok", "ocr_ready": is_ready, "llm_ready": is_ready, "model_ready": is_ready}
    

# Serve the frontend
_STATIC_DIR = os.getenv("STATIC_DIR", str(Path(__file__).resolve().parent / "static"))
app.mount("/", StaticFiles(directory=_STATIC_DIR, html=True), name="static")
