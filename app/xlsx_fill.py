"""Write extracted DTR rows into a client's own timekeeping workbook.

The workbook is edited in place rather than rebuilt: we open the .xlsx (a zip of
XML), rewrite only the cells in the data columns, and repackage. Everything we do
not touch — the client's formulas, number formats, column widths, Shift Schedule
and Rules Per Column sheets, conditional formatting, print setup — survives
byte-for-byte, because it is never parsed or regenerated.

That is the whole point. Rebuilding a workbook means re-deriving that client's
payroll model, and there are a dozen different models here. Filling theirs means
we never have to know what `NDOT` means for them.

Cells keep their existing style index, so a time written into a cell the template
already formats as `h:mm` still displays as `h:mm`.
"""
from __future__ import annotations

import io
import logging
import re
import shutil
import zipfile
import xml.etree.ElementTree as ET
from html import escape as _xml_escape, unescape as _xml_unescape

import formula_compat

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
RNS = "http://schemas.openxmlformats.org/package/2006/relationships"
DOCREL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
ET.register_namespace("", NS)

log = logging.getLogger(__name__)

# Every namespace prefix the document declares has to survive a round trip.
# ElementTree invents ns0/ns1/... for any namespace it was not told about, and
# that quietly breaks mc:Ignorable, which lists prefixes *by name*: the file ends
# up declaring xmlns:ns1 while Ignorable still says "x14ac xr xr2 xr3". Excel
# rejects that with "We found a problem with some content"; LibreOffice does not
# care, which is why it survived every check here until a client opened one.
_NS_DECL = re.compile(rb'xmlns:([A-Za-z_][\w.\-]*)="([^"]+)"')


def _preserve_prefixes(xml: bytes) -> None:
    """Teach ElementTree this document's own prefixes before it re-serialises it."""
    for prefix, uri in _NS_DECL.findall(xml):
        try:
            ET.register_namespace(prefix.decode("utf-8"), uri.decode("utf-8"))
        except (ValueError, UnicodeDecodeError):
            pass


_ROOT_TAG = re.compile(rb"<([A-Za-z_][\w.\-]*(?::[A-Za-z_][\w.\-]*)?)\b[^>]*>", re.S)


def _restore_ns_decls(original: bytes, produced: bytes) -> bytes:
    """Put back namespace declarations ElementTree dropped as unused.

    Registering a prefix is not enough on its own: ElementTree only writes an
    xmlns for a namespace some element or attribute actually uses, and these are
    declared on the root precisely so mc:Ignorable can name them. Dropping them
    leaves Ignorable pointing at prefixes the file no longer declares.
    """
    original_root = _ROOT_TAG.search(original)
    produced_root = _ROOT_TAG.search(produced)
    if not original_root or not produced_root:
        return produced
    tag = produced_root.group(0)
    present = set(_NS_DECL.findall(tag))
    missing = [b' xmlns:%s="%s"' % (pfx, uri)
               for pfx, uri in _NS_DECL.findall(original_root.group(0))
               if (pfx, uri) not in present]
    if not missing:
        return produced
    cut = produced_root.end() - (2 if tag.endswith(b"/>") else 1)
    return produced[:cut] + b"".join(missing) + produced[cut:]

# Accept a trailing AM/PM. The prompt asks for 24-hour HH:MM and most output
# obeys, but gemini-3.5 quite often appends "AM"/"PM" anyway — and an unparsed
# time is written into the client's workbook as TEXT, which silently breaks every
# formula that does arithmetic on it. Parse it rather than lose it.
_HHMM = re.compile(r"^(\d{1,2}):(\d{2})(?::(\d{2}))?\s*([AaPp])?\.?[Mm]?\.?$")


def to_serial(value) -> float | None:
    """"14:46" → 0.6152…, the Excel fraction-of-a-day their formulas do maths on."""
    m = _HHMM.match(str(value or "").strip())
    if not m:
        return None
    h, mi = int(m.group(1)), int(m.group(2))
    s = int(m.group(3) or 0)
    ampm = (m.group(4) or "").lower()
    if ampm and h <= 12:
        if h < 1 or mi > 59:
            return None
        h = (h % 12) + (12 if ampm == "p" else 0)
    # A meridiem on an hour past 12 is redundant, and often wrong ("20:00AM").
    # The 24-hour digits are the reading; the suffix is noise the model tacked
    # on. Rejecting the cell instead wrote "20:04PM" into the workbook as TEXT,
    # and every formula built on it returned #VALUE!.
    if h > 47 or mi > 59 or s > 59:
        return None
    # No modulo: a shift that runs to 27:00 must stay 27:00. Collapsing it to
    # 3:00 makes "out" land before "in" and every duration formula goes negative.
    # Serials >= 1 are displayed with an elapsed-time format (see _Styles).
    return (((h * 60 + mi) * 60 + s)) / 86400.0



# A1-style reference: optional $ before column and before row. Only the row is
# ever shifted, and only when it is relative — "'Shift Schedule'!$B$2" must stay
# pointing at the schedule table no matter which block it is copied into.
_REF = re.compile(r"(\$?)([A-Z]{1,3})(\$?)(\d+)")


def _col_letters(n: int) -> str:
    out = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        out = chr(65 + rem) + out
    return out


def _shift_formula(text: str, delta: int) -> str:
    """Re-point a formula's relative row references by `delta` rows."""
    return _offset_formula(text, delta, 0)


def _offset_formula(text: str, drow: int, dcol: int) -> str:
    """Re-point a formula's relative references by `drow` rows and `dcol` columns.

    Quoted strings are copied verbatim — sheet names like 'Shift Schedule' and
    literals like "WDO" must not be touched — and a reference is only recognised
    when it neither runs on from a preceding word character nor straight into a
    following one or an opening bracket, so a function name such as LOG10( can
    never be mistaken for column LOG, row 10.
    """
    out: list[str] = []
    i, n = 0, len(text)
    while i < n:
        ch = text[i]
        if ch == '"':
            j = text.find('"', i + 1)
            j = n if j == -1 else j + 1
            out.append(text[i:j])
            i = j
            continue
        if ch == "'":
            j = text.find("'", i + 1)
            j = n if j == -1 else j + 1
            out.append(text[i:j])
            i = j
            continue
        m = _REF.match(text, i)
        prev = text[i - 1] if i else ""
        nxt = text[m.end()] if m and m.end() < n else ""
        # Guard the empty string explicitly: "" in "_." is True in Python, which
        # silently skipped every reference at the very start or end of a formula
        # — a cloned "=H6-C6" kept pointing at the first block's row.
        joined_before = bool(prev) and (prev.isalnum() or prev in "_.")
        joined_after = bool(nxt) and (nxt.isalnum() or nxt in "_(")
        if m and not joined_before and not joined_after:
            col_abs, col, row_abs, row = m.groups()
            new_row = row if row_abs else str(int(row) + drow)
            new_col = col if col_abs or not dcol else _col_letters(col_index(col) + dcol)
            out.append(f"{col_abs}{new_col}{row_abs}{new_row}")
            i = m.end()
            continue
        out.append(ch)
        i += 1
    return "".join(out)


def col_index(letter: str) -> int:
    n = 0
    for ch in letter.upper():
        n = n * 26 + (ord(ch) - 64)
    return n


def _q(tag: str) -> str:
    return f"{{{NS}}}{tag}"


# Excel's built-in time formats. A cell carrying one of these already renders a
# serial as a clock value; anything else (notably numFmtId 0, "General") renders
# it as the raw fraction — which is how "0.3125" reached the client instead of
# "07:30". Most template cells are General: 122 of Alberto's 128 time cells are.
_BUILTIN_TIME_FMTS = {18, 19, 20, 21, 22, 45, 46, 47}


def _is_time_code(code: str) -> bool:
    lowered = str(code or "").lower()
    return ":" in lowered and ("h" in lowered or "m" in lowered)


class _Styles:
    """Hand out cell styles that display a serial as a time.

    Formatting is a property of the style record, not the cell, so we cannot just
    stamp a format onto one cell. Instead each source style gets a lazily-created
    twin: identical fonts, fills, borders and alignment, with the number format
    swapped. Cells that already format as a time are left alone.
    """

    # Elapsed-time format: renders serial 1.125 as "27:00" rather than wrapping
    # to "3:00", which is what the client asked for on overnight shifts.
    ELAPSED = "[h]:mm"
    CLOCK = "h:mm"
    # Excel's own built-in h:mm, which is what the templates' native time cells
    # use. Registering a custom "hh:mm" instead made restyled cells show "09:25"
    # beside the template's "9:19" in the same column.
    CLOCK_BUILTIN = 20

    def __init__(self, xml: bytes):
        self._source = xml
        self.root = ET.fromstring(xml)
        self.cellxfs = self.root.find(_q("cellXfs"))
        self.dirty = False
        self._twins: dict[tuple[int, str], int] = {}
        self._fmt_ids: dict[str, int] = {}
        self._codes: dict[int, str] = {}
        numfmts = self.root.find(_q("numFmts"))
        if numfmts is not None:
            for nf in numfmts.findall(_q("numFmt")):
                code = nf.get("formatCode") or ""
                self._codes[int(nf.get("numFmtId"))] = code
                if code in (self.ELAPSED, self.CLOCK):
                    self._fmt_ids[code] = int(nf.get("numFmtId"))

    def _xfs(self) -> list:
        return self.cellxfs.findall(_q("xf")) if self.cellxfs is not None else []

    def _fmt_of(self, s_index: int) -> int:
        xfs = self._xfs()
        if not 0 <= s_index < len(xfs):
            return 0
        return int(xfs[s_index].get("numFmtId", "0"))

    def formats_as_time(self, s_index: int) -> bool:
        fid = self._fmt_of(s_index)
        return fid in _BUILTIN_TIME_FMTS or _is_time_code(self._codes.get(fid, ""))

    def _numfmt_id(self, code: str) -> int:
        """Id of `code`, registering it in the workbook the first time it is used."""
        if code == self.CLOCK:
            return self.CLOCK_BUILTIN
        if code in self._fmt_ids:
            return self._fmt_ids[code]
        numfmts = self.root.find(_q("numFmts"))
        if numfmts is None:
            numfmts = ET.Element(_q("numFmts"))
            # Schema order is strict: numFmts must be the first child of styleSheet.
            self.root.insert(0, numfmts)
        # Ids below 164 are reserved for Excel's built-ins.
        new_id = max([163, *self._codes] ) + 1
        nf = ET.SubElement(numfmts, _q("numFmt"))
        nf.set("numFmtId", str(new_id))
        nf.set("formatCode", code)
        numfmts.set("count", str(len(numfmts.findall(_q("numFmt")))))
        self._codes[new_id] = code
        self._fmt_ids[code] = new_id
        self.dirty = True
        return new_id

    def time_style(self, s_index: int, elapsed: bool) -> int | None:
        """Style index that shows a time, or None when `s_index` already does.

        `elapsed` forces the [h]:mm variant, needed once a serial reaches 1.0 —
        an ordinary hh:mm would wrap 27:00 back round to 3:00.
        """
        code = self.ELAPSED if elapsed else self.CLOCK
        if not elapsed and self.formats_as_time(s_index):
            return None
        target = self.CLOCK_BUILTIN if code == self.CLOCK else self._fmt_ids.get(code)
        if self._fmt_of(s_index) == target:
            return None
        key = (s_index, code)
        if key in self._twins:
            return self._twins[key]
        xfs = self._xfs()
        if self.cellxfs is None:
            return None
        import copy
        src = xfs[s_index] if 0 <= s_index < len(xfs) else None
        twin = copy.deepcopy(src) if src is not None else ET.Element(_q("xf"))
        twin.set("numFmtId", str(self._numfmt_id(code)))
        twin.set("applyNumberFormat", "1")
        self.cellxfs.append(twin)
        self.cellxfs.set("count", str(len(self._xfs())))
        index = len(self._xfs()) - 1
        self._twins[key] = index
        self.dirty = True
        return index

    def tostring(self) -> bytes:
        return _restore_ns_decls(
            self._source,
            ET.tostring(self.root, encoding="utf-8", xml_declaration=True),
        )


class _Sheet:
    """Minimal row/cell accessor over a parsed worksheet, preserving cell styles."""

    def __init__(self, root: ET.Element, styles: "_Styles | None" = None):
        self.root = root
        self.styles = styles
        self.data = root.find(_q("sheetData"))
        self.rows = {int(r.get("r")): r for r in self.data.findall(_q("row"))}

    def _row(self, n: int) -> ET.Element:
        row = self.rows.get(n)
        if row is None:
            row = ET.SubElement(self.data, _q("row"))
            row.set("r", str(n))
            self.rows[n] = row
            # keep sheetData ordered by row number, or Excel complains
            self.data[:] = sorted(self.data, key=lambda e: int(e.get("r") or 0))
        return row

    def cell(self, letter: str, n: int, create: bool = True) -> ET.Element | None:
        ref = f"{letter}{n}"
        row = self.rows.get(n) if not create else self._row(n)
        if row is None:
            return None
        for c in row.findall(_q("c")):
            if c.get("r") == ref:
                return c
        if not create:
            return None
        c = ET.SubElement(row, _q("c"))
        c.set("r", ref)
        row[:] = sorted(row, key=lambda e: col_index(re.match(r"([A-Z]+)", e.get("r")).group(1)))
        return c

    @staticmethod
    def _clear(c: ET.Element) -> None:
        for tag in ("v", "is", "f"):
            for child in c.findall(_q(tag)):
                c.remove(child)
        if c.get("t"):
            del c.attrib["t"]

    def overwrite(self, letter: str, n: int, value) -> None:
        """Replace a cell's content, formula included, keeping its style.

        Only for columns the app calculates itself (see fill's `__cells__`). Call
        unshare_formulas() first: a shared-formula master overwritten in place
        would leave the rest of its group pointing at nothing.
        """
        c = self.cell(letter, n)
        self._clear(c)
        if value is None or value == "":
            return
        if isinstance(value, (int, float)) and not isinstance(value, bool):
            v = ET.SubElement(c, _q("v"))
            v.text = repr(float(value))
        else:
            c.set("t", "inlineStr")
            is_el = ET.SubElement(c, _q("is"))
            t = ET.SubElement(is_el, _q("t"))
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            t.text = str(value)

    def write(self, letter: str, n: int, value, is_time: bool = False) -> None:
        """Write a value, keeping whatever style the template already had here.

        `is_time` marks the value as a clock serial, so the cell is re-styled to
        display it as one. Without that a General cell — which most template
        cells are — shows the bare fraction: 07:30 arrives as "0.3125".
        """
        c = self.cell(letter, n)
        had_formula = c.find(_q("f")) is not None
        if had_formula:
            return          # never overwrite one of the client's computed cells
        self._clear(c)
        if value is None or value == "":
            return
        if isinstance(value, (int, float)) and not isinstance(value, bool):
            if is_time and self.styles is not None:
                twin = self.styles.time_style(int(c.get("s", "0")), elapsed=value >= 1.0)
                if twin is not None:
                    c.set("s", str(twin))
            v = ET.SubElement(c, _q("v"))
            v.text = repr(value) if isinstance(value, float) else str(value)
        else:
            c.set("t", "inlineStr")
            is_el = ET.SubElement(c, _q("is"))
            t = ET.SubElement(is_el, _q("t"))
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            t.text = str(value)


    def unshare_formulas(self) -> int:
        """Turn every shared formula into an ordinary one; returns cells rewritten.

        The templates store most of each block's formulas as shared: row 2 holds
        the formula with ref="I7:I21" si="1", and rows 3 onward hold only
        <f t="shared" si="1"/>, meaning "the master's formula, shifted to here".
        That membership is by range. Cloning a block copies those dependents to
        rows the master's range does not cover, so they belong to nothing; Excel
        reports the file as damaged, and its repair wipes the whole group to
        #N/A — every row from the 3rd onward, in cloned blocks and the original
        alike. Spelling each formula out removes the dependency entirely.
        """
        masters: dict[str, tuple[str, str]] = {}
        for c in self.data.iter(_q("c")):
            f = c.find(_q("f"))
            if f is not None and f.get("t") == "shared" and f.get("ref") and f.text:
                masters[f.get("si")] = (c.get("r"), f.text)
        rewritten = 0
        for c in self.data.iter(_q("c")):
            f = c.find(_q("f"))
            if f is None or f.get("t") != "shared":
                continue
            if not (f.text or "").strip():
                master = masters.get(f.get("si"))
                if master is None:
                    continue            # already orphaned in the template itself
                mref, mtext = master
                mcol, mrow = re.match(r"([A-Z]+)(\d+)$", mref).groups()
                ccol, crow = re.match(r"([A-Z]+)(\d+)$", c.get("r")).groups()
                f.text = _offset_formula(mtext, int(crow) - int(mrow),
                                         col_index(ccol) - col_index(mcol))
            for attr in ("t", "ref", "si"):
                f.attrib.pop(attr, None)
            rewritten += 1
        return rewritten

    def heal_formula_holes(self, first_block: int, stride: int, per_block: int,
                           blocks: int, data_cols: set[str]) -> list[str]:
        """Restore formulas the client's working copy lost to typing-over.

        The templates are the clients' working files, and a formula cell that
        someone once typed "RD" into and then cleared is simply gone. Bikezilla's
        has 13 FINAL OUT cells missing, sitting exactly on one past period's rest
        days, so any later employee who *works* on those rows gets 0 hours —
        TOTAL reads an empty FINAL OUT. 28 such holes across six templates. A
        column that carries a formula on most rows of a block is a formula
        column; each gap is rebuilt from the block's own first row, shifted down
        to it. Returns the cells rebuilt.
        """
        healed: list[str] = []
        by_ref = {c.get("r"): c for c in self.data.iter(_q("c"))}

        def formula_text(ref: str) -> str | None:
            c = by_ref.get(ref)
            f = c.find(_q("f")) if c is not None else None
            return f.text if (f is not None and f.text) else None

        def has_formula(ref: str) -> bool:
            c = by_ref.get(ref)
            return c is not None and c.find(_q("f")) is not None

        for b in range(blocks):
            head = first_block + stride * b
            rows = [head + 2 + n for n in range(per_block)]
            ref_row = rows[0]
            cols = {re.match(r"([A-Z]+)", r).group(1) for r in by_ref
                    if re.fullmatch(r"[A-Z]+" + str(ref_row), r)} - data_cols
            for col in cols:
                template = formula_text(f"{col}{ref_row}")
                if not template:
                    continue
                present = sum(1 for r in rows if has_formula(f"{col}{r}"))
                if present < per_block // 2:
                    continue                      # mostly typed values: not a formula column
                for r in rows:
                    ref = f"{col}{r}"
                    if has_formula(ref):
                        continue
                    c = self.cell(col, r)
                    self._clear(c)                # drops any stale literal left behind
                    f = ET.SubElement(c, _q("f"))
                    f.text = _offset_formula(template, r - ref_row, 0)
                    by_ref[ref] = c
                    healed.append(ref)
        return healed

    def clone_block(self, src_head: int, dst_head: int, stride: int) -> None:
        """Copy a whole employee block, re-pointing its formulas to the new rows.

        The templates ship with a fixed number of blocks — Miniso has exactly one,
        against 284 employees — so without this the export silently serves a single
        person. Cloning keeps every cell's style and formula intact and only moves
        the row references.
        """
        import copy
        delta = dst_head - src_head
        for offset in range(stride):
            src = self.rows.get(src_head + offset)
            if src is None:
                continue
            new_row_n = dst_head + offset
            dup = copy.deepcopy(src)
            dup.set("r", str(new_row_n))
            for c in dup.findall(_q("c")):
                ref = c.get("r") or ""
                col = re.match(r"([A-Z]+)", ref)
                if col:
                    c.set("r", f"{col.group(1)}{new_row_n}")
                f = c.find(_q("f"))
                if f is not None and f.text:
                    f.text = _shift_formula(f.text, delta)
                    # a shared formula copied out of its range becomes wrong; make
                    # this cell own its formula outright
                    for attr in ("t", "ref", "si"):
                        if attr in f.attrib:
                            del f.attrib[attr]
                if f is not None:
                    # Drop the stale cached result, but ONLY for formula cells. On a
                    # shared-string cell <v> holds the string's index, not a cached
                    # value, so clearing it erased the block's own labels — cloned
                    # Miniso blocks came out missing "NAME:" and "Total".
                    for v in c.findall(_q("v")):
                        c.remove(v)
            # Replace, never append: the sheet may already carry a row here (the
            # template is longer than its one block), and two <row> elements with
            # the same r= is invalid and silently corrupts what Excel computes.
            existing = self.rows.get(new_row_n)
            if existing is not None:
                self.data.remove(existing)
            self.rows[new_row_n] = dup
            self.data.append(dup)
        self.data[:] = sorted(self.data, key=lambda e: int(e.get("r") or 0))

    def text_at(self, letter: str, n: int, shared: list[str] | None = None) -> str:
        """The value shown in a cell, resolving a shared-string index."""
        c = self.cell(letter, n, create=False)
        if c is None:
            return ""
        inline = c.find(_q("is"))
        if inline is not None:
            return "".join(t.text or "" for t in inline.iter(_q("t")))
        v = c.find(_q("v"))
        if v is None or v.text is None:
            return ""
        if c.get("t") == "s" and shared is not None:
            try:
                return shared[int(v.text)]
            except (ValueError, IndexError):
                return ""
        return v.text

    def clear_block(self, head: int, per_block: int, columns: dict, name_col: str) -> bool:
        """Blank an employee block the current batch does not reach.

        The shipped templates are the clients' own working files and still carry
        the people who were in them when they were handed over — 19 named blocks
        in Alberto's. Filling only as many blocks as the batch has leaves those
        behind, so an export of three employees comes back with sixteen strangers
        attached and rows that match no uploaded DTR. Returns True if it cleared
        anything, so the caller can say so.
        """
        found = False
        cells = [(name_col, head)] + [
            (letter, head + 2 + n) for n in range(per_block) for letter in columns
        ]
        for letter, row_n in cells:
            c = self.cell(letter, row_n, create=False)
            if c is None or c.find(_q("f")) is not None:
                continue    # absent, or one of the client's own formulas
            if c.find(_q("v")) is not None or c.find(_q("is")) is not None:
                found = True
            self._clear(c)
        return found

    def drop_cached_values(self, first_row: int, last_row: int) -> None:
        """Remove cached results from formula cells so Excel recomputes on open."""
        for n in range(first_row, last_row + 1):
            row = self.rows.get(n)
            if row is None:
                continue
            for c in row.findall(_q("c")):
                if c.find(_q("f")) is not None:
                    for v in c.findall(_q("v")):
                        c.remove(v)


def _sheet_path(z: zipfile.ZipFile, sheet_name: str) -> str:
    wb = ET.fromstring(z.read("xl/workbook.xml"))
    rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    rid2t = {r.get("Id"): r.get("Target") for r in rels.findall(f"{{{RNS}}}Relationship")}
    for sh in wb.find(_q("sheets")):
        if sh.get("name") == sheet_name:
            tgt = rid2t[sh.get(f"{{{DOCREL}}}id")]
            return tgt if tgt.startswith("xl/") else "xl/" + tgt.lstrip("/")
    raise KeyError(f"sheet {sheet_name!r} not found in template")


def _extend_sheet_ranges(xml: bytes, sheet_name: str, min_end: int, added: int) -> tuple[bytes, int]:
    """Stretch ranges on other sheets that stop at the template's last block.

    Miniso's DTR Summary finds each employee with
    VLOOKUP(<block no.>, Timekeeping!$A$3:$C$162, ...). That range covers the
    eight blocks the file ships with, so once a batch is cloned past eight every
    later employee resolves to #N/A.

    The end row is read from the file rather than recomputed, because templates
    carry a row or two of slack past the last block; shifting whatever is there
    by the number of rows we added preserves that slack. Only ranges that name
    this sheet and already reach its last block are touched.
    """
    text = xml.decode("utf-8")
    hits = 0

    def widen(m: re.Match) -> str:
        nonlocal hits
        end = int(m.group("end"))
        if end < min_end:
            return m.group(0)
        hits += 1
        return m.group(0)[: m.start("end") - m.start(0)] + str(end + added)

    for ref in (f"'{sheet_name}'!", f"{sheet_name}!"):
        if ref not in text:
            continue
        pattern = re.compile(
            re.escape(ref)
            + r"\$?[A-Z]{1,3}\$?\d+:\$?[A-Z]{1,3}\$?(?P<end>\d+)\b"
        )
        text = pattern.sub(widen, text)
    return text.encode("utf-8"), hits


def _downgrade_root(root: ET.Element, where: str, problems: list[str]) -> int:
    """Rewrite LET/SWITCH/XLOOKUP in every formula of a parsed sheet (see formula_compat)."""
    changed = 0
    for c in root.iter(_q("c")):
        f = c.find(_q("f"))
        if f is None or not f.text:
            continue
        new, problem = formula_compat.downgrade_safe(f.text)
        if problem:
            problems.append(f"{where}!{c.get('r')}: {problem}")
        elif new != f.text:
            f.text = new
            changed += 1
    return changed


_F_ELEMENT = re.compile(r"(<f\b[^>]*>)([^<]*)(</f>)")


def _downgrade_xml(xml: bytes, where: str, problems: list[str]) -> tuple[bytes, int]:
    """Same as _downgrade_root, on a sheet we otherwise leave byte-for-byte alone."""
    text = xml.decode("utf-8")
    if "_xlfn." not in text and "_xlpm." not in text:
        return xml, 0
    changed = 0

    def sub(m: re.Match) -> str:
        nonlocal changed
        old = _xml_unescape(m.group(2))
        new, problem = formula_compat.downgrade_safe(old)
        if problem:
            problems.append(f"{where}: {problem}")
            return m.group(0)
        if new == old:
            return m.group(0)
        changed += 1
        return m.group(1) + _xml_escape(new, quote=False) + m.group(3)

    return _F_ELEMENT.sub(sub, text).encode("utf-8"), changed


def _force_recalc(wb_xml: bytes) -> bytes:
    """Make Excel recalculate on open, since we strip the cached values."""
    text = wb_xml.decode("utf-8")
    if "<calcPr" in text:
        text = re.sub(r"<calcPr[^/>]*/>", '<calcPr calcId="0" fullCalcOnLoad="1"/>', text)
        text = re.sub(r"<calcPr([^>]*)>", r'<calcPr\1 fullCalcOnLoad="1">', text, count=1) \
            if "fullCalcOnLoad" not in text else text
    else:
        text = text.replace("</workbook>", '<calcPr calcId="0" fullCalcOnLoad="1"/></workbook>')
    return text.encode("utf-8")


_SAME_SHEET_REF = re.compile(
    r"(?<![A-Za-z0-9_.!'\]])\$?([A-Z]{1,3})\$?(\d+)(?::\$?([A-Z]{1,3})\$?(\d+))?(?![A-Za-z0-9_(])")


# A time cell holding only a "no punch" mark: dashes, dots, underscores, "n/a".
_PLACEHOLDER = re.compile(r"^\s*(?:[-‒-―_.~]+|n/?a)\s*$", re.I)


def _has_punch(row: dict) -> bool:
    """A real clock time in any in/out slot. 0:00 alone is how some forms print an
    empty day (Alberto's computed rows), so it doesn't count when choosing rows."""
    return any((to_serial(v) or 0) > 0 for k, v in row.items()
               if k.endswith("_in") or k.endswith("_out"))


def _fit_rows(rows: list[dict], size: int) -> tuple[list[dict], int]:
    """Pick the `size` consecutive rows that hold the most punched days.

    A block holds one pay period (16 rows), but a full-month DTR form comes back
    with a row for every day 1-31, punched or not. Keeping the first 16 rows
    dropped the period's later days whenever it didn't start on the 1st: a June
    11-25 Miniso form lost days 17-25. Order is kept, so the block still reads
    day by day. Among equally good windows, prefer one that starts on a punched
    day, then the earliest.

    Returns (rows, number of punched days that still didn't fit).
    """
    punched = [_has_punch(r) for r in rows]
    best_start, best_key = 0, None
    for start in range(len(rows) - size + 1):
        key = (sum(punched[start:start + size]), punched[start], -start)
        if best_key is None or key > best_key:
            best_start, best_key = start, key
    kept = rows[best_start:best_start + size]
    return kept, sum(punched) - sum(punched[best_start:best_start + size])


def _columns_read_in_row(root: ET.Element, row_n: int) -> set[str]:
    """Columns this row's own formulas read, with ranges expanded (C6:H6 -> C..H)."""
    read: set[str] = set()
    for c in root.iter(_q("c")):
        m = re.match(r"[A-Z]+(\d+)$", c.get("r") or "")
        f = c.find(_q("f"))
        if not m or int(m.group(1)) != row_n or f is None or not f.text:
            continue
        for a, r1, b, r2 in _SAME_SHEET_REF.findall(f.text):
            if b and int(r1) <= row_n <= int(r2):
                read.update(_col_letters(i) for i in range(col_index(a), col_index(b) + 1))
            elif int(r1) == row_n:
                read.add(a)
    return read


def fill(template_path: str, blocks: list[dict], cfg: dict,
         compat: bool = True) -> tuple[bytes, list[str]]:
    """Fill `blocks` (one per employee) into a copy of the client's workbook.

    Each block is {"name": str, "rows": [{slot: value}, ...]}. `cfg` is the
    TEMPLATES entry describing the sheet, block geometry and column→slot map.
    `compat` rewrites Excel 365-only functions so Excel 2016/2019 can recalculate
    the file (see formula_compat); False keeps the client's formulas verbatim and
    exists for side-by-side testing.
    Returns (xlsx bytes, warnings).
    """
    warnings: list[str] = []
    first_block = cfg["first_block"]
    stride = cfg["stride"]
    per_block = cfg["rows_per_block"]
    columns: dict[str, str] = cfg["columns"]
    name_col = cfg.get("name_col", "C")

    with zipfile.ZipFile(template_path) as z:
        names = z.namelist()
        sheet_path = _sheet_path(z, cfg["sheet"])
        parts = {n: z.read(n) for n in names}

    _preserve_prefixes(parts[sheet_path])
    sheet_source = parts[sheet_path]
    root = ET.fromstring(parts[sheet_path])
    if "xl/styles.xml" in parts:
        _preserve_prefixes(parts["xl/styles.xml"])
    styles = _Styles(parts["xl/styles.xml"]) if "xl/styles.xml" in parts else None
    sheet = _Sheet(root, styles)

    # Grow the sheet to fit the batch. The shipped workbooks carry a fixed number
    # of employee blocks (Miniso: one), so without cloning a 284-person batch would
    # deliver a single person and a warning.
    capacity = int(cfg.get("blocks") or 1)
    if len(blocks) > capacity:
        # Before copying anything: cloned shared-formula dependents would point at
        # a group that does not cover them. See _Sheet.unshare_formulas.
        unshared = sheet.unshare_formulas()
        log.info("unshared %d formula cell(s) before cloning", unshared)
        for n in range(capacity, len(blocks)):
            sheet.clone_block(first_block, first_block + stride * n, stride)
        # Routine, so logged rather than returned: the note popped up in the UI
        # after every large batch and read like an error ("7 blocks-Template").
        log.info("template ships %d block(s); cloned out to %d", capacity, len(blocks))

    # Time columns no formula reads: the only safe home for a word like REST DAY.
    # The client's own Bikezilla sheet leaves MORNING IN blank on a rest day and
    # writes "RD" in MORNING OUT, which is exactly such a column.
    read = _columns_read_in_row(root, first_block + 2)
    note_cols = [letter for letter, slot in columns.items()
                 if (slot.endswith("_in") or slot.endswith("_out")) and letter not in read]

    healed = sheet.heal_formula_holes(first_block, stride, per_block,
                                      max(len(blocks), capacity), set(columns))
    if healed:
        log.info("rebuilt %d formula cell(s) missing from the template: %s",
                 len(healed), ", ".join(healed[:12]))

    # Columns the app calculated itself (the timekeeping engine) arrive per row as
    # "__cells__": {letter: value}. They replace the template's formulas outright,
    # so spell shared formulas out first: overwriting a group's master in place
    # would leave the rest of the group pointing at nothing.
    if any(r.get("__cells__") for b in blocks for r in (b.get("rows") or [])):
        sheet.unshare_formulas()

    for i, block in enumerate(blocks):
        head = first_block + stride * i
        first_row = head + 2
        last_row = head + 1 + per_block

        sheet.write(name_col, head, block.get("name", ""))

        rows = block.get("rows") or []
        if len(rows) > per_block:
            rows, lost = _fit_rows(rows, per_block)
            if lost:
                warnings.append(
                    f"{block.get('name','?')}: {lost} day(s) with punches didn't fit the "
                    f"{per_block}-day block and were left out; check this employee's dates."
                )

        for n in range(per_block):
            r = first_row + n
            values = rows[n] if n < len(rows) else {}
            for letter, slot in columns.items():
                raw = values.get(slot, "")
                serial = to_serial(raw)
                if serial is not None:
                    sheet.write(letter, r, serial, is_time=True)
                elif (slot.endswith("_in") or slot.endswith("_out")) and (
                        re.search(r"\d", str(raw or "")) or _PLACEHOLDER.match(str(raw or ""))):
                    # Digits that didn't parse as a time — "20:5PM", a stray "29" — are
                    # never a valid entry for a time column, and as text they break any
                    # formula doing arithmetic on it. Neither is a bare "-" (how Peak 24
                    # and Tokyo forms mark a day with no punch): written as text it made
                    # the day's hours and the block total #VALUE!. Clear the cell instead;
                    # words such as REST DAY or OFF are left to the rules above and below.
                    sheet.write(letter, r, "")
                else:
                    sheet.write(letter, r, raw)
            note = str(values.get("day_note") or "").strip()
            shown = {str(values.get(slot) or "").strip() for slot in columns.values()}
            if note and note not in shown:
                spot = next((L for L in note_cols if not str(values.get(columns[L]) or "").strip()), None)
                if spot:
                    sheet.write(spot, r, note)
            for letter, value in (values.get("__cells__") or {}).items():
                sheet.overwrite(letter, r, value)

        sheet.drop_cached_values(head, last_row)

    # Every template numbers its employee blocks in the first column of the block's
    # head row. A cloned block inherits the source block's number, so a 30-person
    # Miniso export arrived with thirty blocks all numbered "1" and the sheet's own
    # numbering gone. Only renumber where the template actually keeps a number
    # there, so a client who uses that cell for something else is left alone.
    shared: list[str] = []
    if "xl/sharedStrings.xml" in parts:
        sroot = ET.fromstring(parts["xl/sharedStrings.xml"])
        shared = ["".join(t.text or "" for t in si.iter(_q("t")))
                  for si in sroot.findall(_q("si"))]
    seq_col = cfg.get("seq_col", "A")
    if sheet.text_at(seq_col, first_block, shared).strip().isdigit():
        for i in range(max(len(blocks), capacity)):
            sheet.write(seq_col, first_block + stride * i, i + 1)

    # Blocks past the end of this batch still hold whoever the client had in the
    # file when they sent it. Blank them, or the export mixes real strangers in
    # with the batch and the team reconciles rows that came from no DTR at all.
    stale = 0
    for i in range(len(blocks), capacity):
        head = first_block + stride * i
        if sheet.clear_block(head, per_block, columns, name_col):
            stale += 1
        sheet.drop_cached_values(head, head + 1 + per_block)
    if stale:
        log.info("cleared %d pre-filled employee block(s) from the template", stale)

    # A batch bigger than the file's own block count needs the other sheets'
    # lookup ranges stretched to match, or the employees past the original last
    # block resolve to #N/A in the summary.
    if len(blocks) > capacity:
        # Anything reaching the last shipped block gets shifted by the rows added.
        min_end = first_block + stride * (capacity - 1)
        added = stride * (len(blocks) - capacity)
        widened = 0
        for name in list(parts):
            if not name.startswith("xl/worksheets/") or name == sheet_path:
                continue
            parts[name], hits = _extend_sheet_ranges(
                parts[name], cfg["sheet"], min_end, added)
            widened += hits
        if widened:
            log.info("stretched %d lookup range(s) to reach %d employees", widened, len(blocks))

    # A summary sheet with a fixed number of employee rows (24 Chicken's DTR
    # Summary: 40) looks each row up by employee number. Rows past the batch would
    # look up a block that doesn't exist and show #N/A, which then poisons the
    # TOTAL row, so they are blanked; a batch longer than the summary is reported.
    summary = cfg.get("summary")
    if summary:
        with zipfile.ZipFile(template_path) as z:
            spath = _sheet_path(z, summary["sheet"])
        _preserve_prefixes(parts[spath])
        s_source = parts[spath]
        s_root = ET.fromstring(s_source)
        s_sheet = _Sheet(s_root, styles)
        s_sheet.unshare_formulas()
        last_col = col_index(summary["last_col"])
        room = summary["last_row"] - summary["first_row"] + 1
        for r in range(summary["first_row"] + len(blocks), summary["last_row"] + 1):
            for ci in range(1, last_col + 1):
                s_sheet.overwrite(_col_letters(ci), r, "")
        parts[spath] = _restore_ns_decls(s_source, ET.tostring(s_root, encoding="utf-8", xml_declaration=True))
        if len(blocks) > room:
            warnings.append(f"The {summary['sheet']} has rows for {room} employees; the other "
                            f"{len(blocks) - room} are on the {cfg['sheet']} sheet only.")

    # The export drops every cached result and forces a full recalculation, so
    # each formula has to be one the recipient's Excel can evaluate. DPAM's
    # templates use LET/SWITCH/XLOOKUP; Excel 2016/2019 shows #NAME? for those.
    if compat:
        problems: list[str] = []
        rewritten = _downgrade_root(root, cfg["sheet"], problems)
        for name in list(parts):
            if name.startswith("xl/worksheets/") and name.endswith(".xml") and name != sheet_path:
                parts[name], n = _downgrade_xml(parts[name], name.rsplit("/", 1)[-1], problems)
                rewritten += n
        if rewritten:
            log.info("rewrote %d formula(s) for Excel 2016/2019 compatibility", rewritten)
        for p in problems[:20]:
            log.warning("formula left as-is (needs Excel 2021/365): %s", p)
        if problems:
            warnings.append(
                f"{len(problems)} formula(s) use functions older Excel can't calculate; "
                "those cells may show #NAME? in Excel 2019 or earlier.")

    parts[sheet_path] = _restore_ns_decls(
        sheet_source, ET.tostring(root, encoding="utf-8", xml_declaration=True))
    if styles is not None and styles.dirty:
        parts["xl/styles.xml"] = styles.tostring()
    if "xl/workbook.xml" in parts:
        parts["xl/workbook.xml"] = _force_recalc(parts["xl/workbook.xml"])
    # calcChain describes a formula order that no longer matches; Excel rebuilds it.
    parts.pop("xl/calcChain.xml", None)

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as out:
        for name, data in parts.items():
            if name == "[Content_Types].xml":
                data = data.replace(
                    b'<Override PartName="/xl/calcChain.xml" '
                    b'ContentType="application/vnd.openxmlformats-officedocument.'
                    b'spreadsheetml.calcChain+xml"/>', b"")
            out.writestr(name, data)
    return buf.getvalue(), warnings
