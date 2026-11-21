"""Regression tests for xlsx_fill, run by Cloud Build (preflight step).

Each test is a bug DPAM actually hit. Plain asserts, no test framework needed:
    python3 ci/test_fill.py
"""
from __future__ import annotations

import io
import pathlib
import sys
import zipfile
import xml.etree.ElementTree as ET

ROOT = pathlib.Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "app"))
from clients import TEMPLATES  # noqa: E402
import xlsx_fill  # noqa: E402

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
RID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"


def cells(data: bytes, sheet: str) -> dict[str, str]:
    """{ref: value} for the non-formula cells of one sheet (shared strings resolved)."""
    z = zipfile.ZipFile(io.BytesIO(data))
    wb = ET.fromstring(z.read("xl/workbook.xml"))
    rels = {r.get("Id"): r.get("Target") for r in ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))}
    target = next(rels[s.get(RID)] for s in wb.find(NS + "sheets") if s.get("name") == sheet)
    part = target.lstrip("/") if target.lstrip("/").startswith("xl/") else "xl/" + target.lstrip("/")
    shared = []
    if "xl/sharedStrings.xml" in z.namelist():
        shared = ["".join(t.text or "" for t in si.iter(NS + "t"))
                  for si in ET.fromstring(z.read("xl/sharedStrings.xml")).findall(NS + "si")]
    out = {}
    for c in ET.fromstring(z.read(part)).iter(NS + "c"):
        if c.find(NS + "f") is not None:
            continue
        v = c.find(NS + "v")
        if c.get("t") == "s" and v is not None:
            out[c.get("r")] = shared[int(v.text)]
        elif c.get("t") == "inlineStr":
            out[c.get("r")] = "".join(t.text or "" for t in c.iter(NS + "t"))
        elif v is not None and v.text is not None:
            out[c.get("r")] = v.text
    return out


def fill(tpl: str, rows: list[dict]) -> tuple[dict[str, str], list[str], dict]:
    cfg = TEMPLATES[tpl]
    data, warnings = xlsx_fill.fill(str(ROOT / "app" / "templates" / cfg["file"]),
                                    [{"name": "TEST", "rows": rows}], cfg)
    return cells(data, cfg["sheet"]), warnings, cfg


def minutes(serial: str | None) -> int | None:
    return None if serial in (None, "") else round(float(serial) * 1440)


def test_full_month_form_keeps_the_punched_days():
    """Miniso, Nov 2026: a June 11-25 DTR has rows for days 1-31; the export kept
    rows 1-16 and dropped the employee's days 17-25."""
    rows = []
    for day in range(1, 32):
        punched = 11 <= day <= 25
        rows.append({"date": f"6/{day}/2026",
                     "s1_in": "08:00" if punched else "", "s1_out": "12:00" if punched else "",
                     "s2_in": "13:00" if punched else "",
                     "s2_out": f"17:{day:02d}" if punched else ""})
    got, warnings, cfg = fill("miniso", rows)
    first = cfg["first_block"] + 2
    out_col = next(L for L, s in cfg["columns"].items() if s == "s2_out")
    outs = [minutes(got.get(f"{out_col}{first + k}")) for k in range(cfg["rows_per_block"])]
    kept_days = sorted(m - 17 * 60 for m in outs if m is not None)
    assert kept_days == list(range(11, 26)), f"punched days kept: {kept_days}"
    assert not any("didn't fit" in w for w in warnings), warnings


def test_too_many_punched_days_warns():
    """More punched days than the block holds must say so, not drop them silently."""
    rows = [{"date": f"6/{d}/2026", "s1_in": "08:00", "s2_out": "17:00"} for d in range(1, 21)]
    _, warnings, _ = fill("miniso", rows)
    assert any("4 day(s) with punches didn't fit" in w for w in warnings), warnings


def test_dash_in_time_cell_is_blank():
    """Peak 24 / Tokyo, Nov 2026: '-' for a no-punch day was written as text, and
    the day's hours and the block total came out #VALUE!."""
    rows = [{"date": "7/11/26", "s1_in": "-", "final_out": "-"},
            {"date": "7/12/26", "s1_in": "–", "final_out": "n/a"},
            {"date": "7/13/26", "s1_in": "12:00", "final_out": "22:00"}]
    got, _, cfg = fill("peak24", rows)
    first = cfg["first_block"] + 2
    for letter, slot in cfg["columns"].items():
        if slot.endswith("_in") or slot.endswith("_out"):
            for k in (0, 1):
                v = got.get(f"{letter}{first + k}", "")
                assert v == "", f"{letter}{first + k} = {v!r}, expected blank"
    in_col = next(L for L, s in cfg["columns"].items() if s == "s1_in")
    assert minutes(got.get(f"{in_col}{first + 2}")) == 12 * 60


def test_short_month_unchanged():
    """A block with no more rows than it holds is written as-is, in order."""
    rows = [{"date": f"7/{d}/26", "s1_in": "09:00", "final_out": "18:00"} for d in range(11, 27)]
    got, warnings, cfg = fill("peak24", rows)
    date_col = next(L for L, s in cfg["columns"].items() if s == "date")
    first = cfg["first_block"] + 2
    assert got.get(f"{date_col}{first}") == "7/11/26", got.get(f"{date_col}{first}")
    assert not warnings, warnings


if __name__ == "__main__":
    tests = [v for k, v in sorted(globals().items()) if k.startswith("test_")]
    failed = 0
    for t in tests:
        try:
            t()
            print(f"  ok    {t.__name__}")
        except AssertionError as exc:
            failed += 1
            print(f"  FAIL  {t.__name__}: {exc}")
    print(f"{len(tests) - failed}/{len(tests)} passed")
    sys.exit(1 if failed else 0)
