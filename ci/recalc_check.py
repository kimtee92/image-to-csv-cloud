"""Recalculate every client export in a spreadsheet engine and check the results.

Runs in Cloud Build (see cloudbuild.yaml, step `recalc`) on every push, so a
broken export is caught without anyone opening Excel:

  1. Fill each client template with a fixed set of test shifts, exactly as the
     app's /api/dtr-workbook does (xlsx_fill.fill).
  2. Have LibreOffice (headless) fully recalculate each workbook.
  3. Fail if
       - any cell anywhere shows #NAME? (a function the engine doesn't know; the
         Debian LibreOffice, like Excel 2016/2019, has no LET or XLOOKUP), or
       - any formula on the filled rows of the Timekeeping sheet is an error, or
       - a value on those rows differs from what Excel 365 computed for the same
         export (ci/golden/<template>.json, recorded on a PC with Excel).

Modes:
  python3 ci/recalc_check.py export OUTDIR   write the test exports + manifest.json
  python3 ci/recalc_check.py check           export, recalc with soffice, compare

When a change is meant to alter computed values (new template, new rules),
re-record the golden files with ci/record_golden.ps1 on a PC with Excel.
"""
from __future__ import annotations

import json
import os
import pathlib
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
import xml.etree.ElementTree as ET

ROOT = pathlib.Path(__file__).resolve().parent.parent
APP = ROOT / "app"
GOLDEN = ROOT / "ci" / "golden"
sys.path.insert(0, str(APP))

from clients import TEMPLATES  # noqa: E402
import xlsx_fill  # noqa: E402

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
RID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
BLOCKS = 3

# (time in, time out): day shifts, early starts, short days, nights past midnight,
# and a blank day. Each block starts at a different offset so they all differ.
SHIFTS = [
    ("09:00", "18:00"), ("10:15", "20:30"), ("11:45", "22:10"), ("12:30", "23:00"),
    ("13:10", "23:40"), ("14:00", "02:00"), ("22:00", "06:00"), ("06:10", "15:05"),
    ("07:45", "16:50"), ("09:20", "15:00"), ("05:00", "14:00"), ("", ""),
    ("12:00", "22:00"), ("15:30", "01:30"), ("10:00", "19:00"), ("08:05", "17:40"),
]


def _day(i: int, offset: int) -> dict:
    tin, tout = SHIFTS[(i + offset) % len(SHIFTS)]
    if not tin:
        return {}
    return {"s1_in": tin, "s1_out": "12:00", "s2_in": "13:00", "s2_out": tout,
            "final_out": tout, "date": f"7/{11 + i}/26"}


def export(outdir: pathlib.Path) -> dict:
    """Write one test export per template; return {template: manifest entry}."""
    outdir.mkdir(parents=True, exist_ok=True)
    manifest = {}
    for tpl, cfg in sorted(TEMPLATES.items()):
        per = cfg["rows_per_block"]
        blocks = [{"name": f"TEST {b + 1}", "rows": [_day(i, b * 5) for i in range(per)]}
                  for b in range(BLOCKS)]
        data, _ = xlsx_fill.fill(str(APP / "templates" / cfg["file"]), blocks, cfg)
        path = outdir / f"{tpl}.xlsx"
        path.write_bytes(data)
        first = cfg["first_block"]
        last = first + cfg["stride"] * BLOCKS - 1
        # Formula cells come from the file, not from Excel: Excel won't list
        # formula cells on a protected sheet (Tokyo's is), so asking it would
        # silently record nothing for that client.
        with zipfile.ZipFile(path) as z:
            part = _sheets(z)[cfg["sheet"]]
            root = ET.fromstring(z.read(part))
        cells = [c.get("r") for c in root.iter(NS + "c")
                 if c.find(NS + "f") is not None and first <= _row(c.get("r")) <= last]
        manifest[tpl] = {"sheet": cfg["sheet"], "first_row": first, "last_row": last,
                         "cells": cells}
    (outdir / "manifest.json").write_text(json.dumps(manifest, indent=1))
    return manifest


# ── reading a recalculated workbook ──────────────────────────────────────────

def _sheets(z: zipfile.ZipFile) -> dict[str, str]:
    wb = ET.fromstring(z.read("xl/workbook.xml"))
    rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    targets = {r.get("Id"): r.get("Target") for r in rels}
    out = {}
    for sh in wb.find(NS + "sheets"):
        t = targets[sh.get(RID)].lstrip("/")
        out[sh.get("name")] = t if t.startswith("xl/") else "xl/" + t
    return out


def read_values(path: pathlib.Path) -> dict[str, dict[str, object]]:
    """{sheet: {cell: value}} for every formula cell, errors as '#NAME?' etc."""
    z = zipfile.ZipFile(path)
    shared = []
    if "xl/sharedStrings.xml" in z.namelist():
        for si in ET.fromstring(z.read("xl/sharedStrings.xml")).findall(NS + "si"):
            shared.append("".join(t.text or "" for t in si.iter(NS + "t")))
    result = {}
    for name, part in _sheets(z).items():
        cells = {}
        for c in ET.fromstring(z.read(part)).iter(NS + "c"):
            if c.find(NS + "f") is None:
                continue
            v = c.find(NS + "v")
            text = v.text if v is not None and v.text is not None else ""
            t = c.get("t")
            if t == "e":
                val: object = text or "#ERR"
            elif t == "s":
                val = shared[int(text)] if text else ""
            elif t in ("str", "inlineStr"):
                val = text
            elif t == "b":
                val = text == "1"
            else:
                val = float(text) if text else ""
            cells[c.get("r")] = val
        result[name] = cells
    return result


def _row(cell: str) -> int:
    return int(re.sub(r"[A-Z]+", "", cell))


def _same(a: object, b: object) -> bool:
    if isinstance(a, bool) or isinstance(b, bool):
        return a == b
    if isinstance(a, (int, float)) and isinstance(b, (int, float)):
        return abs(a - b) < 1e-6
    # Excel reports "" for an empty-string result; an engine may store 0 or nothing.
    return str(a) == str(b)


def recalc(src: pathlib.Path, dst: pathlib.Path) -> None:
    """Recalculate every .xlsx in src with LibreOffice; write results to dst."""
    profile = pathlib.Path(tempfile.mkdtemp(prefix="lo-profile-"))
    user = profile / "user"
    user.mkdir(parents=True)
    # Always recalculate on load: our exports carry no cached results.
    (user / "registrymodifications.xcu").write_text(
        '<?xml version="1.0" encoding="UTF-8"?>\n'
        '<oor:items xmlns:oor="http://openoffice.org/2001/registry" '
        'xmlns:xs="http://www.w3.org/2001/XMLSchema" '
        'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">\n'
        '<item oor:path="/org.openoffice.Office.Calc/Formula/Load">'
        '<prop oor:name="OOXMLRecalcMode" oor:op="fuse"><value>0</value></prop></item>\n'
        '<item oor:path="/org.openoffice.Office.Calc/Formula/Load">'
        '<prop oor:name="ODFRecalcMode" oor:op="fuse"><value>0</value></prop></item>\n'
        '</oor:items>\n')
    dst.mkdir(parents=True, exist_ok=True)
    files = sorted(str(p) for p in src.glob("*.xlsx"))
    cmd = ["soffice", f"-env:UserInstallation={profile.as_uri()}", "--headless",
           "--calc", "--convert-to", "xlsx", "--outdir", str(dst), *files]
    subprocess.run(cmd, check=True, timeout=900)
    shutil.rmtree(profile, ignore_errors=True)


def check() -> int:
    work = pathlib.Path(tempfile.mkdtemp(prefix="recalc-"))
    manifest = export(work / "exports")
    recalc(work / "exports", work / "recalc")
    failures: list[str] = []
    for tpl, m in manifest.items():
        out = work / "recalc" / f"{tpl}.xlsx"
        if not out.exists():
            failures.append(f"{tpl}: LibreOffice produced no output")
            continue
        values = read_values(out)
        # 1. Unknown functions anywhere.
        for sheet, cells in values.items():
            names = [c for c, v in cells.items() if v == "#NAME?"]
            if names:
                failures.append(f"{tpl}/{sheet}: #NAME? in {len(names)} cell(s), e.g. {names[:5]}")
        rows = range(m["first_row"], m["last_row"] + 1)
        main = {c: v for c, v in values.get(m["sheet"], {}).items() if _row(c) in rows}
        # 2. Errors on the filled rows.
        errs = [c for c, v in main.items() if isinstance(v, str) and v.startswith("#")]
        if errs:
            failures.append(f"{tpl}/{m['sheet']}: {len(errs)} error cell(s) on filled rows, "
                            f"e.g. {[(c, main[c]) for c in errs[:5]]}")
        # 3. Same values as Excel 365 computed.
        gpath = GOLDEN / f"{tpl}.json"
        if not gpath.exists():
            failures.append(f"{tpl}: no golden file {gpath.name} (run ci/record_golden.ps1)")
            continue
        golden = json.loads(gpath.read_text())
        # Every formula cell on the filled rows must have an Excel value, or an
        # empty/stale golden file would pass by comparing nothing.
        uncovered = sorted(set(m["cells"]) - set(golden), key=lambda c: (_row(c), len(c), c))
        if uncovered:
            failures.append(f"{tpl}: {len(uncovered)} formula cell(s) have no Excel 365 value "
                            f"in {gpath.name}, e.g. {uncovered[:5]} (re-run ci/record_golden.ps1)")
        diffs = [(c, g, main.get(c, "<missing>")) for c, g in golden.items()
                 if not _same(g, main.get(c, ""))]
        if diffs:
            failures.append(f"{tpl}/{m['sheet']}: {len(diffs)} value(s) differ from Excel 365, "
                            f"e.g. {diffs[:5]}")
        ok = not diffs and not errs and not uncovered
        print(f"  {tpl:13} {len(golden):5} formula cells checked against Excel 365: "
              f"{'OK' if ok else 'FAIL'}")
    if failures:
        print("\nrecalc check FAILED:", file=sys.stderr)
        for f in failures:
            print("  " + f, file=sys.stderr)
        return 1
    print("recalc check OK")
    return 0


if __name__ == "__main__":
    if len(sys.argv) >= 3 and sys.argv[1] == "export":
        export(pathlib.Path(sys.argv[2]))
    elif len(sys.argv) >= 2 and sys.argv[1] == "check":
        sys.exit(check())
    else:
        print(__doc__)
        sys.exit(2)
