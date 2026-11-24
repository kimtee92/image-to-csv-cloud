"""Check the timekeeping engine against a DPAM answer key (a completed client workbook).

The key holds real employee data, so it is never committed: keep it in a
git-ignored reference*/ folder and pass its path.
    python3 ci/answer_key_check.py reference8/oppa-answer-key-2026-11-23.xlsx [--client oppa_foods] [--days]

Reads the key's Timekeeping sheet (schedule, times and Remarks tag per day) and
its DTR Summary (the expected 26 columns per employee), runs every employee
through app/timekeeping.py, and prints each difference. Known differences that
Dom still has to confirm (see load_known) are reported, not failed.
"""
from __future__ import annotations

import argparse
import pathlib
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from datetime import date

ROOT = pathlib.Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "app"))
import timekeeping as tk  # noqa: E402

NS = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
RID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"

def load_known(key_path: str) -> dict[tuple[str, str], str]:
    """Known differences Dom still has to confirm, kept beside the key (never in git,
    since they name employees): <key>.known.json = [{"employee": "<name prefix>",
    "column": "NDF", "why": "..."}]."""
    side = pathlib.Path(key_path + ".known.json")
    if not side.exists():
        return {}
    import json
    return {(k["employee"], k["column"]): k["why"] for k in json.loads(side.read_text("utf-8"))}


def load_sheets(path: str) -> dict[str, dict[tuple[str, int], object]]:
    z = zipfile.ZipFile(path)
    shared = []
    if "xl/sharedStrings.xml" in z.namelist():
        shared = ["".join(t.text or "" for t in si.iter(NS + "t"))
                  for si in ET.fromstring(z.read("xl/sharedStrings.xml")).findall(NS + "si")]
    wb = ET.fromstring(z.read("xl/workbook.xml"))
    rels = {r.get("Id"): r.get("Target") for r in ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))}
    out = {}
    for s in wb.find(NS + "sheets"):
        target = rels[s.get(RID)].lstrip("/")
        part = target if target.startswith("xl/") else "xl/" + target
        grid = {}
        for c in ET.fromstring(z.read(part)).iter(NS + "c"):
            v = c.find(NS + "v")
            if v is None or v.text is None:
                continue
            val: object = shared[int(v.text)] if c.get("t") == "s" else v.text
            if c.get("t") not in ("s", "str", "e", "b"):
                val = float(v.text)
            m = re.match(r"([A-Z]+)(\d+)$", c.get("r"))
            grid[(m.group(1), int(m.group(2)))] = val
        out[s.get("name")] = grid
    return out


def parse_date(text) -> date | None:
    m = re.match(r"(\d{1,2})/(\d{1,2})/(\d{2,4})$", str(text or "").strip())
    if not m:
        return None
    y = int(m.group(3))
    return date(y + 2000 if y < 100 else y, int(m.group(1)), int(m.group(2)))


def employees(grid) -> dict[str, list[tk.DayInput]]:
    """Blocks of the Timekeeping sheet: 'NAME:' row, header row, then day rows."""
    rows = sorted({r for (_, r) in grid})
    out: dict[str, list[tk.DayInput]] = {}
    name = None
    for r in rows:
        if grid.get(("B", r)) == "NAME:":
            name = str(grid.get(("C", r)) or "").strip()
            out[name] = []
            continue
        d = parse_date(grid.get(("B", r)))
        if name is None or d is None:
            continue
        ins = [grid.get((c, r)) for c in ("D", "F", "H")]
        outs = [grid.get((c, r)) for c in ("E", "G", "I")]
        t_in = next((v for v in ins if isinstance(v, float)), None)
        t_out = next((v for v in reversed(outs) if isinstance(v, float)), None)
        out[name].append(tk.DayInput(d, str(grid.get(("A", r)) or ""), t_in, t_out,
                                     str(grid.get(("AQ", r)) or "")))
    return out


def expected(grid) -> tuple[dict[str, dict[str, float]], list[str]]:
    header_row = next(r for (c, r), v in grid.items() if c == "B" and v == "EMPLOYEE NAME")
    cols = []
    for i in range(3, 3 + len(tk.COLUMNS)):
        letter = chr(64 + i) if i <= 26 else "A" + chr(64 + i - 26)
        cols.append((letter, str(grid.get((letter, header_row)))))
    result = {}
    r = header_row + 1
    while grid.get(("B", r)) not in (None, "TOTAL"):
        result[str(grid[("B", r)]).strip()] = {h: float(grid.get((L, r)) or 0) for L, h in cols}
        r += 1
    return result, [h for _, h in cols]


def expected_from_blocks(grid) -> tuple[dict[str, dict[str, float]], list[str]]:
    """Per-employee totals from the Timekeeping sheet's block rows ("NAME:" in B).

    For a key whose DTR Summary came back #N/A (a lookup that didn't find the
    names); the block's own total row holds the same 26 figures in O..AN.
    """
    letters = []
    for i in range(15, 15 + len(tk.COLUMNS)):          # O .. AN
        letters.append(chr(64 + i) if i <= 26 else "A" + chr(64 + i - 26))
    result = {}
    for (c, r), v in grid.items():
        if c == "B" and str(v).strip().upper() == "NAME:" and str(grid.get(("C", r)) or "").strip():
            name = str(grid[("C", r)]).split("\\n")[0].split("\n")[0].strip()
            vals = {}
            for L, col in zip(letters, tk.COLUMNS):
                try:
                    vals[col] = float(grid.get((L, r)) or 0)
                except (TypeError, ValueError):
                    vals[col] = float("nan")
            result[name] = vals
    return result, list(tk.COLUMNS)


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("key")
    ap.add_argument("--client", default="oppa_foods")
    ap.add_argument("--days", action="store_true", help="print each employee's flags")
    ap.add_argument("--suggest", action="store_true",
                    help="ignore the key's Remarks tags and compare the engine's own suggestions with them")
    args = ap.parse_args()
    if args.suggest:
        return suggest_report(args)
    profile = tk.PROFILES[args.client]
    known_diffs = load_known(args.key)
    sheets = load_sheets(args.key)
    staff = employees(sheets["Timekeeping"])
    want, headers = expected(sheets["DTR Summary"])
    # The key's headers vs the engine's column names (the key says "NDF(10PM-6AM)" and "UT").
    rename = dict(zip(headers, tk.COLUMNS))

    failures = known = checked = 0
    for name, days in staff.items():
        res = tk.calculate_employee(name, days, profile)
        exp = want.get(name)
        if exp is None:
            print(f"  ??  {name}: not in the DTR Summary")
            failures += 1
            continue
        diffs = []
        for key_col, value in exp.items():
            col = rename[key_col]
            checked += 1
            got = res.totals[col]
            if abs(got - value) > 0.01:
                tag = next((why for (who, c), why in known_diffs.items()
                            if name.startswith(who) and c == col), None)
                diffs.append((col, value, got, tag))
        hard = [d for d in diffs if d[3] is None]
        failures += len(hard)
        known += len(diffs) - len(hard)
        status = "OK" if not diffs else ("KNOWN" if not hard else "DIFF")
        print(f"  {status:5} {name[:34]:34} DAYS {res.totals['DAYS']:4.0f}  "
              f"flags {len(res.flags)}")
        for col, value, got, tag in diffs:
            print(f"          {col}: key {value:g}, engine {got:g}" + (f"  (known: {tag})" if tag else ""))
        if args.days:
            for f in res.flags:
                print(f"          flag {f}")
    print(f"{len(staff)} employees, {checked} values checked: {failures} unexplained, {known} known")
    return 1 if failures else 0


def suggest_report(args) -> int:
    """How many of the key's tags the suggestions recover on their own (no scan marks here)."""
    profile = tk.PROFILES[args.client]
    staff = employees(load_sheets(args.key)["Timekeeping"])
    right = wrong = 0
    for name, days in staff.items():
        stripped = [tk.DayInput(d.date, d.schedule, d.time_in, d.time_out) for d in days]
        r = tk.review_employee(name, stripped, profile)
        for d in days:
            want = tk.normalize_tag(d.tag)
            got = r.tags[d.date]
            if not (tk.parse_time(d.time_in) is not None and tk.parse_time(d.time_out) is not None):
                continue
            if want == got:
                right += bool(want)
                continue
            wrong += 1
            s = r.suggestions.get(d.date)
            hint = f"  (hint: {s.tag}, {s.reason})" if s and not s.sure else ""
            print(f"  {name[:28]:28} {d.date}  key {want or '(blank)':6} suggested {got or '(blank)':6}{hint}")
    print(f"tagged days recovered: {right}; days that differ: {wrong}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
