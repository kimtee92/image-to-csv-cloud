"""End-to-end check on a deployed service: scans -> /api/ocr -> /api/timekeeping -> answer key.

Uses real scans and a DPAM answer key, so it runs by hand against a test
environment, never in CI (the data stays in a git-ignored reference*/ folder):
    python3 ci/e2e_check.py --base https://<service>.run.app --client oppa_foods \
        --scans reference8/scans-oct26-nov10 --key reference8/oppa-answer-key-2026-11-23.xlsx \
        --out reference8/e2e

Reports three layers separately, so a reading error is never mistaken for a rule error:
  1. extraction: schedule, time in and time out per day vs the key's day rows
  2. tags: the tag each day was calculated with vs the key's Remarks tag
  3. totals: the 26 DTR Summary columns per employee vs the key
"""
from __future__ import annotations

import argparse
import concurrent.futures as cf
import json
import pathlib
import re
import sys
import urllib.request
import uuid

ROOT = pathlib.Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "app"))
sys.path.insert(0, str(ROOT / "ci"))
import timekeeping as tk  # noqa: E402
from answer_key_check import employees, expected, load_sheets  # noqa: E402


def post_multipart(url: str, fields: dict[str, str], file_path: pathlib.Path) -> dict:
    boundary = uuid.uuid4().hex
    body = bytearray()
    for k, v in fields.items():
        body += f"--{boundary}\r\nContent-Disposition: form-data; name=\"{k}\"\r\n\r\n{v}\r\n".encode()
    body += (f"--{boundary}\r\nContent-Disposition: form-data; name=\"file\"; "
             f"filename=\"{file_path.name}\"\r\nContent-Type: image/jpeg\r\n\r\n").encode()
    body += file_path.read_bytes() + f"\r\n--{boundary}--\r\n".encode()
    req = urllib.request.Request(url, data=bytes(body),
                                 headers={"Content-Type": f"multipart/form-data; boundary={boundary}"})
    with urllib.request.urlopen(req, timeout=300) as r:
        return json.loads(r.read())


def post_json(url: str, payload: dict) -> dict:
    req = urllib.request.Request(url, data=json.dumps(payload).encode(),
                                 headers={"Content-Type": "application/json"})
    with urllib.request.urlopen(req, timeout=300) as r:
        return json.loads(r.read())


def employee_name(ocr: dict, fallback: str) -> str:
    lines = []
    for line in ocr.get("ocr_text_lines") or []:
        lines += re.split(r"\\n|\n", str(line))       # some come back with literal "\n" inside
    for line in lines:
        m = re.match(r"\s*(?:employee(?:'?s)?(?:\s+name)?|name)\s*:\s*(.+)", line, re.I)
        if m and m.group(1).strip():
            return m.group(1).strip()
    return fallback


def key_for(name: str, staff: dict) -> str | None:
    """The key's employee for an extracted name; tolerant of a misread letter ("Annna" for "Anna")."""
    import difflib
    norm = lambda s: re.sub(r"[^a-z ]", "", s.lower()).strip()  # noqa: E731
    keys = {norm(k): k for k in staff}
    hit = difflib.get_close_matches(norm(name), list(keys), n=1, cutoff=0.75)
    return keys[hit[0]] if hit else None


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--base", required=True)
    ap.add_argument("--client", default="oppa_foods")
    ap.add_argument("--scans", required=True)
    ap.add_argument("--key", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--reuse", action="store_true", help="reuse saved OCR results instead of re-reading")
    args = ap.parse_args()
    base, out = args.base.rstrip("/"), pathlib.Path(args.out)
    out.mkdir(parents=True, exist_ok=True)
    scans = sorted(pathlib.Path(args.scans).glob("*.jpg"), key=lambda p: int(re.sub(r"\D", "", p.stem) or 0))

    def read(p: pathlib.Path) -> tuple[pathlib.Path, dict]:
        saved = out / (p.stem + ".json")
        if args.reuse and saved.exists():
            return p, json.loads(saved.read_text("utf-8"))
        data = post_multipart(f"{base}/api/ocr", {"client": args.client}, p)
        saved.write_text(json.dumps(data, indent=1), "utf-8")
        return p, data

    with cf.ThreadPoolExecutor(5) as pool:
        results = list(pool.map(read, scans))

    sheets = load_sheets(args.key)
    staff = employees(sheets["Timekeeping"])
    want, headers = expected(sheets["DTR Summary"])
    rename = dict(zip(headers, tk.COLUMNS))

    files = []
    for p, data in results:
        files.append({"name": employee_name(data, p.stem), "file": p.name,
                      "headers": data.get("headers") or [], "rows": data.get("rows") or []})
        print(f"  read {p.name}: {len(data.get('rows') or [])} rows, layout "
              f"{'in/out' if data.get('is_inout') else 'DTR' if data.get('is_dtr') else 'other'}, "
              f"headers {len(data.get('headers') or [])}")
    calc = post_json(f"{base}/api/timekeeping", {"client": args.client, "files": files})
    (out / "timekeeping.json").write_text(json.dumps(calc, indent=1), "utf-8")

    cells = cell_ok = tag_ok = tag_total = 0
    misses, tag_misses, total_diffs = [], [], []
    for emp in calc["employees"]:
        kname = key_for(emp["name"], staff)
        if kname is None:
            print(f"  ??  {emp['name']}: no matching employee in the key")
            continue
        kdays = {d.date.isoformat(): d for d in staff[kname]}
        read_dates = {d["date"] for d in emp["days"]}
        for iso, k in sorted(kdays.items()):
            if iso not in read_dates:
                cells += 3
                misses.append(f"{kname[:22]:22} {iso} NOT READ (key: {k.schedule!r} "
                              f"{k.time_in or ''} {k.time_out or ''})")
        for d in emp["days"]:
            k = kdays.get(d["date"])
            if k is None:
                continue
            for field, got, exp in (("schedule", d["schedule"], k.schedule),
                                    ("in", d["time_in"], k.time_in), ("out", d["time_out"], k.time_out)):
                if field == "schedule":
                    a, b = re.sub(r"\s", "", str(got)).lower(), re.sub(r"\s", "", str(exp)).lower()
                    same = a == b or (a in ("", "-") and b in ("", "-"))
                else:
                    a, b = tk.parse_time(got), tk.parse_time(exp)
                    same = a == b or (a is not None and b is not None and a % 1440 == b % 1440)
                cells += 1
                cell_ok += same
                if not same:
                    misses.append(f"{kname[:22]:22} {d['date']} {field}: read {got!r}, key {exp!r}")
            ktag = tk.normalize_tag(k.tag)
            if tk.parse_time(k.time_in) is not None:
                tag_total += 1
                if d["tag"] == ktag:
                    tag_ok += 1
                else:
                    s = d.get("suggestion") or {}
                    tag_misses.append(f"{kname[:22]:22} {d['date']} tag {d['tag'] or '(blank)'} vs key "
                                      f"{ktag or '(blank)'}; marks {d['remarks']!r}; "
                                      f"suggestion {s.get('tag')!r} {'' if s.get('sure', True) else '(hint)'}")
        exp_row = want.get(kname, {})
        for key_col, value in exp_row.items():
            col = rename[key_col]
            got = emp["totals"].get(col, 0)
            if abs(got - value) > 0.01:
                total_diffs.append(f"{kname[:22]:22} {col}: key {value:g}, app {got:g}")

    print(f"\n1. extraction: {cell_ok}/{cells} schedule/in/out cells match the key")
    for m in misses:
        print("     " + m)
    print(f"2. tags: {tag_ok}/{tag_total} worked days carry the key's tag")
    for m in tag_misses:
        print("     " + m)
    print(f"3. totals: {len(total_diffs)} of {len(want) * len(tk.COLUMNS)} values differ from the key")
    for m in total_diffs:
        print("     " + m)
    return 0


if __name__ == "__main__":
    sys.exit(main())
