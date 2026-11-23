"""Timekeeping rules engine: DTR days in, the DTR Summary's 26 columns out.

Phase 1 of the "compute in the app" plan (Claude Doc: 24 Chicken timekeeping:
computation build plan). Today the export fills the client's Excel template and
leaves the arithmetic to its formulas, which break on older Excel and disagree
with the client's own rules in places. This module does that arithmetic itself,
as plain Python with no Excel involved, so every rule can be unit tested and the
whole cut-off checked against DPAM's answer key.

Rules (24 Chicken, confirmed by Dom on 23 Nov 2026; his answers win over the
"Timekeeping Rules" sheet where they differ):

  * Schedule comes from the DTR's schedule column: "12-10" = 12:00-22:00. Start
    hours 7-12 are AM/noon, 1-6 are PM; a shift is 10 hrs (8 work + 2 break).
  * RD, rta, med.cert, absent, back out, "-" and blank earn no day credit.
    Time entries on an RD day are suggested as WDO.
  * DAYS counts every worked day (WDO and holidays included). The Remarks tag
    routes the day to its own day column too; OT and night hours go ONLY to the
    tag's columns (Remarks Tag Matrix).
  * Late = minutes after the scheduled start. Undertime = minutes short of the
    10-hr shift, counted from the later of time in and scheduled start. Both are
    summed into one Late/UT figure (Dom: 10-8, in 11:00, out 20:00 = 120 mins).
  * Early in is not OT. OT is time after the scheduled end: from 0.5 hr, in 0.5
    hr steps, rounded down. OT inside 10:00 PM - 6:00 AM is NDF OT instead.
  * NDF is night time inside the scheduled shift. NDF and NDF OT keep the rules
    sheet's 1-hr minimum (then 0.5-hr steps) until Dom says otherwise.
  * ULH: a legal holiday not worked earns 1 day if the last scheduled workday
    before it (rest days skipped) was worked.

Everything is in minutes, measured from midnight of the work date; times past
midnight run on (01:30 the next morning = 25:30 = 1530).
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import date, timedelta

# ── the 26 DTR Summary columns, in the client's order ───────────────────────
# Day columns count days; OT/NDF columns count hours; UT counts minutes.
COLUMNS = [
    "DAYS", "REG. OT", "NDF", "NDF OT",
    "WDO", "WDO OT", "WDO NDF", "WDO NDF OT",
    "SPL. HOL.", "SPL. HOL. OT", "SPL. HOL. NDF", "SPL. HOL. NDF OT",
    "SPL. HOL. / WDO", "SPL. HOL. / WDO OT", "SPL. HOL. / WDO NDF", "SPL. HOL. / WDO NDF OT",
    "LEG. HOL.", "LEG. HOL. OT", "LEG. HOL. NDF", "LEG. HOL. NDF OT",
    "LEG. HOL. / WDO", "LEG. HOL. / WDO OT", "LEG. HOL. / WDO NDF", "LEG. HOL. / WDO NDF OT",
    "U.LEG. HOL.", "UT",
]
DAY_COLUMNS = {"DAYS", "WDO", "SPL. HOL.", "SPL. HOL. / WDO", "LEG. HOL.",
               "LEG. HOL. / WDO", "U.LEG. HOL."}

# Remarks Tag Matrix: tag -> (day credit, OT, NDF, NDF OT) columns.
TAG_MATRIX: dict[str, tuple[str, str, str, str]] = {
    "": ("DAYS", "REG. OT", "NDF", "NDF OT"),
    "WDO": ("WDO", "WDO OT", "WDO NDF", "WDO NDF OT"),
    "SH": ("SPL. HOL.", "SPL. HOL. OT", "SPL. HOL. NDF", "SPL. HOL. NDF OT"),
    "SHWDO": ("SPL. HOL. / WDO", "SPL. HOL. / WDO OT", "SPL. HOL. / WDO NDF",
              "SPL. HOL. / WDO NDF OT"),
    "LH": ("LEG. HOL.", "LEG. HOL. OT", "LEG. HOL. NDF", "LEG. HOL. NDF OT"),
    "LHWDO": ("LEG. HOL. / WDO", "LEG. HOL. / WDO OT", "LEG. HOL. / WDO NDF",
              "LEG. HOL. / WDO NDF OT"),
}
# Tags the clients' templates still use for the same things.
TAG_ALIASES = {"SPL/WD": "SHWDO", "SPLWD": "SHWDO", "SPL": "SH", "LEGWD": "LHWDO",
               "LEG/WD": "LHWDO", "LEG": "LH", "UL": "ULH", "RD/WDO": "WDO"}

# Schedule-column values that are not a shift and earn no day credit.
NON_SHIFT = {"rd", "rta", "med.cert", "med cert", "medcert", "absent", "back out",
             "backout", "-", "--", "", "off", "rest day", "restday"}
REST_DAY = {"rd", "rest day", "restday", "off"}

NIGHT_START, NIGHT_END = 22 * 60, 6 * 60          # 10:00 PM - 6:00 AM


@dataclass(frozen=True)
class Profile:
    """One client's rules. 24 Chicken is the first; Peak 24 differs only in where
    the shift comes from (its DTRs have no schedule column)."""
    name: str
    shift_minutes: int = 600            # 10 hrs = 8 work + 2 break
    ot_minimum: int = 30
    ot_step: int = 30
    nd_minimum: int = 60                # rules sheet; open item with Dom
    nd_step: int = 30
    early_in_is_ot: bool = False
    off_schedule_flag: int = 30         # time in this far from the start -> review
    half_day_minutes: int = 300         # more than this short -> review
    # Peak 24: (latest time in, shift start) pairs from its template's ladder.
    time_in_ladder: tuple[tuple[int, int], ...] = ()
    ladder_default_start: int | None = None


CHICKEN_24 = Profile(name="24 Chicken")
PEAK_24 = Profile(
    name="24 Chicken - Peak 24",
    # The Peak 24 template's own SWITCH: in by 9:30 -> 9-7, by 10:30 -> 10-8, ...
    time_in_ladder=((9 * 60 + 30, 9 * 60), (10 * 60 + 30, 10 * 60), (11 * 60 + 30, 11 * 60),
                    (12 * 60 + 15, 12 * 60), (12 * 60 + 45, 12 * 60 + 30),
                    (13 * 60 + 30, 13 * 60)),
    ladder_default_start=9 * 60,
)
PROFILES = {"oppa_foods": CHICKEN_24, "peak24": PEAK_24}


# ── parsing ─────────────────────────────────────────────────────────────────

_CLOCK = re.compile(r"^\s*(\d{1,2})(?::|\.)(\d{2})(?::(\d{2}))?\s*([ap])?\.?\s*m?\.?\s*$", re.I)


def parse_time(value) -> int | None:
    """'13:05', '1:05 PM', '25:30', 0.5 (an Excel day fraction) -> minutes; else None."""
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return round(float(value) * 1440) if 0 <= value < 3 else None
    m = _CLOCK.match(str(value))
    if not m:
        return None
    h, mi = int(m.group(1)), int(m.group(2))
    ampm = (m.group(4) or "").lower()
    if ampm and 1 <= h <= 12:
        h = h % 12 + (12 if ampm == "p" else 0)
    if h > 47 or mi > 59:
        return None
    return h * 60 + mi


def _hour_to_minutes(h: int, mi: int = 0) -> int:
    """Short-form schedule hour: 7-12 are AM or noon, 1-6 are PM."""
    if 1 <= h <= 6:
        h += 12
    return h * 60 + mi


_SHORT_SCHED = re.compile(r"^\s*(\d{1,2})(?::(\d{2}))?\s*([ap]m?)?\s*(?:-|to|–)\s*(\d{1,2})(?::(\d{2}))?\s*([ap]m?)?\s*$", re.I)


@dataclass
class Schedule:
    start: int | None            # minutes from midnight; None = not a shift
    end: int | None
    kind: str                    # "shift" | "rest" | "off" | "unreadable"
    text: str
    written_end_mismatch: bool = False


def parse_schedule(text, profile: Profile = CHICKEN_24) -> Schedule:
    """'12-10' -> 12:00-22:00; '1-11' -> 13:00-23:00; '9:00-19:00' as written; RD etc. -> no shift."""
    raw = "" if text is None else str(text)
    dotted = re.fullmatch(r"\s*(\d{1,2})\s*\.\s*(\d{1,2})\s*", raw)
    if dotted:                                   # "11.9" is how a smudged "11-9" reads
        raw = f"{dotted.group(1)}-{dotted.group(2)}"
    low = raw.strip().lower()
    if low in REST_DAY:
        return Schedule(None, None, "rest", raw)
    if low in NON_SHIFT:
        return Schedule(None, None, "off", raw)
    m = _SHORT_SCHED.match(raw)
    if not m:
        return Schedule(None, None, "unreadable", raw)
    sh, smi, sap, eh, emi, eap = (m.group(1), m.group(2), m.group(3), m.group(4), m.group(5), m.group(6))
    sh, eh = int(sh), int(eh)
    smi, emi = int(smi or 0), int(emi or 0)
    if sap:
        start = (sh % 12 + (12 if sap.lower().startswith("p") else 0)) * 60 + smi
    elif sh > 12:
        start = sh * 60 + smi                       # already 24-hour, e.g. "13:00-23:00"
    else:
        start = _hour_to_minutes(sh, smi)
    end = start + profile.shift_minutes
    # The written end only cross-checks: "12-10" says 10, and 12:00 + 10 hrs = 22:00.
    if eap:
        written = (eh % 12 + (12 if eap.lower().startswith("p") else 0)) * 60 + emi
    elif eh > 12:
        written = eh * 60 + emi
    else:
        written = (eh % 12) * 60 + emi
    mismatch = (written % 720) != (end % 720) and (written % 1440) != (end % 1440)
    return Schedule(start, end, "shift", raw, written_end_mismatch=mismatch)


def schedule_from_time_in(time_in: int, profile: Profile) -> Schedule:
    """Peak 24: no schedule column, so the shift comes from the time in (template ladder)."""
    for latest, start in profile.time_in_ladder:
        if time_in <= latest:
            return Schedule(start, start + profile.shift_minutes, "shift", "from time in")
    start = profile.ladder_default_start
    return Schedule(start, start + profile.shift_minutes, "shift", "from time in (default)")


def normalize_tag(tag) -> str:
    t = re.sub(r"\s+", "", str(tag or "")).upper()
    t = TAG_ALIASES.get(t, t)
    return t if t in TAG_MATRIX or t == "ULH" else ("" if t in ("", "-") else t)


_MONTHS = ("jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec")


def parse_date(value, default_year: int | None = None) -> date | None:
    """'10/26/26', '11/1/2026', '2026-11-01', 'Nov 1, 2026' -> date (month first, as the forms print)."""
    if isinstance(value, date):
        return value
    text = str(value or "").strip()
    m = re.match(r"^(\d{4})-(\d{1,2})-(\d{1,2})", text)
    if m:
        y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
    else:
        m = re.match(r"^(\d{1,2})[/.-](\d{1,2})(?:[/.-](\d{2,4}))?\b", text)
        if m:
            mo, d = int(m.group(1)), int(m.group(2))
            y = int(m.group(3)) if m.group(3) else default_year
        else:
            m = re.search(r"([A-Za-z]{3,})\.?\s+(\d{1,2}),?\s*(\d{4})?", text)
            if not m or m.group(1)[:3].lower() not in _MONTHS:
                return None
            mo, d = _MONTHS.index(m.group(1)[:3].lower()) + 1, int(m.group(2))
            y = int(m.group(3)) if m.group(3) else default_year
    if y is None:
        return None
    if y < 100:
        y += 2000
    try:
        return date(y, mo, d)
    except ValueError:
        return None


# Where each layout keeps the day's first in and last out (canonical slot names).
_IN_SLOTS = ("time_in", "first_in", "m_in", "a_in", "e_in", "reg_in", "s1_in")
_OUT_SLOTS = ("time_out", "final_out", "e_out", "a_out", "m_out", "reg_out", "s1_out")


# Printed money/hours columns where a checker's handwriting also tends to land.
_MARK_SLOTS = ("ot_hrs", "holiday_pay", "holiday_ot", "deductions", "normal", "total")
_NUMBER = re.compile(r"-?\d+(?:[.,]\d+)?")


def day_from_row(values: dict, default_year: int | None = None) -> "DayInput | None":
    """One extracted row (slot -> text) -> a DayInput; None when it has no readable date.

    Handwritten marks are taken from REMARKS and from any words that landed in
    the printed number columns ("WDO | 180" in HOLIDAY PAY); a bare number in
    REMARKS (an allowance printed beside the grid) is not a mark.
    """
    d = parse_date(values.get("date"), default_year)
    if d is None:
        return None
    t_in = next((values[k] for k in _IN_SLOTS if parse_time(values.get(k)) is not None), None)
    t_out = next((values[k] for k in _OUT_SLOTS if parse_time(values.get(k)) is not None), None)
    marks = [str(values.get("remarks") or "").strip()]
    for k in _MARK_SLOTS:
        text = str(values.get(k) or "")
        if re.search(r"[A-Za-z]{2,}", text):
            marks.append(text.strip())
    marks = [m for m in marks if m and not re.fullmatch(r"[\d.,\s|-]+", m)]
    pay_text = str(values.get("holiday_pay") or "")
    pay = _NUMBER.search(pay_text.replace(",", ""))
    return DayInput(d, str(values.get("schedule") or ""), t_in, t_out,
                    remarks=" | ".join(dict.fromkeys(marks)),
                    holiday_pay=pay.group(0) if pay else None)


# ── per-day calculation ─────────────────────────────────────────────────────

def _overlap(a0: int, a1: int, b0: int, b1: int) -> int:
    return max(0, min(a1, b1) - max(a0, b0))


def night_minutes(start: int, end: int) -> int:
    """Minutes of [start, end) inside any 10:00 PM - 6:00 AM window (spans midnight)."""
    total = 0
    for day in range(-1, 3):
        base = day * 1440
        total += _overlap(start, end, base + NIGHT_START, base + 1440 + NIGHT_END)
    return total


def round_down(minutes: int, minimum: int, step: int) -> int:
    """0 below the minimum, else rounded down to whole steps (OT: 0.5 hr, 1.0 hr, ...)."""
    if minutes < minimum:
        return 0
    return (minutes // step) * step


@dataclass
class DayInput:
    date: date
    schedule: str = ""
    time_in: object = None
    time_out: object = None
    tag: str = ""                 # a Remarks tag already decided (template, review screen)
    remarks: str = ""             # handwritten marks read off the scan: WDO, SPL, OT ...
    holiday_pay: object = None    # what the sheet printed in HOLIDAY PAY


@dataclass
class DayResult:
    date: date
    tag: str                                  # the tag the day was routed by
    worked: bool = False
    late: int = 0
    undertime: int = 0
    ot: int = 0                               # minutes, already rounded
    nd: int = 0
    nd_ot: int = 0
    columns: dict[str, float] = field(default_factory=dict)
    flags: list[str] = field(default_factory=list)
    suggested_tag: str | None = None
    schedule: Schedule | None = None


def calculate_day(day: DayInput, profile: Profile = CHICKEN_24) -> DayResult:
    tag = normalize_tag(day.tag)
    res = DayResult(date=day.date, tag=tag)
    if tag not in TAG_MATRIX and tag != "ULH":
        res.flags.append(f"unknown Remarks tag {day.tag!r}")
        tag = res.tag = ""

    t_in, t_out = parse_time(day.time_in), parse_time(day.time_out)
    if profile.time_in_ladder and t_in is not None:
        sched = schedule_from_time_in(t_in, profile)
    else:
        sched = parse_schedule(day.schedule, profile)
    res.schedule = sched

    if t_in is None and t_out is None:
        if sched.kind == "shift" and tag not in ("ULH",):
            res.flags.append("scheduled shift but no time entries")
        if sched.kind == "unreadable" and str(day.schedule).strip():
            res.flags.append(f"unreadable schedule {day.schedule!r}")
        return res                              # nothing worked; ULH is decided per employee
    if t_in is None or t_out is None:
        res.flags.append("time in without time out" if t_out is None else "time out without time in")
        return res
    if t_out <= t_in:
        t_out += 1440                           # out after midnight

    if tag == "ULH":
        res.flags.append("tagged ULH but has time entries")
        tag = res.tag = ""

    if sched.kind == "rest":
        if not tag:
            res.suggested_tag = "WDO"           # Dom: RD with time entries -> WDO
            tag = res.tag = "WDO"
            res.flags.append("RD with time entries: suggested WDO")
        sched = Schedule(t_in, t_in + profile.shift_minutes, "shift", "assumed from time in")
        res.flags.append("no shift on the schedule; assumed 10 hrs from time in")
    elif sched.kind != "shift":
        res.flags.append(f"unreadable schedule {day.schedule!r}; assumed 10 hrs from time in")
        sched = Schedule(t_in, t_in + profile.shift_minutes, "shift", "assumed from time in")
    if sched.written_end_mismatch:
        res.flags.append(f"schedule {day.schedule!r}: written end doesn't match a 10-hr shift")

    start, end = sched.start, sched.end
    res.worked = True
    if abs(t_in - start) >= profile.off_schedule_flag:
        res.flags.append(f"time in {t_in // 60}:{t_in % 60:02d} is "
                         f"{abs(t_in - start)} mins off the {start // 60}:{start % 60:02d} start")

    # Late + undertime (Dom #2): short time counts from the later of in and start.
    res.late = max(0, t_in - start)
    worked_span = t_out - max(t_in, start)
    res.undertime = max(0, profile.shift_minutes - worked_span)
    if res.late + res.undertime > profile.half_day_minutes:
        res.flags.append(f"{res.late + res.undertime} mins late/short: more than half a day")

    # Night time inside the shift; early-in time is not counted (24 Chicken).
    inside0 = t_in if profile.early_in_is_ot else max(t_in, start)
    inside1 = min(t_out, end)
    nd_raw = night_minutes(inside0, inside1) if inside1 > inside0 else 0

    # OT after the scheduled end, split at the night window.
    ot_raw = nd_ot_raw = 0
    if t_out > end:
        nd_ot_raw = night_minutes(end, t_out)
        ot_raw = (t_out - end) - nd_ot_raw
    if res.late and (ot_raw or nd_ot_raw):
        res.flags.append("late and stayed past the shift end: OT or make-up? (open item)")

    res.ot = round_down(ot_raw, profile.ot_minimum, profile.ot_step)
    res.nd = round_down(nd_raw, profile.nd_minimum, profile.nd_step)
    res.nd_ot = round_down(nd_ot_raw, profile.nd_minimum, profile.nd_step)

    day_col, ot_col, nd_col, ndot_col = TAG_MATRIX[tag]
    cols = res.columns
    cols["DAYS"] = cols.get("DAYS", 0) + 1     # rule #1: every worked day counts in DAYS
    if day_col != "DAYS":
        cols[day_col] = cols.get(day_col, 0) + 1
    cols[ot_col] = res.ot / 60
    cols[nd_col] = res.nd / 60
    cols[ndot_col] = res.nd_ot / 60
    cols["UT"] = res.late + res.undertime
    return res


# ── per-employee summary (one DTR Summary row) ──────────────────────────────

@dataclass
class EmployeeResult:
    name: str
    days: list[DayResult]
    totals: dict[str, float]
    flags: list[str]


def calculate_employee(name: str, days: list[DayInput], profile: Profile = CHICKEN_24,
                       legal_holidays: set[date] | frozenset = frozenset()) -> EmployeeResult:
    """All of one employee's days in a cut-off -> the 26 summary columns.

    `legal_holidays`: dates that are regular holidays. A day tagged ULH, or a
    legal holiday with no time entries, earns 1 U.LEG. HOL. if the employee
    worked the last scheduled workday before it (rest days skipped).
    """
    ordered = sorted(days, key=lambda d: d.date)
    results = [calculate_day(d, profile) for d in ordered]
    flags: list[str] = []
    totals = {c: 0.0 for c in COLUMNS}

    by_date = {r.date: (d, r) for d, r in zip(ordered, results)}
    for d, r in zip(ordered, results):
        is_lh = d.date in legal_holidays or normalize_tag(d.tag) == "ULH"
        if not is_lh or r.worked:
            continue
        verdict = _ulh(d.date, by_date)
        if verdict is None:
            r.flags.append("ULH: the last scheduled workday is before this cut-off; check it")
        elif verdict:
            r.columns["U.LEG. HOL."] = 1
            r.tag = "ULH"

    for r in results:
        for col, v in r.columns.items():
            totals[col] += v
        flags.extend(f"{r.date.isoformat()}: {f}" for f in r.flags)
    return EmployeeResult(name=name, days=results, totals=totals, flags=flags)


# ── Remarks tag suggestions (phase 2) ───────────────────────────────────────

@dataclass
class Suggestion:
    tag: str             # "" = regular day
    reason: str
    sure: bool           # True: pre-filled on the review screen; False: shown as a hint only


def _marks(remarks) -> set[str]:
    words = set(re.findall(r"[A-Z]+", str(remarks or "").upper()))
    out = set()
    if words & {"WDO", "SHWDO", "LHWDO"} or "DAYOFF" in str(remarks or "").upper().replace(" ", ""):
        out.add("WDO")
    if words & {"SPL", "SH", "SPECIAL", "SHWDO"}:
        out.add("SH")
    if words & {"LH", "LEGAL", "LEG", "LHWDO", "REGULAR"}:
        out.add("LH")
    return out


def _amount(value) -> float:
    try:
        return float(str(value).replace(",", "").strip() or 0)
    except ValueError:
        return 0.0


def suggest_tags(days: list[DayInput], profile: Profile = CHICKEN_24,
                 extra_holidays: dict | None = None) -> dict[date, Suggestion]:
    """A suggested Remarks tag for every worked day that needs one, with its reason.

    Sure suggestions come from the holiday calendar and from the schedule saying
    RD; a handwritten WDO is sure too (the checker wrote it). Weaker evidence
    (holiday pay or a holiday mark on a date that is not a national holiday, or
    working a weekday that is a rest day in the other week) is only a hint.
    """
    import holidays as hol

    rest_weekdays: dict[int, list[date]] = {}
    for d in days:
        if parse_schedule(d.schedule, profile).kind == "rest":
            rest_weekdays.setdefault(d.date.weekday(), []).append(d.date)

    out: dict[date, Suggestion] = {}
    for d in days:
        worked = parse_time(d.time_in) is not None and parse_time(d.time_out) is not None
        if not worked:
            continue
        sched = parse_schedule(d.schedule, profile)
        marks = _marks(d.remarks)
        rest_day = sched.kind == "rest" or "WDO" in marks
        rest_why = ("schedule says RD" if sched.kind == "rest"
                    else "WDO written on the scan") if rest_day else ""
        cal = hol.holiday(d.date, extra_holidays)
        if cal:
            kind, name = cal
            base = "LH" if kind == hol.REGULAR else "SH"
            tag = base + ("WDO" if rest_day else "")
            why = f"{name} ({'regular holiday' if base == 'LH' else 'special non-working day'})"
            out[d.date] = Suggestion(tag, why + (f", and {rest_why}" if rest_day else ""), True)
        elif marks & {"SH", "LH"} or _amount(d.holiday_pay) > 0:
            base = "LH" if "LH" in marks else "SH"
            evidence = "holiday mark on the scan" if marks & {"SH", "LH"} else "holiday pay printed"
            out[d.date] = Suggestion(base + ("WDO" if rest_day else ""),
                                     f"{evidence}, but {d.date:%d %b} is not a national holiday:"
                                     " a local holiday?", False)
        elif rest_day:
            out[d.date] = Suggestion("WDO", rest_why[:1].upper() + rest_why[1:], True)
        else:
            others = [x for x in rest_weekdays.get(d.date.weekday(), []) if x != d.date]
            if others and sched.kind == "shift":
                out[d.date] = Suggestion("WDO", f"{d.date:%A} was a rest day on "
                                         f"{others[0]:%d %b}: a rest day worked?", False)
    return out


@dataclass
class ReviewResult:
    employee: EmployeeResult
    suggestions: dict[date, Suggestion]
    tags: dict[date, str]          # the tag each day was calculated with


def review_employee(name: str, days: list[DayInput], profile: Profile = CHICKEN_24,
                    confirmed: dict[date, str] | None = None,
                    extra_holidays: dict | None = None) -> ReviewResult:
    """Suggest tags, apply them under the team's confirmed ones, and calculate.

    Precedence per day: a tag the team confirmed on the review screen, then a tag
    the input already carries, then a sure suggestion. Hints never apply on
    their own.
    """
    import holidays as hol
    confirmed = confirmed or {}
    suggestions = suggest_tags(days, profile, extra_holidays)
    tagged = []
    for d in days:
        if d.date in confirmed:
            tag = confirmed[d.date]
        elif normalize_tag(d.tag):
            tag = d.tag
        elif d.date in suggestions and suggestions[d.date].sure:
            tag = suggestions[d.date].tag
        else:
            tag = ""
        tagged.append(DayInput(d.date, d.schedule, d.time_in, d.time_out, tag,
                               d.remarks, d.holiday_pay))
    emp = calculate_employee(name, tagged, profile, hol.legal_holidays(extra_holidays))
    return ReviewResult(emp, suggestions, {d.date: normalize_tag(d.tag) for d in tagged})


def _ulh(holiday: date, by_date: dict) -> bool | None:
    """Did the employee work the last scheduled workday before `holiday`? None = not in range."""
    day = holiday - timedelta(days=1)
    while day in by_date:
        d, r = by_date[day]
        if r.schedule is not None and r.schedule.kind == "rest" and not r.worked:
            day -= timedelta(days=1)            # Dom #8: skip the rest day, look further back
            continue
        return r.worked
    return None
