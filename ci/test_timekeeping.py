"""Unit tests for app/timekeeping.py, run by Cloud Build (preflight step).

Each test is a rule DPAM confirmed (Dom, 23 Nov 2026) or a day from their answer
key, written with made-up names. Plain asserts:
    python3 ci/test_timekeeping.py
"""
from __future__ import annotations

import pathlib
import sys
from datetime import date

ROOT = pathlib.Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "app"))
import timekeeping as tk  # noqa: E402
from timekeeping import DayInput  # noqa: E402

D = date(2026, 11, 10)


def day(schedule, t_in, t_out, tag="", when=D):
    return tk.calculate_day(DayInput(when, schedule, t_in, t_out, tag))


def test_schedule_short_forms():
    """'12-10' = 12:00-22:00; 1-6 are PM; long forms are read as written."""
    cases = {"12-10": (720, 1320), "9-7": (540, 1140), "10-8": (600, 1200),
             "11-9": (660, 1260), "1-11": (780, 1380), "12:00-22:00": (720, 1320),
             "13:00-23:00": (780, 1380), "12:30-22:30": (750, 1350), "9AM-7PM": (540, 1140)}
    for text, (start, end) in cases.items():
        s = tk.parse_schedule(text)
        assert (s.kind, s.start, s.end) == ("shift", start, end), (text, s)
        assert not s.written_end_mismatch, text
    assert tk.parse_schedule("10-6").written_end_mismatch      # 10:00 + 10 hrs is 20:00, not 6
    for text in ("RD", "rta", "med.cert", "absent", "back out", "-", ""):
        assert tk.parse_schedule(text).start is None, text


def test_regular_day():
    r = day("10-8", "10:00", "20:00")
    assert r.columns["DAYS"] == 1 and r.late == 0 and r.undertime == 0 and r.ot == 0
    assert not r.flags, r.flags


def test_dom_late_and_short_is_120():
    """Dom #2: 10-8, in 11:00, out 20:00 = 1 day, 60 late + 60 short = 120 mins."""
    r = day("10-8", "11:00", "20:00")
    assert r.columns["DAYS"] == 1
    assert (r.late, r.undertime, r.columns["UT"]) == (60, 60, 120)


def test_leaving_early_is_undertime():
    r = day("10-8", "10:00", "19:15")
    assert (r.late, r.undertime) == (0, 45)


def test_early_in_is_not_ot_or_credit():
    """24 Chicken: time before the scheduled start is not OT (Dom #5)."""
    r = day("10-8", "9:00", "20:00")
    assert r.ot == 0 and r.undertime == 0 and r.late == 0


def test_ot_half_hour_minimum_and_steps():
    """Dom #3: OT from 0.5 hr, then 0.5-hr steps, rounded down."""
    assert day("10-8", "10:00", "20:20").ot == 0
    assert day("10-8", "10:00", "20:30").ot == 30
    assert day("10-8", "10:00", "20:40").ot == 30
    assert day("10-8", "10:00", "21:00").columns["REG. OT"] == 1.0


def test_ot_before_10pm_is_regular():
    """Answer key, 28 Oct: 11-9, out 22:00 = 1 hr REG. OT and no NDF OT."""
    r = day("11-9", "11:00", "22:00")
    assert r.columns["REG. OT"] == 1.0 and r.columns["NDF OT"] == 0


def test_ot_after_10pm_is_ndf_ot():
    """Answer key, 29 Oct: 12-10, out 23:00 = 1 hr NDF OT, no REG. OT."""
    r = day("12-10", "12:00", "23:00")
    assert r.columns["NDF OT"] == 1.0 and r.columns["REG. OT"] == 0


def test_out_after_midnight():
    """Rules sheet: shift ending 12:00 AM, out 2:00 AM = 2 hrs NDF OT."""
    r = day("2-12", "14:00", "2:00")
    assert r.schedule.end == 24 * 60
    assert r.columns["NDF OT"] == 2.0 and r.columns["NDF"] == 2.0   # 22:00-24:00 inside the shift


def test_night_inside_shift_is_ndf():
    """Answer key: 1-11 (13:00-23:00) = 1 hr NDF."""
    r = day("1-11", "13:00", "23:00")
    assert r.columns["NDF"] == 1.0 and r.columns["NDF OT"] == 0


def test_ndf_keeps_one_hour_minimum():
    """Rules sheet (open item with Dom): NDF from 1 hr."""
    r = day("12:30-22:30", "12:30", "22:30")
    assert r.nd == 0                      # 30 mins of night time: under the minimum


def test_tag_routes_hours_only_to_its_columns():
    """Matrix + Dom #4: SH hours go to the SPL. HOL. columns only; DAYS still counts the day."""
    r = day("12-10", "12:00", "23:00", tag="SH")
    assert r.columns["DAYS"] == 1 and r.columns["SPL. HOL."] == 1
    assert r.columns["SPL. HOL. NDF OT"] == 1.0
    assert "NDF OT" not in r.columns and "REG. OT" not in r.columns


def test_sh_ot_half_hour():
    """Answer key, 2 Nov: SH day on 10-8, out 20:30 = 0.5 hr SPL. HOL. OT."""
    assert day("10-8", "10:00", "20:30", tag="SH").columns["SPL. HOL. OT"] == 0.5


def test_shwdo_and_legacy_tag():
    """SHWDO routes to SPL. HOL. / WDO; the template's old SPL/WD means the same."""
    for tag in ("SHWDO", "SPL/WD"):
        r = day("10-8", "10:00", "20:00", tag=tag)
        assert r.columns["SPL. HOL. / WDO"] == 1 and r.columns["DAYS"] == 1, tag


def test_non_shift_values_earn_nothing():
    for text in ("RD", "rta", "med.cert", "absent", "back out", "-", ""):
        r = day(text, None, None)
        assert not r.worked and not r.columns, text


def test_rd_with_times_is_suggested_wdo():
    """Dom #7: time entries on an RD day -> suggest WDO."""
    r = day("RD", "10:00", "20:00")
    assert r.suggested_tag == "WDO" and r.columns["WDO"] == 1 and r.columns["DAYS"] == 1
    assert any("suggested WDO" in f for f in r.flags)


def test_flags():
    assert any("mins off" in f for f in day("10-8", "11:00", "20:00").flags)   # catches a misread 10:00
    assert any("time out" in f for f in day("10-8", "10:00", None).flags)
    assert any("half a day" in f for f in day("10-8", "10:00", "14:00").flags)
    assert any("OT or make-up" in f for f in day("10-8", "11:00", "21:00").flags)


def test_ulh_rest_day_before_holiday():
    """Dom #8: LH Monday, RD Sunday, worked Saturday -> ULH 1; absent Saturday -> 0."""
    sat, sun, mon = date(2026, 8, 29), date(2026, 8, 30), date(2026, 8, 31)
    worked = [DayInput(sat, "10-8", "10:00", "20:00"), DayInput(sun, "RD"), DayInput(mon, "10-8")]
    r = tk.calculate_employee("A", worked, legal_holidays={mon})
    assert r.totals["U.LEG. HOL."] == 1 and r.totals["DAYS"] == 1
    absent = [DayInput(sat, "absent"), DayInput(sun, "RD"), DayInput(mon, "10-8")]
    assert tk.calculate_employee("B", absent, legal_holidays={mon}).totals["U.LEG. HOL."] == 0


def test_peak24_shift_from_time_in():
    """Peak 24 has no schedule column: in by 9:30 -> 9-7, by 10:30 -> 10-8, ..."""
    r = tk.calculate_day(DayInput(D, "", "10:05", "20:00"), tk.PEAK_24)
    assert (r.schedule.start, r.schedule.end) == (600, 1200)
    assert r.late == 5 and r.undertime == 5


def test_employee_totals_add_up():
    days = [DayInput(date(2026, 11, d), "10-8", "10:00", "21:00") for d in (3, 4)]
    days.append(DayInput(date(2026, 11, 5), "RD"))
    t = tk.calculate_employee("C", days).totals
    assert t["DAYS"] == 2 and t["REG. OT"] == 2.0 and t["UT"] == 0
    assert list(t) == tk.COLUMNS


def test_holiday_calendar():
    """Proclamation 1006 (2026) and 1427 (2027)."""
    import holidays as hol
    assert hol.holiday(date(2026, 11, 1))[0] == hol.SPECIAL
    assert hol.holiday(date(2026, 11, 2))[0] == hol.SPECIAL          # additional special day
    assert hol.holiday(date(2026, 11, 30))[0] == hol.REGULAR
    assert hol.holiday(date(2026, 8, 31))[0] == hol.REGULAR          # last Monday of August
    assert hol.holiday(date(2027, 8, 30))[0] == hol.REGULAR
    assert hol.holiday(date(2026, 2, 25)) is None                    # special WORKING day
    assert date(2026, 12, 25) in hol.legal_holidays()


def suggest(days):
    return tk.suggest_tags(days)


def test_suggest_calendar_holidays():
    nov1, nov30 = date(2026, 11, 1), date(2026, 11, 30)
    s = suggest([DayInput(nov1, "12-10", "12:00", "22:00"), DayInput(nov30, "10-8", "10:00", "20:00")])
    assert (s[nov1].tag, s[nov1].sure) == ("SH", True)
    assert (s[nov30].tag, s[nov30].sure) == ("LH", True)


def test_suggest_holiday_on_rest_day():
    nov1 = date(2026, 11, 1)
    assert suggest([DayInput(nov1, "RD", "10:00", "20:00")])[nov1].tag == "SHWDO"
    assert suggest([DayInput(nov1, "10-8", "10:00", "20:00", remarks="WDO")])[nov1].tag == "SHWDO"


def test_suggest_wdo():
    d = date(2026, 11, 5)
    assert (suggest([DayInput(d, "RD", "9:00", "19:00")])[d].tag, True) == ("WDO", True)
    s = suggest([DayInput(d, "9-7", "9:00", "19:00", remarks="WDO")])[d]   # handwritten on the scan
    assert (s.tag, s.sure) == ("WDO", True)


def test_suggest_hints_are_not_applied():
    d = date(2026, 11, 5)
    s = suggest([DayInput(d, "9-7", "9:00", "19:00", remarks="SPL")])[d]
    assert (s.tag, s.sure) == ("SH", False) and "local holiday" in s.reason
    assert suggest([DayInput(d, "9-7", "9:00", "19:00", holiday_pay="180")])[d].sure is False
    r = tk.review_employee("D", [DayInput(d, "9-7", "9:00", "19:00", remarks="SPL")])
    assert r.tags[d] == "" and r.employee.totals["SPL. HOL."] == 0


def test_suggest_rest_day_pattern_hint():
    """Worked on a weekday that was RD the week before, with an unreadable mark: a hint."""
    thu1, thu2 = date(2026, 10, 29), date(2026, 11, 5)
    s = suggest([DayInput(thu1, "RD"), DayInput(thu2, "9-7", "9:00", "19:00", remarks="NDP")])
    assert (s[thu2].tag, s[thu2].sure) == ("WDO", False)
    # Without anything written on the row the pattern alone is noise (rosters move).
    assert thu2 not in suggest([DayInput(thu1, "RD"), DayInput(thu2, "9-7", "9:00", "19:00")])
    # Readable marks (OT, NDF) are not a reason either.
    assert thu2 not in suggest([DayInput(thu1, "RD"), DayInput(thu2, "9-7", "9:00", "20:00", remarks="1 OT")])


def test_review_confirmed_tag_wins():
    nov1 = date(2026, 11, 1)
    days = [DayInput(nov1, "10-8", "10:00", "20:00")]
    assert tk.review_employee("E", days).employee.totals["SPL. HOL."] == 1
    r = tk.review_employee("E", days, confirmed={nov1: "SHWDO"})
    assert r.employee.totals["SPL. HOL. / WDO"] == 1 and r.employee.totals["SPL. HOL."] == 0


def test_review_ulh_from_calendar():
    """31 Aug 2026 (regular holiday, Monday): rest day Sunday, worked Saturday -> ULH."""
    days = [DayInput(date(2026, 8, 29), "10-8", "10:00", "20:00"),
            DayInput(date(2026, 8, 30), "RD"), DayInput(date(2026, 8, 31), "10-8")]
    assert tk.review_employee("F", days).employee.totals["U.LEG. HOL."] == 1


def test_parse_date_forms():
    for text in ("10/26/26", "10/26/2026", "2026-10-26", "Oct 26, 2026", "10-26-26"):
        assert tk.parse_date(text) == date(2026, 10, 26), text
    assert tk.parse_date("10/26", 2026) == date(2026, 10, 26)
    assert tk.parse_date("TOTAL") is None


def test_day_from_row_layouts():
    """In/out sheet slots and DTR (morning/afternoon/evening) slots both give in + last out."""
    a = tk.day_from_row({"date": "11/5/26", "schedule": "9-7", "time_in": "9:00",
                         "time_out": "19:00", "remarks": "WDO", "holiday_pay": "180"})
    assert (a.time_in, a.time_out, a.remarks, a.schedule) == ("9:00", "19:00", "WDO", "9-7")
    b = tk.day_from_row({"date": "11/5/26", "m_in": "9:00", "m_out": "13:00", "a_in": "14:00",
                         "a_out": "18:00", "e_in": "", "e_out": "19:00"})
    assert (b.time_in, b.time_out) == ("9:00", "19:00")
    assert tk.day_from_row({"date": "", "time_in": "9:00"}) is None


def test_marks_found_in_number_columns():
    """Real scans: handwriting lands in HOLIDAY PAY / OT HRS; an allowance '60' lands in REMARKS."""
    d = tk.day_from_row({"date": "11/1/26", "schedule": "10-8", "time_in": "10:00", "time_out": "20:00",
                         "holiday_pay": "WDO | 180", "remarks": "60"})
    assert "WDO" in d.remarks and "60" not in d.remarks and d.holiday_pay == "180"
    assert tk.suggest_tags([d])[d.date].tag == "SHWDO"
    plain = tk.day_from_row({"date": "11/5/26", "time_in": "9:00", "time_out": "19:00", "remarks": "50"})
    assert plain.remarks == ""


def test_dotted_schedule():
    """'11.9' is how a smudged '11-9' comes back."""
    for text in ("11.9", "11:9", "11 . 9"):
        s = tk.parse_schedule(text)
        assert (s.kind, s.start, s.end) == ("shift", 660, 1260), text
    assert tk.parse_schedule("12:30").kind == "unreadable"      # a clock time, not a schedule


def test_misread_schedule_start_taken_from_end():
    """Real scan: a tick through '12-10' read as '17-10'. The end (22:00) fixes the start."""
    s = tk.parse_schedule("17-10")
    assert (s.start, s.end, s.repaired_from_end) == (720, 1320, True)
    r = day("17-10", "12:00", "22:00")
    assert r.columns["UT"] == 0 and any("from its end" in f for f in r.flags)


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
