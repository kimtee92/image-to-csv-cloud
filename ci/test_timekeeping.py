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
