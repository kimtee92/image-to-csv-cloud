"""Philippine national holidays, as declared by proclamation.

Used to suggest Remarks tags (SH, LH, SHWDO, LHWDO) and to find unworked legal
holidays (ULH). Local holidays are added per branch on the review screen.

Sources (proclamation texts, read 23 Nov 2026):
  2026: Proclamation No. 1006, s. 2025
        https://lawphil.net/executive/proc/proc2025/proc_1006_2025.html
  2027: Proclamation No. 1427, s. 2026
        https://lawphil.net/executive/proc/proc2026/proc_1427_2026.html
Eid'l Fitr and Eid'l Adha are declared by separate proclamations each year and
are not in either list: add them here when they are proclaimed.
"""
from __future__ import annotations

from datetime import date

REGULAR = "regular"     # legal holiday: LH / LHWDO / ULH
SPECIAL = "special"     # special non-working day: SH / SHWDO

_RAW = {
    2026: [
        ("01-01", REGULAR, "New Year's Day"),
        ("02-17", SPECIAL, "Chinese New Year"),
        ("04-02", REGULAR, "Maundy Thursday"),
        ("04-03", REGULAR, "Good Friday"),
        ("04-04", SPECIAL, "Black Saturday"),
        ("04-09", REGULAR, "Araw ng Kagitingan"),
        ("05-01", REGULAR, "Labor Day"),
        ("06-12", REGULAR, "Independence Day"),
        ("08-21", SPECIAL, "Ninoy Aquino Day"),
        ("08-31", REGULAR, "National Heroes Day"),
        ("11-01", SPECIAL, "All Saints' Day"),
        ("11-02", SPECIAL, "All Souls' Day"),
        ("11-30", REGULAR, "Bonifacio Day"),
        ("12-08", SPECIAL, "Feast of the Immaculate Conception of Mary"),
        ("12-24", SPECIAL, "Christmas Eve"),
        ("12-25", REGULAR, "Christmas Day"),
        ("12-30", REGULAR, "Rizal Day"),
        ("12-31", SPECIAL, "Last Day of the Year"),
    ],
    2027: [
        ("01-01", REGULAR, "New Year's Day"),
        ("02-06", SPECIAL, "Chinese New Year"),
        ("03-25", REGULAR, "Maundy Thursday"),
        ("03-26", REGULAR, "Good Friday"),
        ("03-27", SPECIAL, "Black Saturday"),
        ("04-09", REGULAR, "Araw ng Kagitingan"),
        ("05-01", REGULAR, "Labor Day"),
        ("06-12", REGULAR, "Independence Day"),
        ("08-21", SPECIAL, "Ninoy Aquino Day"),
        ("08-30", REGULAR, "National Heroes Day"),
        ("11-01", SPECIAL, "All Saints' Day"),
        ("11-02", SPECIAL, "All Souls' Day"),
        ("11-30", REGULAR, "Bonifacio Day"),
        ("12-08", SPECIAL, "Feast of the Immaculate Conception of Mary"),
        ("12-24", SPECIAL, "Christmas Eve"),
        ("12-25", REGULAR, "Christmas Day"),
        ("12-30", REGULAR, "Rizal Day"),
        ("12-31", SPECIAL, "Last Day of the Year"),
    ],
}
# EDSA anniversary (25 Feb) is a special WORKING day both years: an ordinary day.

NATIONAL: dict[date, tuple[str, str]] = {
    date(year, int(md[:2]), int(md[3:])): (kind, name)
    for year, rows in _RAW.items() for md, kind, name in rows
}
COVERED_YEARS = frozenset(_RAW)


def holiday(day: date, extra: dict[date, tuple[str, str]] | None = None) -> tuple[str, str] | None:
    """(kind, name) when `day` is a holiday, nationally or in `extra` (local ones)."""
    if extra and day in extra:
        return extra[day]
    return NATIONAL.get(day)


def legal_holidays(extra: dict[date, tuple[str, str]] | None = None) -> frozenset[date]:
    days = {d for d, (kind, _) in NATIONAL.items() if kind == REGULAR}
    days |= {d for d, (kind, _) in (extra or {}).items() if kind == REGULAR}
    return frozenset(days)
