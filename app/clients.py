"""Per-client configuration: which DTR layout comes in, which workbook goes out.

The portal is client-first — you pick the client, and that determines both how the
scans are read and which timekeeping workbook they are compiled into.

Two things vary per client and nothing else does:

  * INPUT LAYOUT — the physical shape of their DTR. Despite 100+ header spellings
    across the historical CSVs, there are only four real forms (see LAYOUTS).
  * OUTPUT TEMPLATE — their own .xlsx, shipped in app/templates/ and filled in
    place. Their formulas, number formats, Shift Schedule and Rules Per Column
    sheets all survive untouched, because we never rebuild the workbook — we only
    write values into its data columns.

That second point is the important one. An earlier version ported one client's
formulas into Python; doing that for every client would mean re-deriving a dozen
different payroll models by hand, and each would be a fresh chance to get someone's
pay wrong. Filling their real workbook sidesteps the whole class of problem, and
onboarding a client becomes: drop their .xlsx in templates/, add an entry here.
"""
from __future__ import annotations

# ── canonical slots ───────────────────────────────────────────────────────
# The neutral vocabulary every layout is mapped onto. Sessions are numbered
# rather than named because "session 2" is the afternoon on a DTR, the lunch
# break on Alberto's form, and "IN 2/OUT 2" on Tokyo's — same position, three
# different words for it.
SLOTS = [
    "schedule", "date", "day",
    "s1_in", "s1_out", "s1_total",
    "s2_in", "s2_out", "s2_total",
    "s3_in", "s3_out", "s3_total",
    "ot_in", "ot_out", "ot_total",
    "final_out", "days_hrs", "ot_hrs", "reason", "remarks",
]

# ── input layouts ─────────────────────────────────────────────────────────
# `headers` is the locked column list the extractor emits; `slots` is the
# parallel canonical name for each, used to drive the workbook fill.
LAYOUTS: dict[str, dict] = {
    "dtr": {
        "label": "DTR — Morning / Afternoon / Evening",
        "headers": [
            "DAILY TIME SCHEDULE", "DATE",
            "MORNING IN", "MORNING OUT",
            "AFTERNOON IN", "AFTERNOON OUT",
            "EVENING IN", "EVENING OUT",
            "NO. OF DAYS/HRS", "O.T. HRS", "REASON FOR O.T.",
        ],
        "slots": [
            "schedule", "date",
            "s1_in", "s1_out", "s2_in", "s2_out", "s3_in", "s3_out",
            "days_hrs", "ot_hrs", "reason",
        ],
        "sections": ["MORNING", "AFTERNOON", "EVENING"],
    },
    "dtr_ot": {
        "label": "DTR — Morning / Afternoon / Overtime",
        "headers": [
            "DAILY TIME SCHEDULE", "DATE",
            "MORNING IN", "MORNING OUT",
            "AFTERNOON IN", "AFTERNOON OUT",
            "OVERTIME IN", "OVERTIME OUT",
            "DAILY TOTAL", "O.T. HRS", "REASON FOR O.T.",
        ],
        "slots": [
            "schedule", "date",
            "s1_in", "s1_out", "s2_in", "s2_out", "ot_in", "ot_out",
            "days_hrs", "ot_hrs", "reason",
        ],
        "sections": ["MORNING", "AFTERNOON", "OVERTIME"],
    },
    "timesheet": {
        "label": "Timesheet — Regular / Lunch / Coffee / Overtime",
        "headers": [
            "DATE", "PLOTTED SCHEDULE", "DUTY LOCATION",
            "REGULAR HOURS - Time in", "REGULAR HOURS - Time out", "REGULAR HOURS - TOTAL",
            "LUNCH BREAK - Break out", "LUNCH BREAK - Break in", "LUNCH BREAK - TOTAL",
            "COFFEE BREAK - Break out", "COFFEE BREAK - Break in", "COFFEE BREAK - TOTAL",
            "OVERTIME - Time in", "OVERTIME - Time out", "OVERTIME - TOTAL",
            "TARDINESS", "REASON/REMARKS",
        ],
        "slots": [
            "date", "schedule", "duty_location",
            "s1_in", "s1_out", "s1_total",
            "s2_in", "s2_out", "s2_total",
            "s3_in", "s3_out", "s3_total",
            "ot_in", "ot_out", "ot_total",
            "tardiness", "reason",
        ],
        "sections": ["REGULAR HOURS", "LUNCH BREAK", "COFFEE BREAK", "OVERTIME"],
    },
    "inout": {
        "label": "Simple in / out payroll sheet",
        "headers": [
            "DATE", "DAY", "SCHEDULE", "TIME IN", "TIME OUT", "BREAK (HR)", "NORMAL",
            "OT HRS", "HOLIDAY PAY", "HOLIDAY+OT", "TOTAL", "DEDUCTIONS", "NET TOTAL",
            "REMARKS",
        ],
        "slots": [
            "date", "day", "schedule", "s1_in", "final_out", "break_hrs", "normal",
            "ot_hrs", "holiday_pay", "holiday_ot", "total", "deductions", "net_total",
            "remarks",
        ],
        "sections": [],
    },
}

# ── output templates ──────────────────────────────────────────────────────
# `blocks` is how many employee blocks the shipped workbook physically contains.
# It is recorded rather than sniffed because the marker is a shared string, and it
# matters: Miniso's template has exactly ONE block against 284 employees, so the
# filler clones blocks up to the batch size (see xlsx_fill.clone_block).
# `columns` maps a spreadsheet column letter to the canonical slot that fills it,
# read straight off each workbook's own header row. `first_block` / `stride` /
# `rows_per_block` describe the employee-block geometry, which is a 20-row stride
# on every one of these workbooks; only the starting row differs.
TEMPLATES: dict[str, dict] = {
    "standard_dtr": {
        "file": "standard_dtr.xlsx", "sheet": "Timekeeping",
        "blocks": 15,
        "first_block": 3, "stride": 20, "rows_per_block": 16, "name_col": "C",
        "columns": {"B": "date", "C": "day", "D": "s1_in", "E": "s1_out",
                    "F": "s2_in", "G": "s2_out", "H": "s3_in", "I": "s3_out"},
    },
    "tokyo": {
        # The revised workbook the client keeps on Drive: one "Sheet1" plus the
        # DTR Summary, with the ND column dropped. The version shipped before was
        # their older four-sheet file, which is what "does not align with the
        # required format" referred to.
        "file": "tokyo.xlsx", "sheet": "Sheet1",
        "blocks": 15,
        "first_block": 3, "stride": 20, "rows_per_block": 16, "name_col": "C",
        # The TOTAL HRS formula is IF(AND(IN<>"", OUT<>""), OUT - IN, 0) and reads only
        # these two columns, so they take the day's first in and last out — not the
        # literal MORNING IN / EVENING OUT boxes, which left every day without an
        # evening punch at 0 hours (or negative, where OUT is a formula).
        "columns": {"B": "date", "C": "day", "D": "first_in", "E": "s1_out",
                    "F": "s2_in", "G": "s2_out", "H": "s3_in", "I": "final_out"},
    },
    "baobab": {
        "file": "baobab.xlsx", "sheet": "Timekeeping",
        "blocks": 3,
        "first_block": 4, "stride": 20, "rows_per_block": 16, "name_col": "C",
        # The TOTAL HRS formula is IF(AND(IN<>"", OUT<>""), OUT - IN, 0) and reads only
        # these two columns, so they take the day's first in and last out — not the
        # literal MORNING IN / EVENING OUT boxes, which left every day without an
        # evening punch at 0 hours (or negative, where OUT is a formula).
        "columns": {"A": "date", "B": "day", "C": "first_in", "D": "s1_out",
                    "E": "s2_in", "F": "s2_out", "G": "s3_in", "H": "final_out"},
    },
    # 24 Chicken (Oppa Foods and Peak 24): DPAM's 26-column template from the
    # Oct 26 - Nov 10, 2026 cut-off, blanked (no employee data) and its DTR Summary
    # extended to 40 rows. The app calculates every computed column itself
    # (app/timekeeping.py) and writes VALUES over the template's formulas, so the
    # file needs no LET/XLOOKUP/SWITCH and opens the same in any Excel. Units are
    # the template's own: a credited day is 10 hrs (its totals divide by 10), hours
    # are fractions of a day, Late/UT is minutes / 1440.
    "chicken24": {
        "file": "chicken24.xlsx", "sheet": "Timekeeping",
        "blocks": 10,
        "first_block": 4, "stride": 20, "rows_per_block": 16, "name_col": "C",
        # TOTAL HRS reads D (first in) and I (last out) only.
        "columns": {"A": "schedule", "B": "date", "C": "day", "D": "first_in", "I": "final_out"},
        "computed": {
            "N": "_hours", "O": "DAYS", "P": "REG. OT", "Q": "NDF", "R": "NDF OT",
            "S": "WDO", "T": "WDO OT", "U": "WDO NDF", "V": "WDO NDF OT",
            "W": "SPL. HOL.", "X": "SPL. HOL. OT", "Y": "SPL. HOL. NDF", "Z": "SPL. HOL. NDF OT",
            "AA": "SPL. HOL. / WDO", "AB": "SPL. HOL. / WDO OT",
            "AC": "SPL. HOL. / WDO NDF", "AD": "SPL. HOL. / WDO NDF OT",
            "AE": "LEG. HOL.", "AF": "LEG. HOL. OT", "AG": "LEG. HOL. NDF", "AH": "LEG. HOL. NDF OT",
            "AI": "LEG. HOL. / WDO", "AJ": "LEG. HOL. / WDO OT",
            "AK": "LEG. HOL. / WDO NDF", "AL": "LEG. HOL. / WDO NDF OT",
            "AM": "U.LEG. HOL.", "AN": "UT", "AP": "_schedule", "AQ": "_tag",
        },
        "summary": {"sheet": "DTR Summary", "first_row": 4, "last_row": 43, "last_col": "AB"},
    },
    "miniso": {
        "file": "miniso.xlsx", "sheet": "Timekeeping",
        # Eight block skeletons ship in the file, numbered 1-8. Declaring 1 made
        # every batch clone over the other seven and left them out of the
        # unused-block clear.
        "blocks": 8,
        "first_block": 3, "stride": 20, "rows_per_block": 16, "name_col": "C",
        # The TOTAL HRS formula is IF(AND(IN<>"", OUT<>""), OUT - IN, 0) and reads only
        # these two columns, so they take the day's first in and last out — not the
        # literal MORNING IN / EVENING OUT boxes, which left every day without an
        # evening punch at 0 hours (or negative, where OUT is a formula).
        "columns": {"A": "schedule", "B": "date", "C": "day",
                    "D": "first_in", "E": "s1_out", "F": "s2_in", "G": "s2_out",
                    "H": "s3_in", "I": "final_out", "J": "ot_hrs", "K": "reason"},
    },
    "bikezilla": {
        "file": "bikezilla.xlsx", "sheet": "Timekeeping",
        "blocks": 7,
        "first_block": 4, "stride": 20, "rows_per_block": 16, "name_col": "C",
        # The TOTAL HRS formula is IF(AND(IN<>"", OUT<>""), OUT - IN, 0) and reads only
        # these two columns, so they take the day's first in and last out — not the
        # literal MORNING IN / EVENING OUT boxes, which left every day without an
        # evening punch at 0 hours (or negative, where OUT is a formula).
        "columns": {"A": "schedule", "B": "date",
                    "C": "first_in", "D": "s1_out", "E": "s2_in", "F": "s2_out",
                    "G": "s3_in", "H": "final_out"},   # I is the client's FINAL OUT formula
    },
    "trinity": {
        "file": "trinity.xlsx", "sheet": "Timekeeping",
        "blocks": 15,
        "first_block": 3, "stride": 20, "rows_per_block": 16, "name_col": "C",
        # D and E are the day's only in/out boxes, and the sheet's own FINAL OUT
        # formula reads E. Mapping E to the afternoon out left it blank whenever
        # the shift ran into the evening. F is that formula — never write to it.
        "columns": {"B": "day", "C": "date", "D": "first_in", "E": "final_out"},
    },
    "sm_asinan": {
        "file": "sm_asinan.xlsx", "sheet": "Timekeeping",
        "blocks": 15,
        "first_block": 3, "stride": 20, "rows_per_block": 16, "name_col": "C",
        # D and E are the day's only in/out boxes, and the sheet's own FINAL OUT
        # formula reads E. Mapping E to the afternoon out left it blank whenever
        # the shift ran into the evening. F is that formula — never write to it.
        "columns": {"B": "day", "C": "date", "D": "first_in", "E": "final_out"},
    },
    "alberto": {
        "file": "alberto.xlsx", "sheet": "Timekeeping",
        "blocks": 176,
        "first_block": 3, "stride": 20, "rows_per_block": 16, "name_col": "C",
        "columns": {"B": "date",
                    "C": "s1_in", "D": "s1_out", "E": "s1_total",
                    "F": "s2_in", "G": "s2_out", "H": "s2_total",
                    "I": "s3_in", "J": "s3_out", "K": "s3_total",
                    "L": "ot_in", "M": "ot_out", "N": "ot_total", "O": "reason"},
    },
    "upson": {
        "file": "upson.xlsx", "sheet": "Sheet1",
        "blocks": 8,
        "first_block": 1, "stride": 20, "rows_per_block": 16, "name_col": "C",
        "columns": {"B": "date", "C": "first_in", "D": "final_out"},
    },
}

# ── clients ───────────────────────────────────────────────────────────────
# `template: None` means the client extracts correctly but has no timekeeping
# workbook of their own yet, so they get CSV export only. That is deliberate —
# inventing a payroll workbook for them would mean inventing their pay rules.
CLIENTS: dict[str, dict] = {
    "alberto":     {"label": "Alberto",              "layout": "timesheet", "template": "alberto"},
    "miniso":      {"label": "Miniso Warehouse",     "layout": "dtr",       "template": "miniso"},
    "tokyo":       {"label": "Tokyo",                "layout": "dtr",       "template": "tokyo"},
    "trinity":     {"label": "Trinity",              "layout": "dtr",       "template": "trinity"},
    "sm_asinan":   {"label": "SM Asinan",            "layout": "dtr",       "template": "sm_asinan"},
    # Oppa Foods' DTRs are the same in/out payroll sheet as Peak 24's, plus a
    # SCHEDULE column. The DTR (morning/afternoon/evening) prompt forced six time
    # columns onto a one-punch-pair sheet and dropped the schedule.
    "oppa_foods":  {"label": "24 Chicken — Oppa Foods", "layout": "inout",  "template": "chicken24"},
    "peak24":      {"label": "24 Chicken — Peak 24", "layout": "inout",     "template": "chicken24"},
    "baobab":      {"label": "Baobab",               "layout": "dtr",       "template": "baobab"},
    "bikezilla":   {"label": "Bikezilla",            "layout": "dtr",       "template": "bikezilla"},
    "upson":       {"label": "Upson",                "layout": "dtr",       "template": "upson"},
    # Awaiting a timekeeping workbook — extraction is configured, export is CSV.
    "booster":     {"label": "Booster",              "layout": "dtr", "template": None},
    "cs":          {"label": "C & S",                "layout": "dtr", "template": None},
    "esn":         {"label": "ESN",                  "layout": "dtr", "template": None},
    "gibi_house":  {"label": "Gibi House",           "layout": "dtr", "template": None},
    "givano":      {"label": "Givano",               "layout": "dtr", "template": None},
    "mhi":         {"label": "MHI",                  "layout": "dtr", "template": None},
    "nectel":      {"label": "Nectel",               "layout": "dtr", "template": None},
    "pied":        {"label": "Pied",                 "layout": "dtr", "template": None},
    "provencal":   {"label": "Provencal",            "layout": "dtr", "template": None},
    "stefano":     {"label": "Stefano",              "layout": "dtr", "template": None},
    "transworld":  {"label": "Transworld",           "layout": "dtr", "template": None},
    "troner":      {"label": "Troner",               "layout": "dtr", "template": None},
    # Not a client: read the form generically, lock nothing.
    "generic":     {"label": "Other / auto-detect",  "layout": None,  "template": None},
}

DEFAULT_CLIENT = "alberto"


def get_client(key: str | None) -> dict:
    return CLIENTS.get((key or "").strip().lower(), CLIENTS[DEFAULT_CLIENT])


def layout_for(key: str | None) -> dict | None:
    return LAYOUTS.get((get_client(key).get("layout") or ""))


def template_for(key: str | None) -> dict | None:
    return TEMPLATES.get((get_client(key).get("template") or ""))


def catalogue() -> list[dict]:
    """What the client picker needs, in display order."""
    out = []
    for key, cfg in CLIENTS.items():
        layout = LAYOUTS.get(cfg.get("layout") or "")
        out.append({
            "key": key,
            "label": cfg["label"],
            "layout": cfg.get("layout"),
            "layout_label": layout["label"] if layout else "Auto-detect",
            "columns": len(layout["headers"]) if layout else 0,
            "has_workbook": bool(cfg.get("template")),
        })
    out.sort(key=lambda c: (c["key"] == "generic", not c["has_workbook"], c["label"].lower()))
    return out
