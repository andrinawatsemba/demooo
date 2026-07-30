"""
Shared constants for the whole app.

Why this file exists: in the old codebase, STANDARD_CATEGORIES was
copy-pasted into pipeline.py, app.py, report.py, and pdf_report.py
(4 separate lists), plus a 5th copy inside an Excel dropdown-validation
string. When "Strategic" was retired, only some copies got updated.
Same problem with colors - pdf_report.py kept a positionally-matched
color list whose order didn't match the other files' category order.

Fix: every file imports from HERE. One list, one color dict. Change a
category once, it's correct everywhere.
"""

# ── CATEGORIES ───────────────────────────────────────────────────
# "Strategic" was retired and folded into "Commercial" - it is
# deliberately NOT in this list. If it ever reappears in old data,
# validate() will flag it as non-standard rather than silently
# accepting it.
STANDARD_CATEGORIES = [
    "Commercial", "Internal", "Trade Marketing", "Swap", "Franchise"
]

# ── CATEGORY COLORS (pie/bar charts) ────────────────────────────
# Keyed by NAME, not position - this is the fix for the pdf_report.py
# ordering landmine. Order in STANDARD_CATEGORIES can now change
# freely without scrambling chart colors.
# Warm palette (per reference image) - red is deliberately excluded
# here since red is reserved for the logo only.
CATEGORY_COLOURS = {
    "Commercial":      "#E49C30",  # amber/orange
    "Internal":        "#6C6090",  # purple
    "Trade Marketing": "#303C60",  # dark navy
    "Swap":            "#3C9CC0",  # teal-blue
    "Franchise":       "#0090F0",  # primary theme blue
}

# ── THEME COLORS (page chrome: headers, backgrounds, borders) ──
# Separate from CATEGORY_COLOURS above - this is the overall look,
# not per-category chart coloring. Red confined to the logo only.
THEME = {
    "PRIMARY_BLUE": "#0090F0",
    "ACCENT_BLUE":  "#00C8F0",
    "WHITE":        "#FFFFFF",
    "LOGO_RED":     "#E30613",  # unchanged, logo only
    "DARK_GREY":    "#404040",
    "LIGHT_GREY":   "#F2F2F2",
}

# ── TIMING ───────────────────────────────────────────────────────
# AD changed from 45s to 40s (confirmed). SQ unchanged at 30s.
AD_SECONDS = 40
SQ_SECONDS = 30

# ── NON-AD / OPERATIONAL TEXT ───────────────────────────────────
# Fixed from substring match to exact match (see pipeline.py) -
# this list is compared against the FULL cleaned cell value now,
# not "is this phrase contained anywhere in the cell".
NON_AD_PHRASES = [
    "end of shift", "shift sign", "signed by", "approved by",
    "scheduler", "programs manager", "chief programs officer",
    "total per hour", "total per block",
    "note:", "morning breeze", "sunrise news", "lunchtime sports",
    "breakfast meeting", "topical discussion",
]
# NOTE: this list is unchanged from the original. Words like
# "TONIGHT" and "PROGRAM LINEUP" show up a lot in the raw file with
# blank time cells and look like they might be operational/program
# text too - but that's a business call, not something to guess at
# in code. Flagged separately for you to confirm.

# ── DAYS ─────────────────────────────────────────────────────────
DAY_ORDER = [
    "Monday", "Tuesday", "Wednesday",
    "Thursday", "Friday", "Saturday", "Sunday"
]

DAY_ALIASES = {
    "mon": "Monday", "monday": "Monday",
    "tue": "Tuesday", "tues": "Tuesday", "tuesday": "Tuesday",
    "wed": "Wednesday", "wednesday": "Wednesday",
    "thu": "Thursday", "thur": "Thursday",
    "thurs": "Thursday", "thursday": "Thursday",
    "fri": "Friday", "friday": "Friday",
    "sat": "Saturday", "saturday": "Saturday",
    "sun": "Sunday", "sunday": "Sunday"
}

HEADER_ALIASES = {
    "time aired": "TIME AIRED",
    "time air": "TIME AIRED",
    "aired": "TIME AIRED",
    "squeeze backs": "SQUEEZE BACKS",
    "squeeze back": "SQUEEZE BACKS",
    "sq backs": "SQUEEZE BACKS",
    "sq bk": "SQUEEZE BACKS",
    "scrolls": "SQUEEZE BACKS",
    "squeeze backs/scrolls": "SQUEEZE BACKS",
}
