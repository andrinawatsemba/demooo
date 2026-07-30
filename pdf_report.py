import pandas as pd
from pathlib import Path
from datetime import datetime
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table,
    TableStyle, Image, HRFlowable, PageBreak
)
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.graphics.shapes import Drawing, Rect, String
from reportlab.graphics.charts.piecharts import Pie

import config
import utils

BASE           = Path(__file__).resolve().parent
WAREHOUSE_PATH = BASE / "warehouse" / "warehouse.xlsx"
OUTPUT_FOLDER  = BASE / "output"
LOGO_PATH      = BASE / "nbs_logo.png"
OUTPUT_FOLDER.mkdir(exist_ok=True)

# ── COLORS ───────────────────────────────────────────────────────
# Page chrome (headers, borders, KPI text): blue/white theme.
# The ONLY red pixels anywhere in this PDF live inside nbs_logo.png -
# nothing here draws red directly.
THEME_BLUE     = colors.HexColor(config.THEME["PRIMARY_BLUE"])
THEME_ACCENT   = colors.HexColor(config.THEME["ACCENT_BLUE"])
NBS_WHITE      = colors.white
NBS_DARK_GREY  = colors.HexColor(config.THEME["DARK_GREY"])
NBS_LIGHT_GREY = colors.HexColor(config.THEME["LIGHT_GREY"])
NBS_MID_GREY   = colors.HexColor("#BFBFBF")

# Category colors - looked up BY NAME, not position. This is the fix
# for the old ordering landmine where a positional color list could
# silently scramble if STANDARD_CATEGORIES' order ever changed.
CATEGORY_COLOR_MAP = {
    cat: colors.HexColor(hexval) for cat, hexval in config.CATEGORY_COLOURS.items()
}

HOUR_SEQUENCE = []
for _h in range(24):
    _period = "am" if _h < 12 else "pm"
    _h12 = _h % 12 or 12
    HOUR_SEQUENCE.append(f"{_h12}{_period}")

# ── STYLES ───────────────────────────────────────────────────────
def build_styles():
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="NBSTitle", fontSize=20, textColor=THEME_BLUE,
                               fontName="Helvetica-Bold", alignment=TA_LEFT, spaceAfter=14))
    styles.add(ParagraphStyle(name="NBSSubtitle", fontSize=10, textColor=NBS_DARK_GREY,
                               fontName="Helvetica", alignment=TA_LEFT, spaceAfter=2))
    styles.add(ParagraphStyle(name="NBSSectionHeader", fontSize=11, textColor=NBS_WHITE,
                               fontName="Helvetica-Bold", alignment=TA_LEFT, spaceAfter=4, spaceBefore=8))
    styles.add(ParagraphStyle(name="NBSBodyText", fontSize=8, textColor=NBS_DARK_GREY,
                               fontName="Helvetica", alignment=TA_LEFT, spaceAfter=4))
    styles.add(ParagraphStyle(name="KPIValue", fontSize=16, textColor=THEME_BLUE,
                               fontName="Helvetica-Bold", alignment=TA_CENTER))
    styles.add(ParagraphStyle(name="KPILabel", fontSize=7, textColor=NBS_DARK_GREY,
                               fontName="Helvetica", alignment=TA_CENTER))
    return styles

# ── LOAD DATA ────────────────────────────────────────────────────
def load_latest_week():
    if not WAREHOUSE_PATH.exists():
        print("[ERROR] Warehouse not found. Run pipeline.py first.")
        return None, None, None, None, None

    df = pd.read_excel(WAREHOUSE_PATH, sheet_name="Clean Data")
    corrections = pd.read_excel(WAREHOUSE_PATH, sheet_name="Corrections Log")
    flags = pd.read_excel(WAREHOUSE_PATH, sheet_name="Flags Log")
    try:
        dropped_rows = pd.read_excel(WAREHOUSE_PATH, sheet_name="Dropped Rows Log")
    except Exception:
        dropped_rows = pd.DataFrame(columns=["Row Reference", "AD/SQ Details", "AD/SQ", "Reason"])
    df.columns = df.columns.str.strip()

    if "Week" in df.columns:
        latest_week = df["Week"].iloc[-1]
        df = df[df["Week"] == latest_week]
        return df, latest_week, corrections, flags, dropped_rows
    return df, "Unknown Week", corrections, flags, dropped_rows

# ── KPI TABLE ────────────────────────────────────────────────────
def build_kpi_table(df, styles):
    total_ads  = len(df[df["AD/SQ"] == "AD"])
    total_sqs  = len(df[df["AD/SQ"] == "SQ"])
    total_secs = int(df["Seconds Aired"].sum())

    kpis = [
        (f"{total_ads:,}", "Total ADs Aired"),
        (f"{total_sqs:,}", "Total Squeeze Backs"),
        (f"{total_secs:,}", "Total Airtime (secs)"),
    ]
    data = [[
        Table([[Paragraph(v, styles["KPIValue"])], [Paragraph(l, styles["KPILabel"])]],
              colWidths=[4.5*cm])
        for v, l in kpis
    ]]
    t = Table(data, colWidths=[4.5*cm] * 3)
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), NBS_LIGHT_GREY),
        ("BOX", (0, 0), (-1, -1), 0.5, NBS_MID_GREY),
        ("LINEAFTER", (0, 0), (-2, -1), 0.5, NBS_MID_GREY),
        ("TOPPADDING", (0, 0), (-1, -1), 10),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]))
    return t

# ── PIE CHART ────────────────────────────────────────────────────
def build_pie_chart(df, ad_sq_filter=None):
    scoped = df if ad_sq_filter is None else df[df["AD/SQ"] == ad_sq_filter]
    airtime = (scoped.groupby("Aired Category")["Seconds Aired"].sum()
               .reindex(config.STANDARD_CATEGORIES, fill_value=0))

    values = list(airtime.values)
    pcts = utils.largest_remainder_percentages(values)  # sums to exactly 100.0
    slice_colors = [CATEGORY_COLOR_MAP[cat] for cat in config.STANDARD_CATEGORIES]

    drawing = Drawing(480, 220)
    pie = Pie()
    pie.x, pie.y = 20, 25
    pie.width = pie.height = 175
    pie.data = [int(v) for v in values]
    pie.labels = [""] * len(config.STANDARD_CATEGORIES)
    pie.sideLabels = False

    for i, colour in enumerate(slice_colors):
        pie.slices[i].fillColor = colour
        pie.slices[i].strokeColor = NBS_WHITE
        pie.slices[i].strokeWidth = 1.5

    if sum(values) > 0:
        max_idx = values.index(max(values))
        pie.slices[max_idx].popout = 6

    drawing.add(pie)

    legend_x, legend_y = 225, 195
    for i, (cat, val, pct) in enumerate(zip(config.STANDARD_CATEGORIES, values, pcts)):
        y = legend_y - (i * 30)
        drawing.add(Rect(legend_x, y, 16, 16, fillColor=slice_colors[i],
                          strokeColor=NBS_WHITE, strokeWidth=0.5))
        drawing.add(String(legend_x + 24, y + 4, cat, fontSize=10,
                            fillColor=NBS_DARK_GREY, fontName="Helvetica-Bold"))
        drawing.add(String(legend_x + 150, y + 4, f"{pct}%", fontSize=10,
                            fillColor=THEME_BLUE, fontName="Helvetica-Bold"))
        drawing.add(String(legend_x + 190, y + 4, f"({int(val):,} secs)", fontSize=9,
                            fillColor=colors.HexColor("#666666"), fontName="Helvetica"))
    return drawing

# ── TIME BLOCK BAR CHART (chronological, not alphabetical) ──────
def build_time_block_chart(df):
    blocks_present = [b for b in HOUR_SEQUENCE if b in df["Time Block"].values]
    secs = [int(df[df["Time Block"] == b]["Seconds Aired"].sum()) for b in blocks_present]

    drawing = Drawing(450, 180)
    chart_x, chart_y, chart_w, chart_h = 40, 30, 400, 130
    n = len(blocks_present) or 1
    bar_w = chart_w / n * 0.6
    gap = chart_w / n
    max_val = max(secs) if secs else 1

    for i, (block, val) in enumerate(zip(blocks_present, secs)):
        x = chart_x + i * gap + (gap - bar_w) / 2
        h = (val / max_val) * chart_h if max_val else 0
        drawing.add(Rect(x, chart_y, bar_w, h, fillColor=THEME_BLUE,
                          strokeColor=NBS_WHITE, strokeWidth=0.5))
        drawing.add(String(x + bar_w / 2, chart_y - 12, block, fontSize=6,
                            fillColor=NBS_DARK_GREY, fontName="Helvetica", textAnchor="middle"))

    for step in range(5):
        val = int(max_val * step / 4)
        y = chart_y + (step / 4) * chart_h
        drawing.add(String(chart_x - 5, y - 3, f"{val:,}", fontSize=7,
                            fillColor=NBS_DARK_GREY, fontName="Helvetica", textAnchor="end"))
    return drawing

# ── CATEGORY TABLE (used for AD-only / SQ-only / Combined) ─────
def build_category_table(df, title, ad_sq_filter=None, colwidths=None):
    scoped = df if ad_sq_filter is None else df[df["AD/SQ"] == ad_sq_filter]
    headers = ["Category", "Airtime (secs)", "% of Airtime"]
    rows = [[title, "", ""], headers]

    values = []
    for cat in config.STANDARD_CATEGORIES:
        c = scoped[scoped["Aired Category"] == cat]
        values.append(int(c["Seconds Aired"].sum()))
    pcts = utils.largest_remainder_percentages(values)

    for cat, secs, pct in zip(config.STANDARD_CATEGORIES, values, pcts):
        rows.append([cat, f"{secs:,}", f"{pct}%"])

    total_secs = sum(values)
    rows.append(["TOTAL", f"{total_secs:,}", "100.0%"])

    # colwidths defaults to a single full-width table (7.5+4+3.5=15cm);
    # when placed side-by-side (2 tables in a 18cm content area) pass
    # a narrower set that actually fits in each ~9cm slot - this is
    # the fix for the overlapping/truncated columns you saw.
    colwidths = colwidths or [4*cm, 3.2*cm, 2.6*cm]
    t = Table(rows, colWidths=colwidths)
    t.setStyle(TableStyle([
        # Title row (e.g. "ADs Only" / "SQs Only") - darker than the
        # header row below it so the two are visually distinct, not
        # something you have to guess by reading column contents.
        ("SPAN", (0, 0), (-1, 0)),
        ("BACKGROUND", (0, 0), (-1, 0), NBS_DARK_GREY),
        ("TEXTCOLOR", (0, 0), (-1, 0), NBS_WHITE),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, 0), 10.5),
        ("ALIGN", (0, 0), (-1, 0), "CENTER"),
        ("TOPPADDING", (0, 0), (-1, 0), 7),
        ("BOTTOMPADDING", (0, 0), (-1, 0), 7),

        ("BACKGROUND", (0, 1), (-1, 1), THEME_BLUE),
        ("TEXTCOLOR", (0, 1), (-1, 1), NBS_WHITE),
        ("FONTNAME", (0, 1), (-1, 1), "Helvetica-Bold"),
        ("FONTSIZE", (0, 1), (-1, -1), 9.5),
        ("ALIGN", (0, 1), (-1, -1), "CENTER"),
        ("ALIGN", (0, 2), (0, -1), "LEFT"),
        ("FONTNAME", (0, 2), (0, -1), "Helvetica-Bold"),
        ("ROWBACKGROUNDS", (0, 2), (-1, -2), [NBS_WHITE, NBS_LIGHT_GREY]),
        ("BACKGROUND", (0, -1), (-1, -1), NBS_DARK_GREY),
        ("TEXTCOLOR", (0, -1), (-1, -1), NBS_WHITE),
        ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
        ("GRID", (0, 1), (-1, -1), 0.5, NBS_MID_GREY),
        ("TOPPADDING", (0, 1), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 1), (-1, -1), 6),
    ]))
    return t

def build_summary_table(df):
    headers = ["Day", "Total ADs", "Total SQs", "Total Entries", "Airtime (secs)"]
    rows = [headers]
    days = [d for d in config.DAY_ORDER if d in df["Date"].values]
    g_ads = g_sqs = g_ent = g_secs = 0


    for day in days:
        d = df[df["Date"] == day]
        ads, sqs, ent = len(d[d["AD/SQ"]=="AD"]), len(d[d["AD/SQ"]=="SQ"]), len(d)
        secs = int(d["Seconds Aired"].sum())
        g_ads += ads; g_sqs += sqs; g_ent += ent; g_secs += secs
        rows.append([day, str(ads), str(sqs), str(ent), f"{secs:,}"])

    rows.append(["GRAND TOTAL", str(g_ads), str(g_sqs), str(g_ent), f"{g_secs:,}"])

    t = Table(rows, colWidths=[3.8*cm, 3*cm, 3*cm, 3.4*cm, 3.8*cm])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), THEME_BLUE),
        ("TEXTCOLOR", (0, 0), (-1, 0), NBS_WHITE),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 9.5),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("ALIGN", (0, 1), (0, -1), "LEFT"),
        ("FONTNAME", (0, 1), (0, -1), "Helvetica-Bold"),
        ("ROWBACKGROUNDS", (0, 1), (-1, -2), [NBS_WHITE, NBS_LIGHT_GREY]),
        ("BACKGROUND", (0, -1), (-1, -1), NBS_DARK_GREY),
        ("TEXTCOLOR", (0, -1), (-1, -1), NBS_WHITE),
        ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"),
        ("GRID", (0, 0), (-1, -1), 0.5, NBS_MID_GREY),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))
    return t

# ── PAGE DECORATION — blue band, not red (red is logo-only) ────
def on_page(canvas, doc):
    canvas.saveState()
    w, h = A4
    canvas.setFillColor(THEME_BLUE)
    canvas.rect(0, h - 1*cm, w, 1*cm, fill=True, stroke=False)
    canvas.setFillColor(NBS_MID_GREY)
    canvas.setFont("Helvetica", 7)
    canvas.drawCentredString(w / 2, 0.5*cm,
        f"NBS Ad Tracker  |  Internal Use Only  |  Next Media Services  |  Page {doc.page}")
    canvas.setStrokeColor(THEME_BLUE)
    canvas.setLineWidth(0.5)
    canvas.line(1.5*cm, 1.1*cm, w - 1.5*cm, 1.1*cm)
    canvas.restoreState()

def section_header(title, styles):
    data = [[Paragraph(title, styles["NBSSectionHeader"])]]
    t = Table(data, colWidths=[17*cm])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), THEME_BLUE),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
    ]))
    return t

# ── BUILD PDF ────────────────────────────────────────────────────
def build_pdf(df, week_label, corrections, flags, dropped_rows):
    styles = build_styles()
    safe_week = week_label.replace(" ", "_").replace(".", "-")
    output = OUTPUT_FOLDER / f"NBS_AdReport_{safe_week}.pdf"

    doc = SimpleDocTemplate(str(output), pagesize=A4, rightMargin=1.5*cm,
                             leftMargin=1.5*cm, topMargin=2.5*cm, bottomMargin=2*cm)
    story = []

    if LOGO_PATH.exists():
        logo = Image(str(LOGO_PATH), width=3.5*cm, height=1.8*cm)
        logo.hAlign = "LEFT"
        story.append(logo)
        story.append(Spacer(1, 0.2*cm))

    story.append(HRFlowable(width="100%", thickness=2, color=THEME_BLUE, spaceAfter=0.3*cm))
    story.append(Paragraph("NBS AD TRACKING REPORT", styles["NBSTitle"]))
    story.append(Paragraph(f"Week: {week_label}", styles["NBSSubtitle"]))
    story.append(Spacer(1, 0.1*cm))
    story.append(Paragraph(
        f"Generated: {datetime.now().strftime('%d %B %Y, %H:%M')}  |  Next Media Services  |  Internal Use Only",
        styles["NBSSubtitle"]))
    story.append(HRFlowable(width="100%", thickness=1, color=NBS_MID_GREY, spaceAfter=0.4*cm))

    story.append(section_header("EXECUTIVE SUMMARY", styles))
    story.append(Spacer(1, 0.3*cm))
    story.append(build_kpi_table(df, styles))
    story.append(Spacer(1, 0.4*cm))

    story.append(section_header("SECTION 1 — AIRTIME BY CATEGORY (COMBINED)", styles))
    story.append(Spacer(1, 0.3*cm))
    story.append(build_pie_chart(df))
    story.append(Spacer(1, 0.3*cm))
    story.append(build_category_table(df, "COMBINED (AD + SQ)"))
    story.append(Spacer(1, 0.3*cm))

    story.append(PageBreak())
    if LOGO_PATH.exists():
        logo2 = Image(str(LOGO_PATH), width=3*cm, height=1.5*cm)
        logo2.hAlign = "LEFT"
        story.append(logo2)
        story.append(Spacer(1, 0.2*cm))
    story.append(HRFlowable(width="100%", thickness=2, color=THEME_BLUE, spaceAfter=0.3*cm))

    story.append(section_header("SECTION 2 — CATEGORY BREAKDOWN (ADs vs SQs)", styles))
    story.append(Spacer(1, 0.3*cm))
    narrow_widths = [3.6*cm, 3*cm, 2.4*cm]  # sums to 9cm - actually fits the side-by-side slot
    col1 = build_category_table(df, "ADS ONLY", ad_sq_filter="AD", colwidths=narrow_widths)
    col2 = build_category_table(df, "SQS ONLY", ad_sq_filter="SQ", colwidths=narrow_widths)
    side_by_side = Table([[col1, col2]], colWidths=[9*cm, 9*cm])
    side_by_side.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP")]))
    story.append(side_by_side)
    story.append(Spacer(1, 0.4*cm))

    story.append(section_header("SECTION 3 — AIRTIME PER DAY", styles))
    story.append(Spacer(1, 0.3*cm))
    story.append(build_summary_table(df))
    story.append(Spacer(1, 0.4*cm))

    story.append(section_header("SECTION 4 — AIRTIME BY TIME BLOCK", styles))
    story.append(Spacer(1, 0.3*cm))
    story.append(build_time_block_chart(df))

    doc.build(story, onFirstPage=on_page, onLaterPages=on_page)
    print(f"[OK] PDF saved to: {output}")
    return output

# ── MAIN ─────────────────────────────────────────────────────────
if __name__ == "__main__":
    print("\n" + "="*50)
    print("PDF REPORT BUILDER")
    print("="*50)

    result = load_latest_week()
    if result[0] is None:
        exit()
    df, week_label, corrections, flags, dropped_rows = result
    build_pdf(df, week_label, corrections, flags, dropped_rows)
    print("="*50)
