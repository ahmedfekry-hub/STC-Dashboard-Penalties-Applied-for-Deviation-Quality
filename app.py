import io
import math
import os
from datetime import datetime
from textwrap import shorten

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    Flowable,
    Image,
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)

# =========================
# Theme
# =========================
st.set_page_config(page_title="STC Quality Executive Dashboard", layout="wide", initial_sidebar_state="expanded")

PURPLE = "#5A0AA2"
PURPLE_2 = "#7C1DEB"
TEAL = "#2EC4D3"
GREEN = "#18B885"
PINK = "#EF476F"
ORANGE = "#FF6B3D"
YELLOW = "#F7B500"
TEXT = "#241B35"
MUTED = "#766C8A"
BORDER = "#E6DDF3"
CARD = "#FFFFFF"
BG = "#FBFAFD"
SOFT_PURPLE = "#F2ECFA"
SOFT_GREEN = "#E9FAF4"
SOFT_PINK = "#FFE9EF"
SOFT_ORANGE = "#FFF0E8"


def hex_color(value: str):
    """Safe ReportLab HexColor. Always prepend # to avoid ValueError."""
    value = str(value).strip()
    if not value.startswith("#"):
        value = "#" + value
    return colors.HexColor(value)


# =========================
# Arabic support for PDF
# =========================
PDF_FONT = "Helvetica"
PDF_FONT_BOLD = "Helvetica-Bold"
for font_path in [
    "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
    "/usr/share/fonts/truetype/dejavu/DejaVuSansCondensed.ttf",
    "/usr/share/fonts/truetype/liberation2/LiberationSans-Regular.ttf",
]:
    if os.path.exists(font_path):
        try:
            pdfmetrics.registerFont(TTFont("DejaVuSans", font_path))
            PDF_FONT = "DejaVuSans"
            PDF_FONT_BOLD = "DejaVuSans"
            break
        except Exception:
            pass

try:
    import arabic_reshaper
    from bidi.algorithm import get_display

    def pdf_text(value):
        text = "" if pd.isna(value) else str(value)
        # Reshape only if Arabic characters exist.
        if any("\u0600" <= ch <= "\u06FF" for ch in text):
            try:
                return get_display(arabic_reshaper.reshape(text))
            except Exception:
                return text
        return text
except Exception:
    def pdf_text(value):
        return "" if pd.isna(value) else str(value)


st.markdown(
    f"""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700;800&display=swap');
        html, body, [class*="css"] {{font-family: 'Inter', sans-serif;}}
        .main .block-container {{padding-top: 1.0rem; padding-bottom: 2rem; max-width: 1540px;}}
        .hero {{
            background: linear-gradient(120deg, {PURPLE}, {PURPLE_2});
            color: white;
            border-radius: 0 0 22px 22px;
            padding: 26px 28px;
            margin-bottom: 20px;
            box-shadow: 0 18px 45px rgba(90,10,162,0.18);
            animation: fadeInDown .65s ease-in-out;
        }}
        .hero h1 {{margin: 0; font-size: 34px; font-weight: 800;}}
        .hero p {{margin-top: 12px; opacity: .95; font-size: 14px;}}
        .metric-card {{
            background: {CARD};
            border: 1px solid {BORDER};
            border-radius: 14px;
            padding: 18px 18px 16px 18px;
            min-height: 112px;
            box-shadow: 0 8px 24px rgba(50, 30, 80, 0.06);
            animation: fadeInUp .6s ease;
        }}
        .metric-label {{font-size: 13px; color: {TEXT}; font-weight: 700;}}
        .metric-value {{font-size: 32px; line-height: 1; font-weight: 800; margin-top: 12px;}}
        .metric-note {{font-size: 12px; color: {MUTED}; margin-top: 12px;}}
        .section-title {{font-size: 23px; font-weight: 800; color: {TEXT}; margin-top: 18px; margin-bottom: 10px;}}
        .small-caption {{font-size: 12px; color: {MUTED};}}
        .stDataFrame {{border-radius: 14px; overflow: hidden;}}
        div[data-testid="stSidebar"] {{background-color: #F7F2FC;}}
        @keyframes fadeInDown {{from {{opacity:0; transform: translateY(-15px);}} to {{opacity:1; transform: translateY(0);}}}}
        @keyframes fadeInUp {{from {{opacity:0; transform: translateY(18px);}} to {{opacity:1; transform: translateY(0);}}}}
    </style>
    """,
    unsafe_allow_html=True,
)


# =========================
# Data helpers
# =========================
def _normalize_text(s: pd.Series) -> pd.Series:
    return s.fillna("").astype(str).str.strip()


def _clean_wo(value):
    if pd.isna(value):
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip().replace(".0", "")


def _clean_bool_series(s: pd.Series) -> pd.Series:
    return s.fillna("").astype(str).str.strip().str.upper().isin(["Y", "YES", "TRUE", "1", "APPLIED", "نعم"])


@st.cache_data(show_spinner=False)
def load_data(file_obj=None) -> pd.DataFrame:
    if file_obj is not None:
        raw = pd.read_excel(file_obj)
    else:
        raw = pd.read_excel("Deviation.xlsx")

    raw.columns = [str(c).strip() for c in raw.columns]
    required = ["District", "WorkOrderNum", "DeviationName"]
    missing = [c for c in required if c not in raw.columns]
    if missing:
        raise ValueError(f"Missing mandatory columns: {missing}")

    df = raw.copy()
    df["District"] = _normalize_text(df.get("District", pd.Series(index=df.index))).str.upper()
    df["WorkOrderNum"] = df.get("WorkOrderNum", pd.Series(index=df.index)).apply(_clean_wo)
    df["DeviationName"] = _normalize_text(df.get("DeviationName", pd.Series(index=df.index)))
    df["Category"] = _normalize_text(df.get("Category", pd.Series(index=df.index)))
    df["SubCategory"] = _normalize_text(df.get("SubCategory", pd.Series(index=df.index)))
    df["Designation"] = _normalize_text(df.get("Designation", pd.Series(index=df.index)))
    df["DeviationStatus"] = _normalize_text(df.get("DeviationStatus", pd.Series(index=df.index)))
    df["Civil Foreman"] = _normalize_text(df.get("Civil Foreman", pd.Series(index=df.index))).replace("", "—")
    df["Inspector"] = _normalize_text(df.get("Inspector", pd.Series(index=df.index))).replace("", "—")
    df["IsPenalty"] = _normalize_text(df.get("IsPenalty", pd.Series(index=df.index)))
    df["ServiceAffecting"] = _normalize_text(df.get("ServiceAffecting", pd.Series(index=df.index)))
    df["Expected Penalties"] = _normalize_text(df.get("Expected Penalties", pd.Series(index=df.index)))

    df["PenaltyApplied"] = _clean_bool_series(df["IsPenalty"])
    df["ServiceAffectingFlag"] = _clean_bool_series(df["ServiceAffecting"])
    # Expected Penalties means: penalty should have applied but was waived/cancelled.
    df["ExpectedPenaltyWaived"] = _clean_bool_series(df["Expected Penalties"])
    df["NoPenaltyApplied"] = ~df["PenaltyApplied"]

    def classify(row) -> str:
        designation = str(row.get("Designation", "")).strip().upper()
        if designation in ["CIVIL", "FIBER", "SAFETY"]:
            return designation.title() if designation != "SAFETY" else "Safety"
        combined = " ".join([
            str(row.get("Designation", "")),
            str(row.get("Category", "")),
            str(row.get("SubCategory", "")),
            str(row.get("DeviationName", "")),
        ]).upper()
        if "SAFETY" in combined or "H&S" in combined or "WORKER" in combined or "ID BADGE" in combined or "UNIFORM" in combined:
            return "Safety"
        if "FIBER" in combined or "CABLE" in combined or "SPLIC" in combined or "LABEL" in combined:
            return "Fiber"
        if "CIVIL" in combined or "TRENCH" in combined or "DUCT" in combined or "BACKFILL" in combined or "MANHOLE" in combined or "HANDHOLE" in combined or "CABINET" in combined or "MATERIAL" in combined:
            return "Civil"
        return "Other"

    df["TeamClassification"] = df.apply(classify, axis=1)
    # Mutually exclusive penalty status for charts/PDF.
    df["PenaltyStatus"] = "Other No Penalty"
    df.loc[df["ExpectedPenaltyWaived"] & ~df["PenaltyApplied"], "PenaltyStatus"] = "Expected Penalties Waived"
    df.loc[df["PenaltyApplied"], "PenaltyStatus"] = "Penalty Applied"
    return df


def pct(n, d):
    return 0.0 if d == 0 else round((n / d) * 100, 1)


def comma_join_unique(values, limit=3):
    vals = [v for v in pd.unique(pd.Series(values).dropna()) if str(v).strip() not in ["", "—", "nan"]]
    if not vals:
        return "—"
    text = " / ".join(map(str, vals[:limit]))
    if len(vals) > limit:
        text += f" +{len(vals) - limit}"
    return text


def top_deviation_nature(group: pd.DataFrame, limit=3) -> str:
    vc = group["DeviationName"].value_counts().head(limit)
    if vc.empty:
        return "—"
    return " | ".join([f"{shorten(str(k), width=42, placeholder='...')} ({int(v)})" for k, v in vc.items()])


def build_wo_summary(df: pd.DataFrame) -> pd.DataFrame:
    rows = []
    total_rows = len(df)
    for wo, g in df.groupby("WorkOrderNum", dropna=False):
        penalty_applied = int(g["PenaltyApplied"].sum())
        expected_waived = int((g["ExpectedPenaltyWaived"] & ~g["PenaltyApplied"]).sum())
        no_penalty = int(g["NoPenaltyApplied"].sum())
        rows.append({
            "Work Order": wo,
            "District": comma_join_unique(g["District"], limit=2),
            "Civil Foreman": comma_join_unique(g["Civil Foreman"], limit=3),
            "Inspector": comma_join_unique(g["Inspector"], limit=3),
            "Total Deviations": int(len(g)),
            "Penalty Applied": penalty_applied,
            "No Penalty Applied": no_penalty,
            "Expected Penalties Waived": expected_waived,
            "Service Affecting": int(g["ServiceAffectingFlag"].sum()),
            "Civil": int((g["TeamClassification"] == "Civil").sum()),
            "Fiber": int((g["TeamClassification"] == "Fiber").sum()),
            "Safety": int((g["TeamClassification"] == "Safety").sum()),
            "Other": int((g["TeamClassification"] == "Other").sum()),
            "% of Total": pct(len(g), total_rows),
            "Top Deviation Nature": top_deviation_nature(g, limit=3),
        })
    return pd.DataFrame(rows).sort_values("Total Deviations", ascending=False).reset_index(drop=True)


def build_deviation_nature(df: pd.DataFrame) -> pd.DataFrame:
    base = (
        df.groupby(["WorkOrderNum", "District", "Civil Foreman", "Inspector", "TeamClassification", "Category", "SubCategory", "DeviationName"], dropna=False)
        .agg(
            Total=("DeviationName", "size"),
            Penalty_Applied=("PenaltyApplied", "sum"),
            No_Penalty=("NoPenaltyApplied", "sum"),
            Expected_Penalties_Waived=("ExpectedPenaltyWaived", "sum"),
            Service_Affecting=("ServiceAffectingFlag", "sum"),
        )
        .reset_index()
        .sort_values(["Total", "Penalty_Applied"], ascending=False)
    )
    return base.rename(columns={"WorkOrderNum": "Work Order", "TeamClassification": "Classification"})


def color_metric(label, value, note, color):
    st.markdown(
        f"""
        <div class="metric-card">
            <div class="metric-label">{label}</div>
            <div class="metric-value" style="color:{color};">{value}</div>
            <div class="metric-note">{note}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


# =========================
# PDF custom donut chart
# =========================
class DonutChart(Flowable):
    def __init__(self, values, labels, chart_colors, width=11.5*cm, height=6.8*cm):
        super().__init__()
        self.values = values
        self.labels = labels
        self.chart_colors = chart_colors
        self.width = width
        self.height = height

    def draw(self):
        total = sum(self.values) or 1
        c = self.canv
        cx = self.width * 0.44
        cy = self.height * 0.50
        r = min(self.width, self.height) * 0.30
        start = 90
        bbox = (cx - r, cy - r, cx + r, cy + r)

        # Wedges
        for value, label, col in zip(self.values, self.labels, self.chart_colors):
            extent = -360 * (value / total)
            c.setFillColor(hex_color(col))
            c.setStrokeColor(colors.white)
            c.wedge(*bbox, start, extent, fill=1, stroke=1)
            start += extent

        # Center hole
        c.setFillColor(colors.white)
        c.circle(cx, cy, r * 0.55, fill=1, stroke=0)
        c.setFillColor(hex_color(TEXT))
        c.setFont(PDF_FONT_BOLD, 10)
        c.drawCentredString(cx, cy + 4, str(int(total)))
        c.setFont(PDF_FONT, 7)
        c.setFillColor(hex_color(MUTED))
        c.drawCentredString(cx, cy - 8, "Total Deviations")

        # Labels with connector lines outside
        start = 90
        for value, label, col in zip(self.values, self.labels, self.chart_colors):
            if value <= 0:
                continue
            extent = -360 * (value / total)
            mid = math.radians(start + extent / 2)
            x1 = cx + math.cos(mid) * r * 0.80
            y1 = cy + math.sin(mid) * r * 0.80
            x2 = cx + math.cos(mid) * r * 1.12
            y2 = cy + math.sin(mid) * r * 1.12
            label_right = math.cos(mid) >= 0
            x3 = x2 + (1.2 * cm if label_right else -1.2 * cm)
            y3 = y2
            c.setStrokeColor(hex_color(col))
            c.setLineWidth(1.1)
            c.line(x1, y1, x2, y2)
            c.line(x2, y2, x3, y3)
            c.setFillColor(hex_color(TEXT))
            c.setFont(PDF_FONT_BOLD, 7.2)
            pct_text = f"{pct(value, total)}%"
            text = f"{label}: {int(value)} ({pct_text})"
            if label_right:
                c.drawString(x3 + 3, y3 - 2, text)
            else:
                c.drawRightString(x3 - 3, y3 - 2, text)
            start += extent


# =========================
# PDF Builder
# =========================
def p(value, style):
    return Paragraph(pdf_text(value), style)


def make_pdf(df: pd.DataFrame, wo_summary: pd.DataFrame, nature: pd.DataFrame) -> bytes:
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=landscape(A4),
        rightMargin=0.7 * cm,
        leftMargin=0.7 * cm,
        topMargin=0.65 * cm,
        bottomMargin=0.65 * cm,
    )
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="TitlePurple", parent=styles["Title"], fontName=PDF_FONT_BOLD, textColor=hex_color(PURPLE), fontSize=22, leading=25, alignment=TA_CENTER, spaceAfter=6))
    styles.add(ParagraphStyle(name="Sub", parent=styles["Normal"], fontName=PDF_FONT, textColor=hex_color(MUTED), fontSize=9.2, alignment=TA_CENTER, spaceAfter=8))
    styles.add(ParagraphStyle(name="Section", parent=styles["Heading2"], fontName=PDF_FONT_BOLD, textColor=hex_color(TEXT), fontSize=14, leading=16, spaceBefore=8, spaceAfter=5))
    styles.add(ParagraphStyle(name="CardLabel", parent=styles["Normal"], fontName=PDF_FONT, textColor=hex_color(MUTED), fontSize=7.6, leading=9, alignment=TA_CENTER))
    styles.add(ParagraphStyle(name="CardValue", parent=styles["Normal"], fontName=PDF_FONT_BOLD, textColor=hex_color(PURPLE), fontSize=16, leading=18, alignment=TA_CENTER))
    styles.add(ParagraphStyle(name="HeaderCell", parent=styles["Normal"], fontName=PDF_FONT_BOLD, textColor=colors.white, fontSize=6.4, leading=7.3, alignment=TA_CENTER))
    styles.add(ParagraphStyle(name="Cell", parent=styles["Normal"], fontName=PDF_FONT, textColor=hex_color(TEXT), fontSize=6.15, leading=7.2, alignment=TA_CENTER))
    styles.add(ParagraphStyle(name="CellLeft", parent=styles["Normal"], fontName=PDF_FONT, textColor=hex_color(TEXT), fontSize=6.0, leading=7.1, alignment=TA_LEFT))
    styles.add(ParagraphStyle(name="Note", parent=styles["Normal"], fontName=PDF_FONT, textColor=hex_color(MUTED), fontSize=7, leading=8))

    total = len(df)
    penalty = int(df["PenaltyApplied"].sum())
    waived = int((df["ExpectedPenaltyWaived"] & ~df["PenaltyApplied"]).sum())
    no_penalty = int(df["NoPenaltyApplied"].sum())
    other_no_penalty = max(no_penalty - waived, 0)
    service = int(df["ServiceAffectingFlag"].sum())
    unique_wo = df["WorkOrderNum"].nunique()

    story = []
    story.append(p("STC Quality Executive Dashboard", styles["TitlePurple"]))
    story.append(p("Deviation WO drilldown - penalty applied vs waived/not applied - Civil / Fiber / Safety - Foreman & Inspector accountability", styles["Sub"]))

    card_data = [
        [p("Total Deviations", styles["CardLabel"]), p("Unique WOs", styles["CardLabel"]), p("Penalty Applied", styles["CardLabel"]), p("No Penalty Applied", styles["CardLabel"]), p("Expected Penalties Waived", styles["CardLabel"]), p("Service Affecting", styles["CardLabel"])],
        [p(f"{total:,}", styles["CardValue"]), p(f"{unique_wo:,}", styles["CardValue"]), p(f"{penalty:,}", styles["CardValue"]), p(f"{no_penalty:,}", styles["CardValue"]), p(f"{waived:,}", styles["CardValue"]), p(f"{service:,}", styles["CardValue"])],
        [p("All filtered records", styles["CardLabel"]), p("Impacted Work Orders", styles["CardLabel"]), p(f"{pct(penalty,total)}%", styles["CardLabel"]), p(f"{pct(no_penalty,total)}%", styles["CardLabel"]), p("Should be penalized but cancelled", styles["CardLabel"]), p(f"{pct(service,total)}%", styles["CardLabel"])]
    ]
    card_table = Table(card_data, colWidths=[4.55*cm]*6, rowHeights=[0.55*cm, 0.72*cm, 0.48*cm], hAlign="CENTER")
    card_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), hex_color(SOFT_PURPLE)),
        ("BOX", (0, 0), (-1, -1), 0.5, hex_color(BORDER)),
        ("GRID", (0, 0), (-1, -1), 0.3, hex_color("#DED2EE")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ]))
    story.append(card_table)
    story.append(Spacer(1, 0.18 * cm))

    story.append(p("Penalty Applied vs Waived / Not Applied", styles["Section"]))
    donut = DonutChart(
        values=[penalty, waived, other_no_penalty],
        labels=["Penalty Applied", "Expected Penalties Waived", "Other No Penalty"],
        chart_colors=[PINK, ORANGE, GREEN],
        width=20.5*cm,
        height=6.6*cm,
    )
    chart_note = p("Expected Penalties Waived = deviations where penalty was expected to be applied, but was cancelled/not applied.", styles["Note"])
    row_chart = Table([[donut, chart_note]], colWidths=[20.8*cm, 6.3*cm], hAlign="CENTER")
    row_chart.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("BACKGROUND", (0, 0), (-1, -1), colors.white),
        ("BOX", (0, 0), (-1, -1), 0.4, hex_color(BORDER)),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
    ]))
    story.append(row_chart)

    story.append(p("Board Work Order Executive Summary", styles["Section"]))
    summary_cols = [
        "Work Order", "District", "Civil Foreman", "Inspector", "Total Deviations", "Penalty Applied", "No Penalty Applied",
        "Expected Penalties Waived", "Service Affecting", "Civil", "Fiber", "Safety", "Other", "% of Total", "Top Deviation Nature",
    ]
    if wo_summary.empty:
        summary_table_data = [[p(c, styles["HeaderCell"]) for c in summary_cols], [p("No data available", styles["Cell"])] + [""]*(len(summary_cols)-1)]
    else:
        summary_table_data = [[p(c, styles["HeaderCell"]) for c in summary_cols]]
        for _, r in wo_summary[summary_cols].iterrows():
            row = []
            for c_name in summary_cols:
                style = styles["CellLeft"] if c_name in ["Civil Foreman", "Inspector", "Top Deviation Nature"] else styles["Cell"]
                row.append(p(str(r[c_name]), style))
            summary_table_data.append(row)

    summary_widths = [1.85*cm, 1.25*cm, 2.0*cm, 1.85*cm, 0.95*cm, 1.05*cm, 1.1*cm, 1.28*cm, 0.95*cm, 0.72*cm, 0.72*cm, 0.72*cm, 0.72*cm, 0.72*cm, 11.0*cm]
    summary_table = Table(summary_table_data, colWidths=summary_widths, repeatRows=1, hAlign="CENTER")
    summary_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), hex_color(PURPLE)),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("GRID", (0, 0), (-1, -1), 0.25, hex_color("#DDD3EA")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, hex_color("#FAF7FE")]),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (0, 1), (-1, -1), "CENTER"),
        ("TOPPADDING", (0, 0), (-1, -1), 2.4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2.4),
        ("BACKGROUND", (7, 1), (7, -1), hex_color(SOFT_ORANGE)),
        ("BACKGROUND", (5, 1), (5, -1), hex_color(SOFT_PINK)),
        ("BACKGROUND", (6, 1), (6, -1), hex_color(SOFT_GREEN)),
    ]))
    story.append(summary_table)

    story.append(PageBreak())
    story.append(p("Detailed Deviation Nature by Work Order", styles["TitlePurple"]))
    story.append(p("Detail view for Board review: WO, District, Civil Foreman, Inspector, classification, deviation nature and penalty status.", styles["Sub"]))

    detail_cols = [
        "Work Order", "District", "Civil Foreman", "Inspector", "Classification", "Category", "SubCategory", "DeviationName",
        "Total", "Penalty_Applied", "No_Penalty", "Expected_Penalties_Waived", "Service_Affecting",
    ]
    detail_df = nature[detail_cols].copy() if not nature.empty else pd.DataFrame(columns=detail_cols)
    detail_table_data = [[p(str(c).replace("_", " "), styles["HeaderCell"]) for c in detail_cols]]
    for _, r in detail_df.iterrows():
        detail_row = []
        for c_name in detail_cols:
            style = styles["CellLeft"] if c_name in ["Civil Foreman", "Inspector", "DeviationName", "Category", "SubCategory"] else styles["Cell"]
            detail_row.append(p(str(r[c_name]), style))
        detail_table_data.append(detail_row)

    detail_widths = [1.9*cm, 1.18*cm, 2.0*cm, 1.75*cm, 1.25*cm, 1.5*cm, 1.75*cm, 11.0*cm, 0.65*cm, 0.88*cm, 0.85*cm, 1.16*cm, 0.85*cm]
    detail_table = Table(detail_table_data, colWidths=detail_widths, repeatRows=1, hAlign="CENTER")
    detail_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), hex_color(PURPLE)),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("GRID", (0, 0), (-1, -1), 0.25, hex_color("#DDD3EA")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, hex_color("#FAF7FE")]),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (8, 1), (-1, -1), "CENTER"),
        ("TOPPADDING", (0, 0), (-1, -1), 2.4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2.4),
        ("BACKGROUND", (11, 1), (11, -1), hex_color(SOFT_ORANGE)),
    ]))
    story.append(detail_table)

    def _footer(canvas, doc):
        canvas.saveState()
        canvas.setFont(PDF_FONT, 7)
        canvas.setFillColor(hex_color(MUTED))
        canvas.drawString(0.7 * cm, 0.35 * cm, f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')} | Source: Deviation.xlsx")
        canvas.drawRightString(28.8 * cm, 0.35 * cm, f"Page {doc.page}")
        canvas.restoreState()

    doc.build(story, onFirstPage=_footer, onLaterPages=_footer)
    buffer.seek(0)
    return buffer.read()


# =========================
# Sidebar / Data
# =========================
with st.sidebar:
    st.markdown("### Source File")
    uploaded = st.file_uploader("Upload updated Deviation.xlsx", type=["xlsx"], help="Optional. If not uploaded, app uses the included Deviation.xlsx file.")
    st.markdown("### Filters")

try:
    df = load_data(uploaded)
except Exception as exc:
    st.error(f"Unable to read Deviation.xlsx: {exc}")
    st.stop()

with st.sidebar:
    districts = sorted(df["District"].dropna().unique())
    wos = sorted(df["WorkOrderNum"].dropna().unique())
    inspectors = sorted([x for x in df["Inspector"].dropna().unique() if x != "—"])
    foremen = sorted([x for x in df["Civil Foreman"].dropna().unique() if x != "—"])
    selected_district = st.multiselect("District", districts)
    selected_wo = st.multiselect("Work Order", wos)
    selected_foreman = st.multiselect("Civil Foreman", foremen)
    selected_inspector = st.multiselect("Inspector", inspectors)
    selected_class = st.multiselect("Classification", ["Civil", "Fiber", "Safety", "Other"])
    penalty_view = st.radio("Penalty View", ["All", "Penalty Applied", "No Penalty Applied", "Expected Penalties Waived"], horizontal=False)

filtered = df.copy()
if selected_district:
    filtered = filtered[filtered["District"].isin(selected_district)]
if selected_wo:
    filtered = filtered[filtered["WorkOrderNum"].isin(selected_wo)]
if selected_foreman:
    filtered = filtered[filtered["Civil Foreman"].isin(selected_foreman)]
if selected_inspector:
    filtered = filtered[filtered["Inspector"].isin(selected_inspector)]
if selected_class:
    filtered = filtered[filtered["TeamClassification"].isin(selected_class)]
if penalty_view == "Penalty Applied":
    filtered = filtered[filtered["PenaltyApplied"]]
elif penalty_view == "No Penalty Applied":
    filtered = filtered[filtered["NoPenaltyApplied"]]
elif penalty_view == "Expected Penalties Waived":
    filtered = filtered[filtered["ExpectedPenaltyWaived"] & ~filtered["PenaltyApplied"]]

wo_summary = build_wo_summary(filtered)
nature = build_deviation_nature(filtered)

# =========================
# Dashboard UI
# =========================
st.markdown(
    """
    <div class="hero">
        <h1>STC Quality Executive Dashboard</h1>
        <p>Deviation analysis • Work Order performance • penalty applied vs waived/not applied • Civil / Fiber / Safety classification • Foreman & Inspector accountability</p>
    </div>
    """,
    unsafe_allow_html=True,
)

total = len(filtered)
penalty = int(filtered["PenaltyApplied"].sum())
no_penalty = int(filtered["NoPenaltyApplied"].sum())
waived = int((filtered["ExpectedPenaltyWaived"] & ~filtered["PenaltyApplied"]).sum())
other_no_penalty = max(no_penalty - waived, 0)
service = int(filtered["ServiceAffectingFlag"].sum())
unique_wo = filtered["WorkOrderNum"].nunique()

m1, m2, m3, m4, m5, m6 = st.columns(6)
with m1:
    color_metric("Total Deviations", f"{total:,}", "All records after filters", PURPLE)
with m2:
    color_metric("Unique WOs", f"{unique_wo:,}", "Impacted Work Orders", TEAL)
with m3:
    color_metric("Penalty Applied", f"{penalty:,}", f"{pct(penalty,total)}% of deviations", PINK)
with m4:
    color_metric("No Penalty Applied", f"{no_penalty:,}", f"{pct(no_penalty,total)}% of deviations", GREEN)
with m5:
    color_metric("Expected Penalties Waived", f"{waived:,}", "Should be penalized but cancelled", ORANGE)
with m6:
    color_metric("Service Affecting", f"{service:,}", f"{pct(service,total)}% of deviations", YELLOW)

st.markdown('<div class="section-title">Executive Visualizations</div>', unsafe_allow_html=True)

c1, c2 = st.columns([1.35, 1])
with c1:
    st.subheader("Top Work Orders by Deviation Count")
    if not wo_summary.empty:
        fig_wo = px.bar(
            wo_summary.sort_values("Total Deviations"),
            x="Total Deviations",
            y="Work Order",
            color="District",
            orientation="h",
            text="Total Deviations",
            custom_data=["Civil Foreman", "Inspector", "Penalty Applied", "Expected Penalties Waived", "Safety", "Civil", "Fiber"],
            color_discrete_sequence=[PURPLE, TEAL, ORANGE, PINK, GREEN],
        )
        fig_wo.update_traces(
            textposition="outside",
            hovertemplate=(
                "<b>WO:</b> %{y}<br>"
                "<b>Total:</b> %{x}<br>"
                "<b>Foreman:</b> %{customdata[0]}<br>"
                "<b>Inspector:</b> %{customdata[1]}<br>"
                "<b>Penalty Applied:</b> %{customdata[2]}<br>"
                "<b>Expected Penalties Waived:</b> %{customdata[3]}<br>"
                "<b>Safety / Civil / Fiber:</b> %{customdata[4]} / %{customdata[5]} / %{customdata[6]}<extra></extra>"
            ),
        )
        fig_wo.update_layout(height=440, margin=dict(l=10, r=25, t=10, b=10), plot_bgcolor="white", paper_bgcolor="white")
        st.plotly_chart(fig_wo, use_container_width=True)
    else:
        st.warning("No data available for selected filters.")

with c2:
    st.subheader("Penalty Applied vs Waived/Not Applied")
    penalty_df = pd.DataFrame({
        "Status": ["Penalty Applied", "Expected Penalties Waived", "Other No Penalty"],
        "Count": [penalty, waived, other_no_penalty],
    })
    fig_pen = px.pie(
        penalty_df,
        values="Count",
        names="Status",
        hole=0.56,
        color="Status",
        color_discrete_map={"Penalty Applied": PINK, "Expected Penalties Waived": ORANGE, "Other No Penalty": GREEN},
    )
    fig_pen.update_traces(
        textposition="outside",
        textinfo="label+value+percent",
        pull=[0.02, 0.035, 0.02],
        marker=dict(line=dict(color="white", width=2)),
        hovertemplate="%{label}<br>%{value} deviations<br>%{percent}<extra></extra>",
    )
    fig_pen.update_layout(
        height=440,
        margin=dict(l=20, r=20, t=10, b=10),
        legend=dict(orientation="h", y=-0.08, x=0.02),
        annotations=[dict(text=f"{total}<br>Total", x=0.5, y=0.5, showarrow=False, font=dict(size=18, color=TEXT))],
    )
    st.plotly_chart(fig_pen, use_container_width=True)

c3, c4 = st.columns([1, 1])
with c3:
    st.subheader("Civil / Fiber / Safety Classification")
    cls = filtered["TeamClassification"].value_counts().reset_index()
    cls.columns = ["Classification", "Count"]
    fig_cls = px.bar(
        cls,
        x="Classification",
        y="Count",
        color="Classification",
        text="Count",
        color_discrete_map={"Civil": PURPLE, "Fiber": TEAL, "Safety": ORANGE, "Other": MUTED},
    )
    fig_cls.update_traces(textposition="outside")
    fig_cls.update_layout(height=360, showlegend=False, margin=dict(l=10, r=10, t=10, b=10), plot_bgcolor="white")
    st.plotly_chart(fig_cls, use_container_width=True)

with c4:
    st.subheader("Foreman / Inspector Accountability")
    people = filtered.groupby(["Civil Foreman", "Inspector"], dropna=False).size().reset_index(name="Deviation Count").sort_values("Deviation Count", ascending=False)
    fig_people = px.treemap(
        people,
        path=["Civil Foreman", "Inspector"],
        values="Deviation Count",
        color="Deviation Count",
        color_continuous_scale=["#F3E8FF", PURPLE],
    )
    fig_people.update_layout(height=360, margin=dict(l=10, r=10, t=10, b=10))
    st.plotly_chart(fig_people, use_container_width=True)

st.markdown('<div class="section-title">Board Work Order Executive Summary</div>', unsafe_allow_html=True)
st.caption("Main board table included in PDF export. Includes Civil Foreman, Inspector, Expected Penalties Waived and full WO numbers.")
show_cols = [
    "Work Order", "District", "Civil Foreman", "Inspector", "Total Deviations",
    "Penalty Applied", "No Penalty Applied", "Expected Penalties Waived", "Service Affecting",
    "Civil", "Fiber", "Safety", "Other", "% of Total", "Top Deviation Nature",
]
st.dataframe(wo_summary[show_cols] if not wo_summary.empty else wo_summary, use_container_width=True, hide_index=True, height=360)

st.markdown('<div class="section-title">Detailed Deviation Nature by Work Order</div>', unsafe_allow_html=True)
st.dataframe(nature, use_container_width=True, hide_index=True, height=420)

st.markdown('<div class="section-title">Export Board-ready PDF</div>', unsafe_allow_html=True)
try:
    pdf_bytes = make_pdf(filtered, wo_summary, nature)
    st.download_button(
        "Export Executive PDF Report",
        data=pdf_bytes,
        file_name="STC_Quality_Executive_Dashboard_Board_Report.pdf",
        mime="application/pdf",
        use_container_width=True,
    )
except Exception as exc:
    st.error(f"PDF export failed: {exc}")

st.caption("Source file: Deviation.xlsx | Expected Penalties = penalties that should have been applied but were waived/cancelled.")
