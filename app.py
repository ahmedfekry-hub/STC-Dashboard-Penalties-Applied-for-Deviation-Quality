import io
from datetime import datetime
from textwrap import shorten

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
    PageBreak,
)

# =========================
# Page & Theme
# =========================
st.set_page_config(
    page_title="STC Quality Executive Dashboard",
    layout="wide",
    initial_sidebar_state="expanded",
)

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

st.markdown(
    f"""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700;800&display=swap');
        html, body, [class*="css"] {{font-family: 'Inter', sans-serif;}}
        .main .block-container {{padding-top: 1.0rem; padding-bottom: 2rem; max-width: 1500px;}}
        .hero {{
            background: linear-gradient(120deg, {PURPLE}, {PURPLE_2});
            color: white;
            border-radius: 0 0 22px 22px;
            padding: 24px 28px;
            margin-bottom: 18px;
            box-shadow: 0 18px 45px rgba(90,10,162,0.18);
            animation: fadeInDown .65s ease-in-out;
        }}
        .hero h1 {{margin: 0; font-size: 32px; font-weight: 800;}}
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
        .metric-label {{font-size: 13px; color: {TEXT}; font-weight: 600;}}
        .metric-value {{font-size: 32px; line-height: 1; font-weight: 800; margin-top: 12px;}}
        .metric-note {{font-size: 12px; color: {MUTED}; margin-top: 12px;}}
        .section-title {{font-size: 22px; font-weight: 800; color: {TEXT}; margin-top: 16px; margin-bottom: 8px;}}
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
# Helpers
# =========================
def _clean_bool_series(s: pd.Series) -> pd.Series:
    return s.astype(str).str.strip().str.upper().isin(["Y", "YES", "TRUE", "1", "APPLIED"])


def _normalize_text(s: pd.Series) -> pd.Series:
    return s.fillna("").astype(str).str.strip()


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
    df["WorkOrderNum"] = _normalize_text(df.get("WorkOrderNum", pd.Series(index=df.index)))
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
    df["ServiceAffectingFlag"] = df["ServiceAffecting"].astype(str).str.strip().str.upper().isin(["Y", "YES", "TRUE", "1"])
    # Expected Penalties means: penalty should have applied but was waived/cancelled.
    df["ExpectedPenaltyWaived"] = df["Expected Penalties"].astype(str).str.strip().str.upper().isin(["Y", "YES", "TRUE", "1"])
    df["NoPenaltyApplied"] = ~df["PenaltyApplied"]

    def classify(row) -> str:
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
    return df


def pct(n, d):
    return 0 if d == 0 else round((n / d) * 100, 1)


def comma_join_unique(values, limit=3):
    vals = [v for v in pd.unique(pd.Series(values).dropna()) if str(v).strip() not in ["", "—", "nan"]]
    if not vals:
        return "—"
    text = " / ".join(map(str, vals[:limit]))
    if len(vals) > limit:
        text += f" +{len(vals)-limit}"
    return text


def top_deviation_nature(group: pd.DataFrame, limit=3) -> str:
    vc = group["DeviationName"].value_counts().head(limit)
    if vc.empty:
        return "—"
    return " | ".join([f"{shorten(str(k), width=38, placeholder='…')} ({int(v)})" for k, v in vc.items()])


def build_wo_summary(df: pd.DataFrame) -> pd.DataFrame:
    grouped = []
    total_rows = len(df)
    for wo, g in df.groupby("WorkOrderNum", dropna=False):
        row = {
            "Work Order": wo,
            "District": comma_join_unique(g["District"], limit=2),
            "Civil Foreman": comma_join_unique(g["Civil Foreman"], limit=3),
            "Inspector": comma_join_unique(g["Inspector"], limit=3),
            "Total Deviations": int(len(g)),
            "Penalty Applied": int(g["PenaltyApplied"].sum()),
            "No Penalty Applied": int(g["NoPenaltyApplied"].sum()),
            "Expected Penalties Waived": int(g["ExpectedPenaltyWaived"].sum()),
            "Service Affecting": int(g["ServiceAffectingFlag"].sum()),
            "Civil": int((g["TeamClassification"] == "Civil").sum()),
            "Fiber": int((g["TeamClassification"] == "Fiber").sum()),
            "Safety": int((g["TeamClassification"] == "Safety").sum()),
            "Other": int((g["TeamClassification"] == "Other").sum()),
            "% of Total": pct(len(g), total_rows),
            "Top Deviation Nature": top_deviation_nature(g, limit=3),
        }
        grouped.append(row)
    out = pd.DataFrame(grouped).sort_values("Total Deviations", ascending=False).reset_index(drop=True)
    return out


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
# PDF Builder
# =========================
def make_pdf(df: pd.DataFrame, wo_summary: pd.DataFrame, nature: pd.DataFrame) -> bytes:
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=landscape(A4),
        rightMargin=0.8 * cm,
        leftMargin=0.8 * cm,
        topMargin=0.7 * cm,
        bottomMargin=0.7 * cm,
    )
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="TitlePurple", parent=styles["Title"], textColor=colors.HexColor(PURPLE), fontSize=22, leading=25, alignment=TA_CENTER, spaceAfter=6))
    styles.add(ParagraphStyle(name="Sub", parent=styles["Normal"], textColor=colors.HexColor(MUTED), fontSize=9.5, alignment=TA_CENTER, spaceAfter=8))
    styles.add(ParagraphStyle(name="Section", parent=styles["Heading2"], textColor=colors.HexColor(TEXT), fontSize=14, leading=17, spaceBefore=8, spaceAfter=5))
    styles.add(ParagraphStyle(name="Cell", parent=styles["Normal"], fontSize=6.5, leading=7.6, alignment=TA_LEFT))
    styles.add(ParagraphStyle(name="HeaderCell", parent=styles["Normal"], fontSize=6.6, leading=7.8, textColor=colors.white, alignment=TA_CENTER))

    story = []
    story.append(Paragraph("STC Quality Executive Dashboard", styles["TitlePurple"]))
    story.append(Paragraph("Deviation WO drilldown • penalty applied vs waived • Civil / Fiber / Safety classification • Foreman & Inspector accountability", styles["Sub"]))

    total = len(df)
    penalty = int(df["PenaltyApplied"].sum())
    no_penalty = int(df["NoPenaltyApplied"].sum())
    waived = int(df["ExpectedPenaltyWaived"].sum())
    service = int(df["ServiceAffectingFlag"].sum())
    unique_wo = df["WorkOrderNum"].nunique()

    metrics = [
        ["Total Deviations", f"{total:,}"],
        ["Unique WOs", f"{unique_wo:,}"],
        ["Penalty Applied", f"{penalty:,} ({pct(penalty,total)}%)"],
        ["No Penalty Applied", f"{no_penalty:,} ({pct(no_penalty,total)}%)"],
        ["Expected Penalties Waived", f"{waived:,} ({pct(waived,total)}%)"],
        ["Service Affecting", f"{service:,}"],
    ]
    t = Table(metrics, colWidths=[4.6 * cm, 3.1 * cm], hAlign="CENTER")
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("F2ECFA")),
        ("TEXTCOLOR", (0, 0), (-1, -1), colors.HexColor(TEXT)),
        ("FONTNAME", (0, 0), (-1, -1), "Helvetica-Bold"),
        ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("D9CCE9")),
        ("ALIGN", (1, 0), (1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ROWBACKGROUNDS", (0, 0), (-1, -1), [colors.HexColor("F7F2FC"), colors.HexColor("FFFFFF")]),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))
    story.append(t)
    story.append(Spacer(1, 0.25 * cm))

    # Board table: compact but complete
    story.append(Paragraph("Board Work Order Executive Summary", styles["Section"]))
    board_cols = [
        "Work Order", "District", "Civil Foreman", "Inspector", "Total Deviations",
        "Penalty Applied", "No Penalty Applied", "Expected Penalties Waived", "Service Affecting",
        "Civil", "Fiber", "Safety", "Other", "% of Total", "Top Deviation Nature",
    ]
    board_df = wo_summary[board_cols].copy()
    data = [[Paragraph(str(c), styles["HeaderCell"]) for c in board_cols]]
    for _, r in board_df.iterrows():
        data.append([Paragraph(str(r[c]), styles["Cell"]) for c in board_cols])
    col_widths = [2.05*cm, 1.45*cm, 2.65*cm, 2.25*cm, 1.35*cm, 1.35*cm, 1.45*cm, 1.65*cm, 1.25*cm, 0.85*cm, 0.85*cm, 0.95*cm, 0.85*cm, 1.05*cm, 5.2*cm]
    table = Table(data, colWidths=col_widths, repeatRows=1, hAlign="CENTER")
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(PURPLE)),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("DDD3EA")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("FAF7FE")]),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (4, 1), (13, -1), "CENTER"),
        ("TOPPADDING", (0, 0), (-1, -1), 3.5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3.5),
    ]))
    story.append(table)

    # Deviation nature table on next page
    story.append(PageBreak())
    story.append(Paragraph("Detailed Deviation Nature by Work Order", styles["TitlePurple"]))
    story.append(Paragraph("This table shows the exact nature of repeated deviations by WO with responsible foreman and inspector names.", styles["Sub"]))

    detail_cols = [
        "Work Order", "District", "Civil Foreman", "Inspector", "Classification", "Category", "SubCategory", "DeviationName",
        "Total", "Penalty_Applied", "No_Penalty", "Expected_Penalties_Waived", "Service_Affecting",
    ]
    detail_df = nature[detail_cols].copy()
    detail_data = [[Paragraph(str(c).replace("_", " "), styles["HeaderCell"]) for c in detail_cols]]
    for _, r in detail_df.iterrows():
        detail_data.append([Paragraph(str(r[c]), styles["Cell"]) for c in detail_cols])

    detail_widths = [2.0*cm, 1.3*cm, 2.35*cm, 2.1*cm, 1.45*cm, 1.6*cm, 1.85*cm, 6.3*cm, 0.75*cm, 0.95*cm, 0.95*cm, 1.15*cm, 0.95*cm]
    detail_table = Table(detail_data, colWidths=detail_widths, repeatRows=1, hAlign="CENTER")
    detail_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(PURPLE)),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("DDD3EA")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("FAF7FE")]),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (8, 1), (-1, -1), "CENTER"),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
    ]))
    story.append(detail_table)

    def _footer(canvas, doc):
        canvas.saveState()
        canvas.setFont("Helvetica", 7)
        canvas.setFillColor(colors.HexColor(MUTED))
        canvas.drawString(0.8 * cm, 0.35 * cm, f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')} | Source: Deviation.xlsx")
        canvas.drawRightString(28.8 * cm, 0.35 * cm, f"Page {doc.page}")
        canvas.restoreState()

    doc.build(story, onFirstPage=_footer, onLaterPages=_footer)
    buffer.seek(0)
    return buffer.read()


# =========================
# Sidebar / Data
# =========================
with st.sidebar:
    st.markdown("### 📂 Source File")
    uploaded = st.file_uploader("Upload updated Deviation.xlsx", type=["xlsx"], help="Optional. If not uploaded, app uses the included Deviation.xlsx file.")
    st.markdown("### 🔎 Filters")

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
    penalty_view = st.radio("Penalty View", ["All", "Penalty Applied", "No Penalty", "Expected Penalties Waived"], horizontal=False)

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
elif penalty_view == "No Penalty":
    filtered = filtered[filtered["NoPenaltyApplied"]]
elif penalty_view == "Expected Penalties Waived":
    filtered = filtered[filtered["ExpectedPenaltyWaived"]]

wo_summary = build_wo_summary(filtered)
nature = build_deviation_nature(filtered)

# =========================
# Dashboard UI
# =========================
st.markdown(
    """
    <div class="hero">
        <h1>STC Quality Executive Dashboard</h1>
        <p>Deviation analysis • Work Order performance • Penalty applied vs waived • Civil / Fiber / Safety accountability • Foreman & Inspector control view</p>
    </div>
    """,
    unsafe_allow_html=True,
)

total = len(filtered)
penalty = int(filtered["PenaltyApplied"].sum())
no_penalty = int(filtered["NoPenaltyApplied"].sum())
waived = int(filtered["ExpectedPenaltyWaived"].sum())
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
            animation_frame=None,
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
        "Count": [penalty, waived, max(no_penalty - waived, 0)],
    })
    fig_pen = px.pie(
        penalty_df,
        values="Count",
        names="Status",
        hole=0.58,
        color="Status",
        color_discrete_map={"Penalty Applied": PINK, "Expected Penalties Waived": ORANGE, "Other No Penalty": GREEN},
    )
    fig_pen.update_traces(textinfo="label+value+percent", hovertemplate="%{label}<br>%{value} deviations<br>%{percent}<extra></extra>")
    fig_pen.update_layout(height=440, margin=dict(l=10, r=10, t=10, b=10), legend=dict(orientation="h", y=-0.05))
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
st.caption("This is the main table that will appear in the PDF export. It includes Civil Foreman, Inspector, Expected Penalties Waived and full WO numbers.")
show_cols = [
    "Work Order", "District", "Civil Foreman", "Inspector", "Total Deviations",
    "Penalty Applied", "No Penalty Applied", "Expected Penalties Waived", "Service Affecting",
    "Civil", "Fiber", "Safety", "Other", "% of Total", "Top Deviation Nature",
]
st.dataframe(wo_summary[show_cols], use_container_width=True, hide_index=True, height=360)

st.markdown('<div class="section-title">Detailed Deviation Nature by Work Order</div>', unsafe_allow_html=True)
st.dataframe(nature, use_container_width=True, hide_index=True, height=420)

st.markdown('<div class="section-title">Export Board-ready PDF</div>', unsafe_allow_html=True)
pdf_bytes = make_pdf(filtered, wo_summary, nature)
st.download_button(
    "📄 Export Executive PDF Report",
    data=pdf_bytes,
    file_name="STC_Quality_Executive_Dashboard_Board_Report.pdf",
    mime="application/pdf",
    use_container_width=True,
)

st.caption("Source file: Deviation.xlsx | Expected Penalties = penalties that should have been applied but were waived/cancelled.")
