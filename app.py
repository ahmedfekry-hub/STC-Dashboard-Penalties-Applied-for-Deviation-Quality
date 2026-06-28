
import os
import re
import io
import textwrap
from datetime import datetime

import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go

# Optional PDF dependency
try:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import cm
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
    REPORTLAB_READY = True
except Exception:
    REPORTLAB_READY = False


# =====================
# Page Configuration
# =====================
st.set_page_config(
    page_title="STC Quality Executive Dashboard",
    layout="wide",
    initial_sidebar_state="expanded",
)

PURPLE = "#5A0AA2"
PURPLE_2 = "#7B1FE6"
TEAL = "#26C6DA"
ORANGE = "#FF7A45"
PINK = "#EF476F"
GREEN = "#10B981"
GRAY = "#6B7280"
LIGHT_BG = "#F8F6FB"
CARD_BG = "#FFFFFF"
BORDER = "#E7E0F0"

st.markdown(
    f"""
    <style>
        .block-container {{
            padding-top: 1.0rem;
            padding-bottom: 1.0rem;
        }}
        .main {{
            background-color: {LIGHT_BG};
        }}
        .title-box {{
            background: linear-gradient(90deg, {PURPLE}, {PURPLE_2});
            color: white;
            padding: 20px 26px;
            border-radius: 18px;
            margin-bottom: 16px;
            box-shadow: 0 8px 24px rgba(90,10,162,0.18);
        }}
        .title-box h1 {{
            font-size: 34px;
            margin: 0;
            font-weight: 800;
        }}
        .title-box p {{
            font-size: 15px;
            margin: 8px 0 0 0;
            opacity: 0.92;
        }}
        .metric-card {{
            background: {CARD_BG};
            border: 1px solid {BORDER};
            border-radius: 16px;
            padding: 16px 18px;
            box-shadow: 0 3px 14px rgba(40, 20, 70, 0.05);
            min-height: 112px;
        }}
        .metric-label {{
            color: #6B6178;
            font-size: 14px;
            font-weight: 600;
            margin-bottom: 8px;
        }}
        .metric-value {{
            color: #2F2144;
            font-size: 34px;
            line-height: 1.1;
            font-weight: 800;
        }}
        .metric-sub {{
            color: #8A8198;
            font-size: 13px;
            margin-top: 8px;
        }}
        .section-title {{
            font-size: 24px;
            color: #2F2144;
            font-weight: 800;
            margin: 12px 0 10px 0;
        }}
        .small-note {{
            color: #7A7288;
            font-size: 13px;
        }}
        div[data-testid="stDataFrame"] {{
            border: 1px solid {BORDER};
            border-radius: 12px;
            overflow: hidden;
        }}
    </style>
    """,
    unsafe_allow_html=True,
)


# =====================
# Helpers
# =====================
def clean_text(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def normalize_yes_no(value):
    v = clean_text(value).upper()
    if v in ["Y", "YES", "TRUE", "1", "APPLIED"]:
        return "Y"
    if v in ["N", "NO", "FALSE", "0", "NOT APPLIED", "NONE"]:
        return "N"
    return v if v else "N"


def shorten_deviation(text, max_len=70):
    text = clean_text(text)
    text = re.sub(r"\s+", " ", text)
    if len(text) <= max_len:
        return text
    return text[: max_len - 3] + "..."


def classify_scope(row):
    # Priority 1: use Designation if it exists.
    designation = clean_text(row.get("Designation", "")).upper()
    if "SAF" in designation:
        return "Safety"
    if "FIB" in designation:
        return "Fiber"
    if "CIV" in designation:
        return "Civil"

    # Priority 2: classify from category/subcategory/deviation keywords.
    text = " ".join([
        clean_text(row.get("Category", "")),
        clean_text(row.get("SubCategory", "")),
        clean_text(row.get("DeviationName", "")),
    ]).upper()

    safety_kw = ["SAFETY", "PPE", "UNIFORM", "SIGNBOARD", "PEDESTRIAN", "WORKER", "ID BADGE", "REGISTER"]
    fiber_kw = ["FIBER", "CABLE", "SPLIC", "LABEL", "ODF", "PATCH", "JOINT", "CLOSURE"]
    civil_kw = ["CIVIL", "TRENCH", "BORE", "DUCT", "CONDUIT", "MANHOLE", "HANDHOLE", "ASPHALT", "BACKFILL", "PERMIT", "DEBRIS", "ROUTE", "WIDTH", "DEPTH", "SPACER"]

    if any(k in text for k in safety_kw):
        return "Safety"
    if any(k in text for k in fiber_kw):
        return "Fiber"
    if any(k in text for k in civil_kw):
        return "Civil"
    return "Other"


@st.cache_data(show_spinner=False)
def load_excel(uploaded_file=None):
    if uploaded_file is not None:
        data = pd.read_excel(uploaded_file)
    else:
        default_path = os.path.join(os.path.dirname(__file__), "Deviation.xlsx")
        data = pd.read_excel(default_path)

    data.columns = [str(c).strip() for c in data.columns]

    required = ["District", "WorkOrderNum", "DeviationName", "IsPenalty"]
    missing = [c for c in required if c not in data.columns]
    if missing:
        st.error(f"Missing required columns in Deviation.xlsx: {', '.join(missing)}")
        st.stop()

    data["District"] = data["District"].apply(clean_text).str.upper()
    data["WorkOrderNum"] = data["WorkOrderNum"].apply(clean_text)
    data["DeviationName"] = data["DeviationName"].apply(clean_text)
    data["DeviationShort"] = data["DeviationName"].apply(shorten_deviation)
    data["PenaltyFlag"] = data["IsPenalty"].apply(normalize_yes_no)
    data["PenaltyStatus"] = np.where(data["PenaltyFlag"].eq("Y"), "Penalty Applied", "No Penalty")
    if "ServiceAffecting" in data.columns:
        data["ServiceAffectingClean"] = data["ServiceAffecting"].apply(normalize_yes_no)
    else:
        data["ServiceAffectingClean"] = "N"

    if "DateOfDeviation" in data.columns:
        data["DateOfDeviation"] = pd.to_datetime(data["DateOfDeviation"], errors="coerce")

    data["Scope"] = data.apply(classify_scope, axis=1)

    return data


def filter_data(data):
    with st.sidebar:
        st.markdown("### Filters")
        district = st.multiselect("District", sorted(data["District"].dropna().unique()))
        wo = st.multiselect("Work Order", sorted(data["WorkOrderNum"].dropna().unique()))
        scope = st.multiselect("Classification", ["Civil", "Fiber", "Safety", "Other"])
        penalty = st.multiselect("Penalty Status", ["Penalty Applied", "No Penalty"])

        if "DeviationStatus" in data.columns:
            status = st.multiselect("Deviation Status", sorted(data["DeviationStatus"].dropna().astype(str).unique()))
        else:
            status = []

        search = st.text_input("Search in Deviation Name / WO", placeholder="Type keyword or WO...")

    f = data.copy()

    if district:
        f = f[f["District"].isin(district)]
    if wo:
        f = f[f["WorkOrderNum"].isin(wo)]
    if scope:
        f = f[f["Scope"].isin(scope)]
    if penalty:
        f = f[f["PenaltyStatus"].isin(penalty)]
    if status and "DeviationStatus" in f.columns:
        f = f[f["DeviationStatus"].astype(str).isin(status)]
    if search:
        s = search.strip().lower()
        f = f[
            f["DeviationName"].str.lower().str.contains(s, na=False)
            | f["WorkOrderNum"].str.lower().str.contains(s, na=False)
        ]

    return f


def pct(part, total):
    if total == 0:
        return 0
    return round((part / total) * 100, 1)


def metric_card(label, value, sub="", accent=PURPLE):
    st.markdown(
        f"""
        <div class="metric-card">
            <div class="metric-label">{label}</div>
            <div class="metric-value" style="color:{accent};">{value}</div>
            <div class="metric-sub">{sub}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def make_pdf_report(data, wo_summary, top_deviation, scope_summary):
    if not REPORTLAB_READY:
        return None

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
    title_style = ParagraphStyle(
        "TitlePurple",
        parent=styles["Title"],
        fontSize=20,
        textColor=colors.HexColor(PURPLE),
        spaceAfter=10,
    )
    h_style = ParagraphStyle(
        "Heading",
        parent=styles["Heading2"],
        fontSize=14,
        textColor=colors.HexColor("#2F2144"),
        spaceBefore=8,
        spaceAfter=6,
    )
    normal = ParagraphStyle(
        "NormalSmall",
        parent=styles["Normal"],
        fontSize=8,
        leading=10,
    )

    story = []
    story.append(Paragraph("STC Quality Executive Dashboard", title_style))
    story.append(Paragraph("Deviation WO drilldown • penalty status • Civil / Fiber / Safety classification", styles["Normal"]))
    story.append(Spacer(1, 8))

    total = len(data)
    penalty_y = int((data["PenaltyFlag"] == "Y").sum())
    penalty_n = int((data["PenaltyFlag"] != "Y").sum())
    service_affecting = int((data["ServiceAffectingClean"] == "Y").sum())
    unique_wo = data["WorkOrderNum"].nunique()

    cards = [
        ["Total Deviations", f"{total:,}"],
        ["Unique WOs", f"{unique_wo:,}"],
        ["Penalty Applied", f"{penalty_y:,} ({pct(penalty_y, total)}%)"],
        ["No Penalty", f"{penalty_n:,} ({pct(penalty_n, total)}%)"],
        ["Service Affecting", f"{service_affecting:,}"],
    ]
    card_table = Table(cards, colWidths=[3.2 * cm, 3.0 * cm])
    card_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#F4EEFB")),
        ("TEXTCOLOR", (0, 0), (-1, -1), colors.HexColor("#2F2144")),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#DFD2EC")),
        ("FONTNAME", (0, 0), (-1, -1), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
    ]))
    story.append(card_table)
    story.append(Spacer(1, 10))

    story.append(Paragraph("Top Work Orders", h_style))
    wo_pdf = wo_summary.head(20).copy()
    table_data = [["WorkOrder", "District", "Total", "Penalty", "No Penalty", "Civil", "Fiber", "Safety", "% of Total"]]
    for _, r in wo_pdf.iterrows():
        table_data.append([
            str(r["WorkOrderNum"]), str(r["District"]), str(r["Total Deviations"]),
            str(r["Penalty Applied"]), str(r["No Penalty"]),
            str(r.get("Civil", 0)), str(r.get("Fiber", 0)), str(r.get("Safety", 0)),
            f'{r["% of Total"]:.1f}%'
        ])
    t = Table(table_data, repeatRows=1)
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(PURPLE)),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#E6E0ED")),
        ("FONTSIZE", (0, 0), (-1, -1), 7),
        ("ALIGN", (2, 1), (-1, -1), "CENTER"),
    ]))
    story.append(t)
    story.append(PageBreak())

    story.append(Paragraph("Classification Summary", h_style))
    scope_table = [["Scope", "Count", "%"]]
    for _, r in scope_summary.iterrows():
        scope_table.append([str(r["Scope"]), str(r["Count"]), f'{r["%"]:.1f}%'])
    t2 = Table(scope_table, colWidths=[5 * cm, 3 * cm, 3 * cm])
    t2.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(PURPLE)),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#E6E0ED")),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
    ]))
    story.append(t2)
    story.append(Spacer(1, 10))

    story.append(Paragraph("Top Deviation Types", h_style))
    dev_table = [["Deviation Type", "Count", "% of Total", "Penalty Applied"]]
    for _, r in top_deviation.head(25).iterrows():
        dev_table.append([
            Paragraph(str(r["DeviationShort"]), normal),
            str(r["Count"]),
            f'{r["% of Total"]:.1f}%',
            str(r["Penalty Applied"])
        ])
    t3 = Table(dev_table, colWidths=[14 * cm, 3 * cm, 3 * cm, 3 * cm], repeatRows=1)
    t3.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(PURPLE)),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#E6E0ED")),
        ("FONTSIZE", (0, 0), (-1, -1), 7),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
    ]))
    story.append(t3)

    doc.build(story)
    buffer.seek(0)
    return buffer


# =====================
# Load Data
# =====================
with st.sidebar:
    st.markdown("## Data Source")
    uploaded_file = st.file_uploader("Upload Deviation.xlsx", type=["xlsx"])
    st.caption("The dashboard runs on one file only: Deviation.xlsx")

data = load_excel(uploaded_file)
filtered = filter_data(data)

# =====================
# Header
# =====================
st.markdown(
    """
    <div class="title-box">
        <h1>STC Quality Executive Dashboard</h1>
        <p>Deviation analysis • Work Order performance • Penalty applied vs not applied • Civil / Fiber / Safety classification</p>
    </div>
    """,
    unsafe_allow_html=True,
)

# =====================
# KPIs
# =====================
total = len(filtered)
unique_wo = filtered["WorkOrderNum"].nunique()
penalty_applied = int((filtered["PenaltyFlag"] == "Y").sum())
no_penalty = int((filtered["PenaltyFlag"] != "Y").sum())
service_affecting = int((filtered["ServiceAffectingClean"] == "Y").sum())

c1, c2, c3, c4, c5 = st.columns(5)
with c1:
    metric_card("Total Deviations", f"{total:,}", "All records after filters", PURPLE)
with c2:
    metric_card("Unique WOs", f"{unique_wo:,}", "Impacted Work Orders", TEAL)
with c3:
    metric_card("Penalty Applied", f"{penalty_applied:,}", f"{pct(penalty_applied, total)}% of deviations", PINK)
with c4:
    metric_card("No Penalty", f"{no_penalty:,}", f"{pct(no_penalty, total)}% of deviations", GREEN)
with c5:
    metric_card("Service Affecting", f"{service_affecting:,}", f"{pct(service_affecting, total)}% of deviations", ORANGE)

st.markdown('<div class="section-title">Executive Charts</div>', unsafe_allow_html=True)

# =====================
# Summary Frames
# =====================
scope_summary = (
    filtered.groupby("Scope")
    .size()
    .reset_index(name="Count")
    .sort_values("Count", ascending=False)
)
scope_summary["%"] = scope_summary["Count"].apply(lambda x: pct(x, total))

penalty_summary = (
    filtered.groupby("PenaltyStatus")
    .size()
    .reset_index(name="Count")
    .sort_values("Count", ascending=False)
)

top_deviation = (
    filtered.groupby("DeviationShort")
    .agg(
        Count=("DeviationShort", "size"),
        **{"Penalty Applied": ("PenaltyFlag", lambda x: int((x == "Y").sum()))},
        **{"No Penalty": ("PenaltyFlag", lambda x: int((x != "Y").sum()))},
    )
    .reset_index()
    .sort_values("Count", ascending=False)
)
top_deviation["% of Total"] = top_deviation["Count"].apply(lambda x: pct(x, total))

scope_pivot = pd.crosstab(filtered["WorkOrderNum"], filtered["Scope"])
penalty_pivot = pd.crosstab(filtered["WorkOrderNum"], filtered["PenaltyStatus"])
wo_base = (
    filtered.groupby(["WorkOrderNum", "District"])
    .size()
    .reset_index(name="Total Deviations")
    .sort_values("Total Deviations", ascending=False)
)

wo_summary = wo_base.merge(scope_pivot, on="WorkOrderNum", how="left").merge(penalty_pivot, on="WorkOrderNum", how="left")
for col in ["Civil", "Fiber", "Safety", "Other", "Penalty Applied", "No Penalty"]:
    if col not in wo_summary.columns:
        wo_summary[col] = 0
wo_summary["% of Total"] = wo_summary["Total Deviations"].apply(lambda x: pct(x, total))
wo_summary = wo_summary.sort_values("Total Deviations", ascending=False)

# =====================
# Charts Row 1
# =====================
col_a, col_b = st.columns([1.15, 0.85])

with col_a:
    st.markdown("#### Top Work Orders by Deviation Count")
    if wo_summary.empty:
        st.info("No records to display.")
    else:
        top_wo_chart = wo_summary.head(15).sort_values("Total Deviations")
        fig = px.bar(
            top_wo_chart,
            x="Total Deviations",
            y="WorkOrderNum",
            color="District",
            orientation="h",
            text="Total Deviations",
            color_discrete_sequence=[PURPLE, TEAL, ORANGE, PINK, GREEN],
        )
        fig.update_traces(textposition="outside", cliponaxis=False)
        fig.update_layout(
            height=460,
            margin=dict(l=5, r=20, t=20, b=20),
            xaxis_title="Deviation Count",
            yaxis_title="Work Order",
            legend_title="District",
        )
        st.plotly_chart(fig, use_container_width=True)

with col_b:
    st.markdown("#### Penalty Applied vs No Penalty")
    if penalty_summary.empty:
        st.info("No records to display.")
    else:
        fig = px.pie(
            penalty_summary,
            names="PenaltyStatus",
            values="Count",
            hole=0.55,
            color="PenaltyStatus",
            color_discrete_map={
                "Penalty Applied": PINK,
                "No Penalty": GREEN,
            },
        )
        fig.update_traces(textinfo="label+percent+value")
        fig.update_layout(height=460, margin=dict(l=10, r=10, t=20, b=20))
        st.plotly_chart(fig, use_container_width=True)

# =====================
# Charts Row 2
# =====================
col_c, col_d = st.columns([0.85, 1.15])

with col_c:
    st.markdown("#### Civil / Fiber / Safety Classification")
    if scope_summary.empty:
        st.info("No records to display.")
    else:
        fig = px.bar(
            scope_summary.sort_values("Count"),
            x="Count",
            y="Scope",
            orientation="h",
            text="Count",
            color="Scope",
            color_discrete_map={
                "Civil": PURPLE,
                "Fiber": TEAL,
                "Safety": ORANGE,
                "Other": GRAY,
            },
        )
        fig.update_traces(textposition="outside", cliponaxis=False)
        fig.update_layout(height=430, margin=dict(l=10, r=20, t=20, b=20), showlegend=False)
        st.plotly_chart(fig, use_container_width=True)

with col_d:
    st.markdown("#### Most Common Deviation Types")
    if top_deviation.empty:
        st.info("No records to display.")
    else:
        top_dev_chart = top_deviation.head(12).sort_values("Count")
        fig = px.bar(
            top_dev_chart,
            x="Count",
            y="DeviationShort",
            orientation="h",
            text="Count",
            color_discrete_sequence=[PURPLE],
        )
        fig.update_traces(textposition="outside", cliponaxis=False)
        fig.update_layout(
            height=430,
            margin=dict(l=10, r=20, t=20, b=20),
            xaxis_title="Deviation Count",
            yaxis_title="Deviation Type",
        )
        st.plotly_chart(fig, use_container_width=True)

# =====================
# Heatmap
# =====================
st.markdown('<div class="section-title">Work Order Classification Heatmap</div>', unsafe_allow_html=True)
heat = pd.crosstab(filtered["WorkOrderNum"], filtered["Scope"])
if not heat.empty:
    heat = heat.loc[heat.sum(axis=1).sort_values(ascending=False).head(20).index]
    fig = px.imshow(
        heat,
        text_auto=True,
        aspect="auto",
        color_continuous_scale=[[0, "#F3E8FF"], [1, PURPLE]],
    )
    fig.update_layout(height=520, margin=dict(l=10, r=10, t=10, b=10), coloraxis_showscale=False)
    st.plotly_chart(fig, use_container_width=True)
else:
    st.info("No records to display.")

# =====================
# Tables
# =====================
st.markdown('<div class="section-title">WO Executive Summary</div>', unsafe_allow_html=True)
show_cols = [
    "WorkOrderNum", "District", "Total Deviations", "% of Total",
    "Penalty Applied", "No Penalty", "Civil", "Fiber", "Safety", "Other"
]
st.dataframe(wo_summary[show_cols], use_container_width=True, hide_index=True)

st.markdown('<div class="section-title">Deviation Nature by WO</div>', unsafe_allow_html=True)
wo_nature = (
    filtered.groupby(["WorkOrderNum", "District", "Scope", "DeviationShort", "PenaltyStatus"])
    .size()
    .reset_index(name="Count")
    .sort_values(["WorkOrderNum", "Count"], ascending=[True, False])
)
st.dataframe(wo_nature, use_container_width=True, hide_index=True)

# =====================
# Export
# =====================
st.markdown('<div class="section-title">Export</div>', unsafe_allow_html=True)
export_cols = [
    "District", "WorkOrderNum", "DeviationName", "Scope", "PenaltyStatus",
    "IsPenalty", "ServiceAffectingClean"
]
for optional in ["Category", "SubCategory", "Designation", "DeviationStatus", "CorrectionTime", "DateOfDeviation"]:
    if optional in filtered.columns and optional not in export_cols:
        export_cols.append(optional)

excel_buffer = io.BytesIO()
with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
    wo_summary[show_cols].to_excel(writer, index=False, sheet_name="WO Summary")
    top_deviation.to_excel(writer, index=False, sheet_name="Top Deviations")
    scope_summary.to_excel(writer, index=False, sheet_name="Classification")
    filtered[export_cols].to_excel(writer, index=False, sheet_name="Clean Data")
excel_buffer.seek(0)

st.download_button(
    "Download Excel Summary",
    data=excel_buffer,
    file_name=f"STC_Quality_Deviation_Summary_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

pdf_buffer = make_pdf_report(filtered, wo_summary, top_deviation, scope_summary)
if pdf_buffer is not None:
    st.download_button(
        "Download Executive PDF Report",
        data=pdf_buffer,
        file_name=f"STC_Quality_Executive_Report_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
        mime="application/pdf",
    )
else:
    st.caption("PDF export requires reportlab. Install dependencies from requirements.txt.")

st.caption("Source: Deviation.xlsx only. No additional CSV/database files are required.")
