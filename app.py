import os
import re
from io import BytesIO
from datetime import datetime

import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go

# PDF / charts
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch
from reportlab.lib import colors
from reportlab.lib.colors import HexColor
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
    Image,
    PageBreak,
)

try:
    import arabic_reshaper
    from bidi.algorithm import get_display
except Exception:
    arabic_reshaper = None
    get_display = None

# -----------------------------
# Page setup / colors
# -----------------------------
st.set_page_config(
    page_title="STC Quality Executive Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

PURPLE = "#5A0AA2"
PURPLE2 = "#7B1FE6"
TEAL = "#2EC4D3"
GREEN = "#19B985"
PINK = "#EF476F"
ORANGE = "#FF7A45"
GOLD = "#F7B500"
DARK = "#201635"
MUTED = "#6B6380"
LIGHT_BG = "#F7F4FB"
BORDER = "#E5DDF1"
CARD = "#FFFFFF"

DATA_FILE = "Deviation.xlsx"

# -----------------------------
# Utility functions
# -----------------------------
def _norm_col(col: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(col).strip().lower())


def find_col(df: pd.DataFrame, candidates, required=False):
    mapping = {_norm_col(c): c for c in df.columns}
    for cand in candidates:
        key = _norm_col(cand)
        if key in mapping:
            return mapping[key]
    # loose match
    for cand in candidates:
        key = _norm_col(cand)
        for k, v in mapping.items():
            if key in k or k in key:
                return v
    if required:
        raise ValueError(f"Required column not found. Tried: {candidates}")
    return None


def as_text(series):
    return series.fillna("—").astype(str).str.strip().replace({"": "—", "nan": "—", "None": "—"})


def yes_flag(series):
    s = series.fillna("").astype(str).str.strip().str.upper()
    return s.isin(["Y", "YES", "TRUE", "1", "APPLIED", "PENALTY"])


def nonblank_flag(series):
    s = series.fillna("").astype(str).str.strip().str.upper()
    return ~s.isin(["", "N", "NO", "FALSE", "0", "NAN", "NONE", "—", "-"])


def classify_work(row):
    text = " ".join([
        str(row.get("Designation", "")),
        str(row.get("Category", "")),
        str(row.get("SubCategory", "")),
        str(row.get("DeviationName", "")),
    ]).upper()
    if any(x in text for x in ["FIBER", "FIBRE", "CABLE", "SPLIC"]):
        return "Fiber"
    if any(x in text for x in ["SAFETY", "PPE", "UNIFORM", "WORKER", "STC ID", "BADGE"]):
        return "Safety"
    return "Civil"


def short_deviation(name):
    if pd.isna(name):
        return "Unknown Deviation"
    t = str(name).upper().replace("\n", " ").strip()
    patterns = [
        ("DAMAGE TO PROPERTY", "Damage to Property"),
        ("NO DEBRIS", "Site Housekeeping"),
        ("SAFETY MEASUREMENTS", "STC Safety Measures"),
        ("MUNICIPALITY PERMITS ARE VALID", "Municipality Permit Validity"),
        ("MUNCIPALITY PERMIT", "Municipality Permit Signboard"),
        ("MUNICIPALITY PERMIT", "Municipality Permit"),
        ("PEDESTRIAN PASSES", "Pedestrian Crossing"),
        ("CONTRACTOR DETAILS", "Contractor Details Signboard"),
        ("CONTRACTOR'S SIGNBOARD", "Signboard Compliance"),
        ("SIGNBOARDS AVAILABLE", "Site Signboards"),
        ("WORKER UNIFORM", "PPE Compliance"),
        ("STC ID BADGE", "STC ID Badge"),
        ("UNREGISTERED WORKER", "Unregistered Worker"),
        ("UNCERTIFIED WORKER", "Uncertified Worker"),
        ("BACKFILL MATERIAL", "Backfill Material"),
        ("HANDHOLES / MANHOLES", "Manhole Installation"),
        ("HANDHOLE", "Manhole / Handhole"),
        ("TRENCH ROUTE", "Trench Alignment"),
        ("TRENCH WIDTH", "Trench Width"),
        ("TRENCH DEPTH", "Trench Depth"),
        ("CONDUITS", "Conduit Installation"),
        ("SPACERS", "Spacer Installation"),
        ("MARKING TAPE", "Warning Tape"),
        ("DETECTABLE EMS", "Detectable EMS"),
        ("TRENCH BEEN CLOSED", "Trench Closure"),
        ("REINSTATEMENT", "Reinstatement Delay"),
        ("CABLES ROUTED", "Cable Routing"),
        ("U-GUARDS", "U-Guard Installation"),
        ("POWER SOURCE", "Power Source"),
    ]
    for key, val in patterns:
        if key in t:
            return val
    # fallback title-case shortened
    cleaned = re.sub(r"\s+", " ", str(name)).strip(" ?")
    return cleaned[:45] + ("..." if len(cleaned) > 45 else "")


def first_valid(series):
    vals = [str(x).strip() for x in series.dropna().tolist() if str(x).strip() and str(x).strip().lower() not in ["nan", "none", "—", "-"]]
    return vals[0] if vals else "—"


def top_devs_text(group, top_n=3):
    counts = group["DeviationShort"].value_counts().head(top_n)
    if counts.empty:
        return "—"
    return " | ".join([f"{idx} ({int(val)})" for idx, val in counts.items()])


def register_pdf_font():
    # Do not bundle fonts. Use system font when available for Arabic names.
    candidates = [
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
        "/usr/share/fonts/truetype/freefont/FreeSans.ttf",
    ]
    for f in candidates:
        if os.path.exists(f):
            try:
                pdfmetrics.registerFont(TTFont("ReportFont", f))
                return "ReportFont"
            except Exception:
                pass
    return "Helvetica"


def rtl_text(text):
    if text is None:
        return "—"
    text = str(text)
    if arabic_reshaper and get_display and re.search(r"[\u0600-\u06FF]", text):
        try:
            return get_display(arabic_reshaper.reshape(text))
        except Exception:
            return text
    return text


@st.cache_data
def load_data():
    base_dir = os.path.dirname(__file__)
    file_path = os.path.join(base_dir, DATA_FILE)
    if not os.path.exists(file_path):
        st.error(f"Missing required file: {DATA_FILE}")
        st.stop()

    raw = pd.read_excel(file_path)

    col_map = {
        "District": find_col(raw, ["District"], required=True),
        "WorkOrder": find_col(raw, ["WorkOrderNum", "Work Order", "WO", "WorkOrder"], required=True),
        "DeviationName": find_col(raw, ["DeviationName", "Deviation Name", "Deviation"], required=True),
        "Category": find_col(raw, ["Category"], required=False),
        "SubCategory": find_col(raw, ["SubCategory", "Sub Category"], required=False),
        "Designation": find_col(raw, ["Designation", "Classification", "Class"], required=False),
        "IsPenalty": find_col(raw, ["IsPenalty", "Penalty", "Penalty Applied"], required=False),
        "ExpectedPenalties": find_col(raw, ["Expected Penalties", "ExpectedPenalty", "Expected Penalty"], required=False),
        "CivilForeman": find_col(raw, ["Civil Foreman", "Foreman"], required=False),
        "Inspector": find_col(raw, ["Inspector"], required=False),
        "DateOfDeviation": find_col(raw, ["DateOfDeviation", "Date Of Deviation", "Date"], required=False),
    }

    df = pd.DataFrame()
    df["District"] = as_text(raw[col_map["District"]]).str.upper()
    df["WorkOrder"] = raw[col_map["WorkOrder"]].fillna("—").astype(str).str.replace(r"\.0$", "", regex=True).str.strip()
    df["DeviationName"] = as_text(raw[col_map["DeviationName"]])
    df["Category"] = as_text(raw[col_map["Category"]]) if col_map["Category"] else "—"
    df["SubCategory"] = as_text(raw[col_map["SubCategory"]]) if col_map["SubCategory"] else "—"
    df["Designation"] = as_text(raw[col_map["Designation"]]) if col_map["Designation"] else "—"
    df["Civil Foreman"] = as_text(raw[col_map["CivilForeman"]]) if col_map["CivilForeman"] else "—"
    df["Inspector"] = as_text(raw[col_map["Inspector"]]) if col_map["Inspector"] else "—"
    df["DateOfDeviation"] = raw[col_map["DateOfDeviation"]] if col_map["DateOfDeviation"] else pd.NaT
    df["Penalty Applied Flag"] = yes_flag(raw[col_map["IsPenalty"]]) if col_map["IsPenalty"] else False
    df["Expected Penalties Waived Flag"] = nonblank_flag(raw[col_map["ExpectedPenalties"]]) if col_map["ExpectedPenalties"] else False
    df["No Penalty Flag"] = ~df["Penalty Applied Flag"]
    df["Classification"] = df.apply(classify_work, axis=1)
    df["DeviationShort"] = df["DeviationName"].apply(short_deviation)
    return df


def build_wo_summary(df):
    if df.empty:
        return pd.DataFrame()

    rows = []
    total_all = len(df)
    for wo, g in df.groupby("WorkOrder", dropna=False):
        total = len(g)
        rows.append({
            "Work Order": wo,
            "District": first_valid(g["District"]),
            "Civil Foreman": first_valid(g["Civil Foreman"]),
            "Inspector": first_valid(g["Inspector"]),
            "Total Deviations": total,
            "Penalty Applied": int(g["Penalty Applied Flag"].sum()),
            "No Penalty": int(g["No Penalty Flag"].sum()),
            "Expected Penalties Waived": int(g["Expected Penalties Waived Flag"].sum()),
            "Civil": int((g["Classification"] == "Civil").sum()),
            "Fiber": int((g["Classification"] == "Fiber").sum()),
            "Safety": int((g["Classification"] == "Safety").sum()),
            "% of Total": round(total / total_all * 100, 1) if total_all else 0,
            "Top 3 Deviations": top_devs_text(g, 3),
        })
    summary = pd.DataFrame(rows)
    return summary.sort_values(["Total Deviations", "Penalty Applied", "Expected Penalties Waived"], ascending=False).reset_index(drop=True)


def apply_filters(df):
    with st.sidebar:
        st.markdown("### Filters")
        districts = st.multiselect("District", sorted(df["District"].dropna().unique()))
        wos = st.multiselect("Work Order", sorted(df["WorkOrder"].dropna().unique()))
        foremen = st.multiselect("Civil Foreman", sorted(df["Civil Foreman"].dropna().unique()))
        inspectors = st.multiselect("Inspector", sorted(df["Inspector"].dropna().unique()))
        classes = st.multiselect("Classification", ["Civil", "Fiber", "Safety"])
        penalty_status = st.multiselect("Penalty Status", ["Penalty Applied", "No Penalty", "Expected Penalties Waived"])

    out = df.copy()
    if districts:
        out = out[out["District"].isin(districts)]
    if wos:
        out = out[out["WorkOrder"].isin(wos)]
    if foremen:
        out = out[out["Civil Foreman"].isin(foremen)]
    if inspectors:
        out = out[out["Inspector"].isin(inspectors)]
    if classes:
        out = out[out["Classification"].isin(classes)]
    if penalty_status:
        mask = pd.Series(False, index=out.index)
        if "Penalty Applied" in penalty_status:
            mask |= out["Penalty Applied Flag"]
        if "No Penalty" in penalty_status:
            mask |= out["No Penalty Flag"]
        if "Expected Penalties Waived" in penalty_status:
            mask |= out["Expected Penalties Waived Flag"]
        out = out[mask]
    return out


def metric_card(label, value, subtext, color=PURPLE):
    st.markdown(
        f"""
        <div class='metric-card'>
            <div class='metric-label'>{label}</div>
            <div class='metric-value' style='color:{color}'>{value}</div>
            <div class='metric-sub'>{subtext}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def make_donut(labels, values, colors_list, title=""):
    fig = go.Figure(data=[go.Pie(
        labels=labels,
        values=values,
        hole=.58,
        marker=dict(colors=colors_list, line=dict(color="#FFFFFF", width=3)),
        textinfo="label+value+percent",
        textposition="outside",
        pull=[0.02] * len(labels),
    )])
    fig.update_layout(
        title=title,
        height=430,
        margin=dict(l=20, r=20, t=50, b=30),
        legend=dict(orientation="h", y=-0.1, x=0.5, xanchor="center"),
        font=dict(family="Arial", size=13, color=DARK),
    )
    return fig


def wrap_table_dataframe(df, max_rows=None):
    show = df.copy()
    if max_rows:
        show = show.head(max_rows)
    return show


# -----------------------------
# PDF generation
# -----------------------------
def fig_to_image(fig, w=8, h=4, dpi=180):
    buf = BytesIO()
    fig.savefig(buf, format="png", dpi=dpi, bbox_inches="tight", facecolor="white")
    plt.close(fig)
    buf.seek(0)
    return buf


def pdf_kpi_cards(total_dev, unique_wos, penalty, no_penalty, waived, civil, fiber, safety):
    fig, ax = plt.subplots(figsize=(11.5, 2.0))
    ax.axis("off")
    cards = [
        ("Total Deviations", total_dev, PURPLE),
        ("Unique WOs", unique_wos, TEAL),
        ("Penalty Applied", penalty, PINK),
        ("No Penalty", no_penalty, GREEN),
        ("Expected Waived", waived, ORANGE),
        ("Civil / Fiber / Safety", f"{civil} / {fiber} / {safety}", GOLD),
    ]
    n = len(cards)
    for i, (label, value, color) in enumerate(cards):
        x = i / n + 0.008
        width = 1 / n - 0.016
        box = FancyBboxPatch((x, 0.1), width, 0.78, boxstyle="round,pad=0.015,rounding_size=0.04",
                             linewidth=1, edgecolor="#E6DDF4", facecolor="#FFFFFF")
        ax.add_patch(box)
        ax.text(x + 0.03, 0.62, label, ha="left", va="center", fontsize=9, color="#4B3A61", weight="bold")
        ax.text(x + 0.03, 0.36, str(value), ha="left", va="center", fontsize=18, color=color, weight="bold")
    return fig_to_image(fig, 11.5, 2.0)


def pdf_penalty_donut(penalty, no_penalty, waived):
    labels = ["Penalty Applied", "No Penalty", "Expected Penalties Waived"]
    values = [penalty, no_penalty, waived]
    colors_list = [PINK, GREEN, ORANGE]
    total = sum(values) if sum(values) else 1
    fig, ax = plt.subplots(figsize=(6.5, 4.2))
    wedges, _ = ax.pie(values, colors=colors_list, startangle=90, wedgeprops=dict(width=0.38, edgecolor="white", linewidth=2))
    ax.set(aspect="equal")
    ax.set_title("Penalty Applied vs No Penalty vs Expected Waived", fontsize=14, weight="bold", color=DARK, pad=18)
    for wedge, label, val, color in zip(wedges, labels, values, colors_list):
        angle = (wedge.theta2 + wedge.theta1) / 2.0
        x = np.cos(np.deg2rad(angle))
        y = np.sin(np.deg2rad(angle))
        xy = (0.82 * x, 0.82 * y)
        xytext = (1.42 * np.sign(x), 1.18 * y)
        ha = "left" if x >= 0 else "right"
        pct = val / total * 100
        ax.annotate(f"{label}\n{val:,} ({pct:.1f}%)", xy=xy, xytext=xytext, ha=ha, va="center",
                    fontsize=9, color=DARK, arrowprops=dict(arrowstyle="-", color="#8A809B", lw=1.0,
                    connectionstyle="angle3,angleA=0,angleB=90"),
                    bbox=dict(boxstyle="round,pad=0.25", fc="white", ec="#E6DDF4", lw=0.8))
    ax.set_xlim(-1.75, 1.75)
    ax.set_ylim(-1.35, 1.35)
    return fig_to_image(fig, 6.5, 4.2)


def pdf_bar_chart(data, x_col, y_col, title, color=PURPLE, horizontal=True, top_n=10):
    plot_df = data.head(top_n).copy()
    fig, ax = plt.subplots(figsize=(6.5, 4.2))
    if plot_df.empty:
        ax.text(0.5, 0.5, "No data", ha="center", va="center")
        ax.axis("off")
    elif horizontal:
        p = plot_df.iloc[::-1]
        ax.barh(p[y_col].astype(str), p[x_col], color=color, alpha=0.92)
        for i, v in enumerate(p[x_col]):
            ax.text(v + max(plot_df[x_col].max() * 0.015, 0.5), i, str(int(v)), va="center", fontsize=8, color=DARK)
        ax.set_xlabel(x_col, fontsize=9, color=MUTED)
    else:
        ax.bar(plot_df[x_col].astype(str), plot_df[y_col], color=color, alpha=0.92)
    ax.set_title(title, fontsize=14, weight="bold", color=DARK)
    ax.spines[["top", "right", "left", "bottom"]].set_visible(False)
    ax.grid(axis="x" if horizontal else "y", color="#ECE7F4", lw=0.8)
    ax.tick_params(axis="both", labelsize=8, colors=MUTED)
    return fig_to_image(fig, 6.5, 4.2)


def para(text, style):
    return Paragraph(rtl_text(text), style)


def make_pdf(df, summary):
    font_name = register_pdf_font()
    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=landscape(A4),
        rightMargin=0.8 * cm,
        leftMargin=0.8 * cm,
        topMargin=0.7 * cm,
        bottomMargin=0.7 * cm,
    )

    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="TitleWhite", fontName=font_name, fontSize=22, leading=26, textColor=colors.white, alignment=TA_LEFT, spaceAfter=4))
    styles.add(ParagraphStyle(name="SubWhite", fontName=font_name, fontSize=9.5, leading=12, textColor=colors.white, alignment=TA_LEFT))
    styles.add(ParagraphStyle(name="H1Custom", fontName=font_name, fontSize=16, leading=20, textColor=HexColor(PURPLE), alignment=TA_LEFT, spaceAfter=8))
    styles.add(ParagraphStyle(name="BodyCustom", fontName=font_name, fontSize=8.3, leading=10, textColor=HexColor(DARK), alignment=TA_LEFT))
    styles.add(ParagraphStyle(name="Cell", fontName=font_name, fontSize=6.5, leading=7.4, textColor=HexColor(DARK), alignment=TA_LEFT))
    styles.add(ParagraphStyle(name="CellCenter", fontName=font_name, fontSize=6.5, leading=7.4, textColor=HexColor(DARK), alignment=TA_CENTER))
    styles.add(ParagraphStyle(name="HeaderCell", fontName=font_name, fontSize=6.4, leading=7.2, textColor=colors.white, alignment=TA_CENTER))

    story = []

    header_table = Table([
        [para("STC Quality Executive Dashboard", styles["TitleWhite"])],
        [para("Deviation analysis - Work Order performance - Penalty applied vs waived/not applied - Civil / Fiber / Safety classification", styles["SubWhite"])]
    ], colWidths=[27.8 * cm])
    header_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), HexColor(PURPLE)),
        ("BOX", (0, 0), (-1, -1), 0, HexColor(PURPLE)),
        ("LEFTPADDING", (0, 0), (-1, -1), 14),
        ("RIGHTPADDING", (0, 0), (-1, -1), 14),
        ("TOPPADDING", (0, 0), (-1, -1), 10),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
    ]))
    story.append(header_table)
    story.append(Spacer(1, 0.28 * cm))

    total_dev = len(df)
    unique_wos = df["WorkOrder"].nunique()
    penalty = int(df["Penalty Applied Flag"].sum())
    no_penalty = int(df["No Penalty Flag"].sum())
    waived = int(df["Expected Penalties Waived Flag"].sum())
    civil = int((df["Classification"] == "Civil").sum())
    fiber = int((df["Classification"] == "Fiber").sum())
    safety = int((df["Classification"] == "Safety").sum())

    story.append(Image(pdf_kpi_cards(total_dev, unique_wos, penalty, no_penalty, waived, civil, fiber, safety), width=27.4 * cm, height=4.15 * cm))
    story.append(Spacer(1, 0.15 * cm))

    # Charts row
    wo_chart_data = summary[["Work Order", "Total Deviations"]].rename(columns={"Work Order": "WO", "Total Deviations": "Deviation Count"})
    dev_chart_data = df["DeviationShort"].value_counts().head(10).reset_index()
    dev_chart_data.columns = ["Deviation", "Count"]
    donut = Image(pdf_penalty_donut(penalty, no_penalty, waived), width=12.9 * cm, height=8.2 * cm)
    wo_bar = Image(pdf_bar_chart(wo_chart_data, "Deviation Count", "WO", "Top Work Orders by Deviation Count", PURPLE), width=12.9 * cm, height=8.2 * cm)
    chart_table = Table([[wo_bar, donut]], colWidths=[13.6 * cm, 13.6 * cm])
    chart_table.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "MIDDLE")]))
    story.append(chart_table)

    story.append(PageBreak())
    story.append(para("Executive Work Order Summary", styles["H1Custom"]))
    story.append(para("Each row represents one Work Order. Service Affecting and Status are intentionally excluded from the board report.", styles["BodyCustom"]))
    story.append(Spacer(1, 0.15 * cm))

    display_cols = [
        "Work Order", "District", "Civil Foreman", "Inspector", "Total Deviations", "Penalty Applied",
        "No Penalty", "Expected Penalties Waived", "Civil", "Fiber", "Safety", "Top 3 Deviations"
    ]
    pdf_summary = summary[display_cols].copy()
    headers = ["Work Order", "District", "Civil Foreman", "Inspector", "Total", "Penalty", "No Penalty", "Expected Waived", "Civil", "Fiber", "Safety", "Top 3 Deviations"]

    table_data = [[para(h, styles["HeaderCell"]) for h in headers]]
    for _, row in pdf_summary.iterrows():
        r = []
        for c in display_cols:
            sty = styles["CellCenter"] if c not in ["Civil Foreman", "Inspector", "Top 3 Deviations"] else styles["Cell"]
            r.append(para(row[c], sty))
        table_data.append(r)

    col_widths = [1.9*cm, 1.4*cm, 2.5*cm, 2.3*cm, 1.2*cm, 1.25*cm, 1.25*cm, 1.55*cm, 0.9*cm, 0.9*cm, 0.95*cm, 9.0*cm]
    t = Table(table_data, colWidths=col_widths, repeatRows=1)
    style_cmds = [
        ("BACKGROUND", (0, 0), (-1, 0), HexColor(PURPLE)),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("GRID", (0, 0), (-1, -1), 0.35, HexColor("#D9D0E9")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 3),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]
    for i in range(1, len(table_data)):
        style_cmds.append(("BACKGROUND", (0, i), (-1, i), HexColor("#F6F1FA") if i % 2 == 0 else colors.white))
    # highlight key columns
    idx_penalty = headers.index("Penalty")
    idx_waived = headers.index("Expected Waived")
    idx_no = headers.index("No Penalty")
    for r in range(1, len(table_data)):
        style_cmds.append(("BACKGROUND", (idx_penalty, r), (idx_penalty, r), HexColor("#FFE9EF")))
        style_cmds.append(("BACKGROUND", (idx_waived, r), (idx_waived, r), HexColor("#FFF0E8")))
        style_cmds.append(("BACKGROUND", (idx_no, r), (idx_no, r), HexColor("#E7FAF3")))
    t.setStyle(TableStyle(style_cmds))
    story.append(t)

    story.append(PageBreak())
    story.append(para("Deviation Distribution Analysis", styles["H1Custom"]))
    dev_bar = Image(pdf_bar_chart(dev_chart_data, "Count", "Deviation", "Top Deviations - Short Names", PURPLE2), width=13.2 * cm, height=8.0 * cm)
    class_data = df["Classification"].value_counts().reset_index()
    class_data.columns = ["Class", "Count"]
    class_bar = Image(pdf_bar_chart(class_data.rename(columns={"Class": "Class"}), "Count", "Class", "Civil / Fiber / Safety Distribution", TEAL), width=13.2 * cm, height=8.0 * cm)
    story.append(Table([[dev_bar, class_bar]], colWidths=[13.8 * cm, 13.8 * cm]))

    story.append(Spacer(1, 0.4 * cm))
    story.append(para("Top Civil Foremen / Inspectors", styles["H1Custom"]))
    foreman = df["Civil Foreman"].value_counts().head(10).reset_index()
    foreman.columns = ["Civil Foreman", "Deviation Count"]
    inspector = df["Inspector"].value_counts().head(10).reset_index()
    inspector.columns = ["Inspector", "Deviation Count"]
    small_headers = [[para("Civil Foreman", styles["HeaderCell"]), para("Deviation Count", styles["HeaderCell"]), para("Inspector", styles["HeaderCell"]), para("Deviation Count", styles["HeaderCell"])]]
    max_rows = max(len(foreman), len(inspector))
    small_rows = []
    for i in range(max_rows):
        small_rows.append([
            para(foreman.iloc[i, 0] if i < len(foreman) else "", styles["Cell"]),
            para(int(foreman.iloc[i, 1]) if i < len(foreman) else "", styles["CellCenter"]),
            para(inspector.iloc[i, 0] if i < len(inspector) else "", styles["Cell"]),
            para(int(inspector.iloc[i, 1]) if i < len(inspector) else "", styles["CellCenter"]),
        ])
    people_table = Table(small_headers + small_rows, colWidths=[8.0 * cm, 3.0 * cm, 8.0 * cm, 3.0 * cm], repeatRows=1)
    people_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), HexColor(PURPLE)),
        ("GRID", (0, 0), (-1, -1), 0.35, HexColor("#D9D0E9")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    story.append(people_table)

    doc.build(story)
    buffer.seek(0)
    return buffer


# -----------------------------
# Dashboard UI
# -----------------------------
st.markdown(
    f"""
    <style>
        .block-container {{padding-top: 1rem; padding-bottom: 1rem; max-width: 1500px;}}
        .hero {{
            background: linear-gradient(100deg, {PURPLE}, {PURPLE2});
            padding: 26px 30px; border-radius: 0 0 22px 22px;
            color: white; box-shadow: 0 12px 30px rgba(90,10,162,.22); margin-bottom: 18px;
        }}
        .hero h1 {{margin:0; font-size: 34px; line-height:1.1;}}
        .hero p {{margin:12px 0 0 0; font-size: 14px; opacity:.95;}}
        .metric-card {{
            background: white; border:1px solid {BORDER}; border-radius: 15px; padding: 18px 18px;
            min-height: 118px; box-shadow: 0 6px 20px rgba(61,30,95,.06);
        }}
        .metric-label {{font-size: 13px; color:{DARK}; font-weight:700;}}
        .metric-value {{font-size: 32px; font-weight:800; margin-top:8px;}}
        .metric-sub {{font-size: 12px; color:{MUTED}; margin-top:8px;}}
        div[data-testid="stDataFrame"] {{border:1px solid {BORDER}; border-radius:14px; overflow:hidden;}}
        .section-title {{font-size: 22px; font-weight: 800; color:{DARK}; margin: 18px 0 12px 0;}}
    </style>
    <div class="hero">
        <h1>STC Quality Executive Dashboard</h1>
        <p>Executive Work Order summary • Penalty applied vs waived/not applied • Civil / Fiber / Safety classification • Board-ready PDF export</p>
    </div>
    """,
    unsafe_allow_html=True,
)

base_df = load_data()
filtered = apply_filters(base_df)
summary = build_wo_summary(filtered)

# KPI row
st.markdown('<div class="section-title">Executive KPIs</div>', unsafe_allow_html=True)
cols = st.columns(8)
total = len(filtered)
unique_wos = filtered["WorkOrder"].nunique()
penalty_applied = int(filtered["Penalty Applied Flag"].sum())
no_penalty = int(filtered["No Penalty Flag"].sum())
waived = int(filtered["Expected Penalties Waived Flag"].sum())
civil = int((filtered["Classification"] == "Civil").sum())
fiber = int((filtered["Classification"] == "Fiber").sum())
safety = int((filtered["Classification"] == "Safety").sum())

with cols[0]: metric_card("Total Deviations", f"{total:,}", "All records after filters", PURPLE)
with cols[1]: metric_card("Unique WOs", f"{unique_wos:,}", "Impacted Work Orders", TEAL)
with cols[2]: metric_card("Penalty Applied", f"{penalty_applied:,}", f"{penalty_applied / total * 100:.1f}%" if total else "0%", PINK)
with cols[3]: metric_card("No Penalty", f"{no_penalty:,}", f"{no_penalty / total * 100:.1f}%" if total else "0%", GREEN)
with cols[4]: metric_card("Expected Waived", f"{waived:,}", "Should be penalty but cancelled", ORANGE)
with cols[5]: metric_card("Civil", f"{civil:,}", "Civil deviations", "#8B5E34")
with cols[6]: metric_card("Fiber", f"{fiber:,}", "Fiber deviations", TEAL)
with cols[7]: metric_card("Safety", f"{safety:,}", "Safety deviations", GOLD)

st.markdown('<div class="section-title">Executive Visualizations</div>', unsafe_allow_html=True)
left, right = st.columns([1.05, 1])
with left:
    if not summary.empty:
        fig = px.bar(
            summary.head(10).sort_values("Total Deviations"),
            x="Total Deviations", y="Work Order", orientation="h",
            color="District", text="Total Deviations",
            color_discrete_sequence=[PURPLE, TEAL, ORANGE, PINK],
            title="Top Work Orders by Deviation Count",
        )
        fig.update_traces(textposition="outside")
        fig.update_layout(height=440, margin=dict(l=10, r=20, t=60, b=20), font=dict(color=DARK), plot_bgcolor="white")
        st.plotly_chart(fig, use_container_width=True)
with right:
    fig2 = make_donut(
        ["Penalty Applied", "No Penalty", "Expected Penalties Waived"],
        [penalty_applied, no_penalty, waived],
        [PINK, GREEN, ORANGE],
        title="Penalty Applied vs Waived/Not Applied",
    )
    st.plotly_chart(fig2, use_container_width=True)

left2, right2 = st.columns([1, 1])
with left2:
    dev_counts = filtered["DeviationShort"].value_counts().head(12).reset_index()
    dev_counts.columns = ["Deviation", "Count"]
    fig3 = px.bar(dev_counts.sort_values("Count"), x="Count", y="Deviation", orientation="h", text="Count", title="Top Deviation Types - Short Names", color_discrete_sequence=[PURPLE2])
    fig3.update_traces(textposition="outside")
    fig3.update_layout(height=480, margin=dict(l=10, r=20, t=60, b=20), plot_bgcolor="white")
    st.plotly_chart(fig3, use_container_width=True)
with right2:
    class_counts = filtered["Classification"].value_counts().reset_index()
    class_counts.columns = ["Classification", "Count"]
    fig4 = px.bar(class_counts, x="Classification", y="Count", color="Classification", text="Count", title="Civil / Fiber / Safety Classification", color_discrete_sequence=["#8B5E34", TEAL, GOLD])
    fig4.update_traces(textposition="outside")
    fig4.update_layout(height=480, margin=dict(l=10, r=20, t=60, b=20), plot_bgcolor="white", showlegend=False)
    st.plotly_chart(fig4, use_container_width=True)

st.markdown('<div class="section-title">Executive Work Order Summary</div>', unsafe_allow_html=True)
st.caption("Aggregated table: one row per Work Order. Service Affecting and Status are intentionally excluded.")
show_cols = [
    "Work Order", "District", "Civil Foreman", "Inspector", "Total Deviations", "Penalty Applied",
    "No Penalty", "Expected Penalties Waived", "Civil", "Fiber", "Safety", "% of Total", "Top 3 Deviations"
]
st.dataframe(summary[show_cols], use_container_width=True, hide_index=True, height=420)

with st.expander("Detailed Deviation Records - Optional operational drilldown", expanded=False):
    detail_cols = ["WorkOrder", "District", "Civil Foreman", "Inspector", "Classification", "Category", "SubCategory", "DeviationShort", "DeviationName", "Penalty Applied Flag", "No Penalty Flag", "Expected Penalties Waived Flag"]
    st.dataframe(filtered[detail_cols], use_container_width=True, hide_index=True, height=400)

st.markdown('<div class="section-title">Export</div>', unsafe_allow_html=True)
pdf_buffer = make_pdf(filtered, summary)
st.download_button(
    label="Export Board-ready PDF",
    data=pdf_buffer,
    file_name=f"STC_Quality_Executive_Board_Report_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
    mime="application/pdf",
    type="primary",
)

st.caption("Prepared for Quality Director / Executive Management. Data source: Deviation.xlsx")
