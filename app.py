
import io
import math
import os
import re
from datetime import datetime

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import plotly.graph_objects as go
import streamlit as st
from reportlab.lib import colors
from reportlab.lib.colors import HexColor
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import A3, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak,
    Image as RLImage, KeepTogether
)

# ============================================================
# CONFIG
# ============================================================
st.set_page_config(
    page_title="STC Quality Executive Dashboard",
    layout="wide",
    initial_sidebar_state="expanded"
)

PURPLE = "#5A0AA2"
PURPLE_2 = "#7B1FE8"
TEAL = "#2EC4D3"
PINK = "#EF476F"
GREEN = "#18B77B"
ORANGE = "#FF7043"
YELLOW = "#F4B400"
DARK = "#1E1733"
MUTED = "#6F6A85"
CARD_BG = "#FFFFFF"
SOFT_BG = "#F8F5FC"
LINE = "#E7DDF6"

# ============================================================
# DATA
# ============================================================
@st.cache_data
def load_data():
    xlsx_path = os.path.join(os.path.dirname(__file__), "Deviation.xlsx")
    xls = pd.ExcelFile(xlsx_path)
    # Prefer raw detail sheet, fallback to first sheet
    sheet = "Sheet2" if "Sheet2" in xls.sheet_names else xls.sheet_names[0]
    df = pd.read_excel(xlsx_path, sheet_name=sheet)

    # Standardize expected columns
    df.columns = [str(c).strip() for c in df.columns]
    required = [
        "WorkOrderNum", "District", "DeviationName", "Designation",
        "IsPenalty", "Civil Foreman", "Inspector", "Expected Penalties"
    ]
    for c in required:
        if c not in df.columns:
            df[c] = ""

    df["WorkOrderNum"] = df["WorkOrderNum"].astype(str).str.replace(".0", "", regex=False).str.strip()
    df["District"] = df["District"].astype(str).str.strip().str.upper()
    df["DeviationName"] = df["DeviationName"].astype(str).str.strip()
    df["Designation"] = df["Designation"].astype(str).str.strip().str.title().replace({"Safety": "Safety", "Civil": "Civil", "Fiber": "Fiber"})
    df["IsPenalty"] = df["IsPenalty"].astype(str).str.strip().str.upper()
    df["Expected Penalties"] = df["Expected Penalties"].fillna("").astype(str).str.strip().str.upper()
    df["Civil Foreman"] = df["Civil Foreman"].fillna("—").astype(str).str.strip().replace({"": "—", "nan": "—"})
    df["Inspector"] = df["Inspector"].fillna("—").astype(str).str.strip().replace({"": "—", "nan": "—"})

    df["Penalty Applied Flag"] = df["IsPenalty"].eq("Y").astype(int)
    df["No Penalty Flag"] = df["IsPenalty"].ne("Y").astype(int)
    # Expected Penalties = was expected to apply penalty but it was waived/cancelled
    df["Expected Penalty Waived Flag"] = df["Expected Penalties"].isin(["YES", "Y", "TRUE", "1"]).astype(int)

    df["Civil Flag"] = df["Designation"].eq("Civil").astype(int)
    df["Fiber Flag"] = df["Designation"].eq("Fiber").astype(int)
    df["Safety Flag"] = df["Designation"].eq("Safety").astype(int)

    df["Civil Penalty Flag"] = ((df["Designation"].eq("Civil")) & (df["Penalty Applied Flag"].eq(1))).astype(int)
    df["Fiber Penalty Flag"] = ((df["Designation"].eq("Fiber")) & (df["Penalty Applied Flag"].eq(1))).astype(int)
    df["Safety Penalty Flag"] = ((df["Designation"].eq("Safety")) & (df["Penalty Applied Flag"].eq(1))).astype(int)

    df["Short Deviation"] = df["DeviationName"].apply(shorten_deviation)
    return df


def shorten_deviation(name: str) -> str:
    s = str(name).upper()
    rules = [
        ("DAMAGE TO PROPERTY", "Damage to Property"),
        ("NO DEBRIS", "Site Housekeeping"),
        ("SAFETY MEASUREMENTS", "STC Safety Measures"),
        ("MUNICIPALITY PERMITS", "Municipality Permit"),
        ("MUNICIPALITY PERMIT IS AVAILABLE", "Municipality Permit"),
        ("PEDESTRIAN", "Pedestrian Crossing"),
        ("REINSTATEMENT", "Reinstatement Delay"),
        ("MARKING TAPE", "Warning Tape"),
        ("SIGNBOARD", "Site Signboard"),
        ("WORKER UNIFORM", "PPE Compliance"),
        ("UNCERTIFIED WORKER", "Uncertified Worker"),
        ("BACKFILL", "Backfill Material"),
        ("U-GUARDS", "U-Guard Installation"),
        ("CABLES ROUTED", "Cable Routing"),
        ("MANHOLES / HANDHOLES ARE CLEAN", "Manhole Cleanliness"),
        ("STAND ALONE POWER SOURCE", "Power Source"),
    ]
    for key, val in rules:
        if key in s:
            return val
    cleaned = re.sub(r"\s+", " ", str(name).title()).strip(" ?")
    return cleaned[:42] + "..." if len(cleaned) > 45 else cleaned


def unique_join(series):
    vals = [str(v).strip() for v in series.dropna().astype(str) if str(v).strip() and str(v).strip().lower() != "nan"]
    vals = list(dict.fromkeys(vals))
    return " / ".join(vals) if vals else "—"


def top_deviations_text(g):
    vc = g["Short Deviation"].value_counts().head(3)
    return " | ".join([f"{idx} ({int(val)})" for idx, val in vc.items()])


def build_workorder_summary(df):
    total_all = len(df)
    rows = []
    for wo, g in df.groupby("WorkOrderNum", dropna=False):
        total = len(g)
        rows.append({
            "WorkOrderNum": wo,
            "Total Deviation": total,
            "Penalty applied": int(g["Penalty Applied Flag"].sum()),
            "No penalty applied": int(g["No Penalty Flag"].sum()),
            "Expected Penalty Waived": int(g["Expected Penalty Waived Flag"].sum()),
            "Civil Deviation": int(g["Civil Flag"].sum()),
            "Civil Penalty applied": int(g["Civil Penalty Flag"].sum()),
            "Fiber Deviation": int(g["Fiber Flag"].sum()),
            "Fiber Penalty applied": int(g["Fiber Penalty Flag"].sum()),
            "Safety Deviation": int(g["Safety Flag"].sum()),
            "Safety Penalty applied": int(g["Safety Penalty Flag"].sum()),
            "% OF Total": (total / total_all * 100) if total_all else 0,
            "Civil Foreman": unique_join(g["Civil Foreman"]),
            "Inspector": unique_join(g["Inspector"]),
            "District": unique_join(g["District"]),
            "Top 3 Deviations": top_deviations_text(g),
        })
    out = pd.DataFrame(rows).sort_values(["Total Deviation", "Penalty applied", "Expected Penalty Waived"], ascending=False)
    out["% OF Total"] = out["% OF Total"].map(lambda x: f"{x:.2f}%")
    return out


# ============================================================
# VISUAL HELPERS
# ============================================================
def kpi_card(title, value, subtitle="", color=PURPLE, icon=""):
    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-title">{icon} {title}</div>
            <div class="kpi-value" style="color:{color};">{value}</div>
            <div class="kpi-subtitle">{subtitle}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def fig_penalty_donut(applied, no_penalty, expected_waived):
    labels = ["Penalty Applied", "No Penalty", "Expected Penalty Waived"]
    values = [applied, no_penalty, expected_waived]
    colors_list = [PINK, GREEN, ORANGE]
    total = sum(values) if sum(values) else 1
    text = [f"{l}<br>{v}<br>{v/total:.1%}" for l, v in zip(labels, values)]
    fig = go.Figure(
        data=[go.Pie(
            labels=labels,
            values=values,
            hole=.62,
            marker=dict(colors=colors_list, line=dict(color="white", width=3)),
            text=text,
            textinfo="text",
            textposition="outside",
            pull=[0.02, 0.02, 0.02],
            showlegend=True
        )]
    )
    fig.update_layout(
        title=dict(text="Penalty Applied vs No Penalty vs Waived", x=0.02, xanchor="left"),
        height=430,
        margin=dict(t=60, b=40, l=30, r=30),
        legend=dict(orientation="h", y=-0.05, x=0.05),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
    )
    return fig


def fig_top_workorders(summary):
    plot = summary.head(10).copy()
    plot["PercentNumeric"] = plot["% OF Total"].str.replace("%", "", regex=False).astype(float)
    fig = go.Figure()
    fig.add_trace(go.Bar(
        y=plot["WorkOrderNum"][::-1],
        x=plot["Total Deviation"][::-1],
        orientation="h",
        marker=dict(color=PURPLE),
        name="Total Deviations",
        text=plot["Total Deviation"][::-1],
        textposition="outside"
    ))
    fig.add_trace(go.Bar(
        y=plot["WorkOrderNum"][::-1],
        x=plot["Penalty applied"][::-1],
        orientation="h",
        marker=dict(color=PINK),
        name="Penalty Applied",
        text=plot["Penalty applied"][::-1],
        textposition="inside"
    ))
    fig.update_layout(
        title=dict(text="Top Work Orders by Deviation Count", x=0.02, xanchor="left"),
        barmode="overlay",
        height=430,
        margin=dict(l=20, r=30, t=60, b=40),
        xaxis_title="Deviation Count",
        yaxis_title="Work Order",
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        legend=dict(orientation="h", y=-0.15)
    )
    return fig


def wo_score_card(row):
    st.markdown(
        f"""
        <div class="wo-card">
            <div class="wo-header">
                <div>
                    <div class="wo-title">WO {row['WorkOrderNum']}</div>
                    <div class="wo-sub">{row.get('District','')} • Foreman: {row['Civil Foreman']} • Inspector: {row['Inspector']}</div>
                </div>
                <div class="wo-percent">{row['% OF Total']}</div>
            </div>
            <div class="wo-metrics">
                <div class="mini-card"><span>Total</span><b style="color:{PURPLE}">{row['Total Deviation']}</b></div>
                <div class="mini-card"><span>Penalty</span><b style="color:{PINK}">{row['Penalty applied']}</b></div>
                <div class="mini-card"><span>No Penalty</span><b style="color:{GREEN}">{row['No penalty applied']}</b></div>
                <div class="mini-card"><span>Expected Waived</span><b style="color:{ORANGE}">{row['Expected Penalty Waived']}</b></div>
                <div class="mini-card"><span>Civil</span><b style="color:{PURPLE}">{row['Civil Deviation']} / {row['Civil Penalty applied']}</b></div>
                <div class="mini-card"><span>Fiber</span><b style="color:{TEAL}">{row['Fiber Deviation']} / {row['Fiber Penalty applied']}</b></div>
                <div class="mini-card"><span>Safety</span><b style="color:{YELLOW}">{row['Safety Deviation']} / {row['Safety Penalty applied']}</b></div>
            </div>
            <div class="wo-top"><b>Top Deviations:</b> {row['Top 3 Deviations']}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def styled_summary_table(df):
    cols = [
        "WorkOrderNum", "Total Deviation", "Penalty applied", "No penalty applied",
        "Expected Penalty Waived", "Civil Deviation", "Civil Penalty applied",
        "Fiber Deviation", "Fiber Penalty applied", "Safety Deviation",
        "Safety Penalty applied", "% OF Total", "Civil Foreman", "Inspector"
    ]
    return df[cols].style.set_properties(**{
        "font-size": "13px",
        "border-color": "#E8DEF8",
    }).set_table_styles([
        {"selector": "th", "props": [
            ("background-color", PURPLE),
            ("color", "white"),
            ("font-weight", "700"),
            ("text-align", "center"),
            ("border", "1px solid #FFFFFF"),
        ]},
        {"selector": "td", "props": [
            ("border", "1px solid #E8DEF8"),
            ("padding", "8px"),
        ]},
    ]).format(precision=0)


# ============================================================
# PDF EXPORT
# ============================================================
def matplotlib_donut_for_pdf(applied, no_penalty, waived):
    labels = ["Penalty Applied", "No Penalty", "Expected Waived"]
    values = [applied, no_penalty, waived]
    colors_mpl = [PINK, GREEN, ORANGE]
    total = sum(values) if sum(values) else 1

    fig, ax = plt.subplots(figsize=(6.5, 4.2), dpi=170)
    wedges, _ = ax.pie(
        values, labels=None, colors=colors_mpl, startangle=90,
        wedgeprops=dict(width=0.38, edgecolor="white", linewidth=2)
    )

    for i, w in enumerate(wedges):
        ang = (w.theta2 + w.theta1) / 2
        x = np.cos(np.deg2rad(ang))
        y = np.sin(np.deg2rad(ang))
        ha = "left" if x >= 0 else "right"
        label = f"{labels[i]}\n{values[i]} ({values[i]/total:.1%})"
        ax.annotate(
            label,
            xy=(0.78 * x, 0.78 * y),
            xytext=(1.25 * np.sign(x), 1.15 * y),
            ha=ha, va="center",
            fontsize=8.5,
            fontweight="bold",
            color=DARK,
            arrowprops=dict(arrowstyle="-", color="#6F6A85", lw=1.0, connectionstyle="angle3,angleA=0,angleB=90")
        )

    ax.text(0, 0.05, "Penalty\nStatus", ha="center", va="center", fontsize=13, fontweight="bold", color=DARK)
    ax.set_title("Penalty Applied vs No Penalty vs Waived", fontsize=13, fontweight="bold", color=DARK)
    ax.set_aspect("equal")
    plt.tight_layout()
    buf = io.BytesIO()
    fig.savefig(buf, format="png", transparent=True, bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    return buf


def matplotlib_bar_for_pdf(summary):
    plot = summary.head(10).copy().sort_values("Total Deviation")
    fig, ax = plt.subplots(figsize=(7.2, 4.2), dpi=170)
    ax.barh(plot["WorkOrderNum"], plot["Total Deviation"], color=PURPLE, label="Total")
    ax.barh(plot["WorkOrderNum"], plot["Penalty applied"], color=PINK, label="Penalty")
    for y, total, penalty in zip(plot["WorkOrderNum"], plot["Total Deviation"], plot["Penalty applied"]):
        ax.text(total + 1, y, str(int(total)), va="center", fontsize=8, color=DARK, fontweight="bold")
        if penalty > 0:
            ax.text(max(penalty / 2, 1), y, str(int(penalty)), va="center", ha="center", fontsize=8, color="white", fontweight="bold")
    ax.set_title("Top Work Orders by Deviation Count", fontsize=13, fontweight="bold", color=DARK)
    ax.set_xlabel("Deviation Count")
    ax.spines[["top", "right", "left", "bottom"]].set_visible(False)
    ax.grid(axis="x", color="#E7DDF6", linewidth=0.8)
    ax.legend(frameon=False, loc="lower right", fontsize=8)
    plt.tight_layout()
    buf = io.BytesIO()
    fig.savefig(buf, format="png", transparent=True, bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    return buf


def make_pdf(df, summary):
    buffer = io.BytesIO()
    pagesize = landscape(A3)
    doc = SimpleDocTemplate(
        buffer,
        pagesize=pagesize,
        rightMargin=0.8*cm, leftMargin=0.8*cm,
        topMargin=0.7*cm, bottomMargin=0.7*cm
    )
    W, H = pagesize

    styles = getSampleStyleSheet()
    title = ParagraphStyle(
        "Title", parent=styles["Title"], fontName="Helvetica-Bold", fontSize=22,
        textColor=HexColor(PURPLE), alignment=TA_CENTER, spaceAfter=6
    )
    subtitle = ParagraphStyle(
        "SubTitle", parent=styles["Normal"], fontName="Helvetica", fontSize=10,
        textColor=HexColor(DARK), alignment=TA_CENTER, spaceAfter=10
    )
    h2 = ParagraphStyle(
        "H2", parent=styles["Heading2"], fontName="Helvetica-Bold", fontSize=13,
        textColor=HexColor(DARK), spaceBefore=6, spaceAfter=6
    )
    normal = ParagraphStyle(
        "NormalSmall", parent=styles["Normal"], fontName="Helvetica", fontSize=7.1,
        textColor=HexColor(DARK), leading=8.5
    )

    story = []
    story.append(Paragraph("STC Quality Executive Dashboard", title))
    story.append(Paragraph("Work Order executive summary • penalty applied vs waived/not applied • Civil / Fiber / Safety classification", subtitle))

    # KPI cards
    total = len(df)
    unique_wos = summary["WorkOrderNum"].nunique()
    penalty = int(df["Penalty Applied Flag"].sum())
    no_penalty = int(df["No Penalty Flag"].sum())
    waived = int(df["Expected Penalty Waived Flag"].sum())
    civil = int(df["Civil Flag"].sum())
    fiber = int(df["Fiber Flag"].sum())
    safety = int(df["Safety Flag"].sum())

    kpis = [
        ["Total Deviations", str(total), PURPLE],
        ["Unique WOs", str(unique_wos), TEAL],
        ["Penalty Applied", str(penalty), PINK],
        ["No Penalty", str(no_penalty), GREEN],
        ["Expected Waived", str(waived), ORANGE],
        ["Civil / Fiber / Safety", f"{civil} / {fiber} / {safety}", YELLOW],
    ]
    kpi_data = []
    for label, val, color in kpis:
        kpi_data.append([
            Paragraph(f"<b>{label}</b><br/><font size='18' color='{color}'><b>{val}</b></font>", normal)
        ])
    kpi_table = Table([kpi_data], colWidths=[(W-1.6*cm)/6]*6, rowHeights=[1.55*cm])
    kpi_table.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,-1), HexColor("#FFFFFF")),
        ("BOX", (0,0), (-1,-1), 0.8, HexColor("#DDCFF5")),
        ("INNERGRID", (0,0), (-1,-1), 0.8, HexColor("#DDCFF5")),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("LEFTPADDING", (0,0), (-1,-1), 10),
    ]))
    story.append(kpi_table)
    story.append(Spacer(1, 0.3*cm))

    # charts
    donut_buf = matplotlib_donut_for_pdf(penalty, no_penalty, waived)
    bar_buf = matplotlib_bar_for_pdf(summary)
    chart_tbl = Table(
        [[RLImage(bar_buf, width=15.3*cm, height=8.1*cm), RLImage(donut_buf, width=15.3*cm, height=8.1*cm)]],
        colWidths=[(W-1.8*cm)/2]*2
    )
    chart_tbl.setStyle(TableStyle([
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("ALIGN", (0,0), (-1,-1), "CENTER")
    ]))
    story.append(chart_tbl)
    story.append(Spacer(1, 0.2*cm))
    story.append(Paragraph("Executive Work Order Summary", h2))

    # main executive table exactly requested
    main_cols = [
        "WorkOrderNum", "Total Deviation", "Penalty applied", "No penalty applied",
        "Expected Penalty Waived", "Civil Deviation", "Civil Penalty applied",
        "Fiber Deviation", "Fiber Penalty applied", "Safety Deviation",
        "Safety Penalty applied", "% OF Total", "Civil Foreman", "Inspector"
    ]
    header_labels = [
        "Work Order", "Total<br/>Deviation", "Penalty<br/>Applied", "No Penalty<br/>Applied",
        "Expected<br/>Penalty<br/>Waived", "Civil<br/>Deviation", "Civil<br/>Penalty",
        "Fiber<br/>Deviation", "Fiber<br/>Penalty", "Safety<br/>Deviation",
        "Safety<br/>Penalty", "% of<br/>Total", "Civil Foreman", "Inspector"
    ]

    pdf_df = summary[main_cols].copy()
    table_data = [[Paragraph(f"<b>{x}</b>", normal) for x in header_labels]]
    for _, r in pdf_df.iterrows():
        table_data.append([Paragraph(str(r[c]), normal) for c in main_cols])

    col_widths = [
        2.0*cm, 1.35*cm, 1.35*cm, 1.45*cm, 1.65*cm, 1.28*cm, 1.30*cm,
        1.28*cm, 1.30*cm, 1.35*cm, 1.35*cm, 1.10*cm, 2.9*cm, 2.8*cm
    ]
    table = Table(table_data, colWidths=col_widths, repeatRows=1)
    ts = [
        ("BACKGROUND", (0,0), (-1,0), HexColor(PURPLE)),
        ("TEXTCOLOR", (0,0), (-1,0), colors.white),
        ("ALIGN", (1,1), (11,-1), "CENTER"),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("GRID", (0,0), (-1,-1), 0.35, HexColor("#DDD6EA")),
        ("ROWBACKGROUNDS", (0,1), (-1,-1), [HexColor("#FFFFFF"), HexColor("#F7F2FC")]),
        ("BACKGROUND", (2,1), (2,-1), HexColor("#FFE5EE")),
        ("BACKGROUND", (3,1), (3,-1), HexColor("#EAF8F2")),
        ("BACKGROUND", (4,1), (4,-1), HexColor("#FFF1E8")),
        ("BACKGROUND", (5,1), (6,-1), HexColor("#F1E9FB")),
        ("BACKGROUND", (7,1), (8,-1), HexColor("#E7F8FB")),
        ("BACKGROUND", (9,1), (10,-1), HexColor("#FFF6D7")),
        ("LEFTPADDING", (0,0), (-1,-1), 3),
        ("RIGHTPADDING", (0,0), (-1,-1), 3),
        ("TOPPADDING", (0,0), (-1,-1), 4),
        ("BOTTOMPADDING", (0,0), (-1,-1), 4),
    ]
    table.setStyle(TableStyle(ts))
    story.append(table)

    # Per WO board-style cards
    story.append(PageBreak())
    story.append(Paragraph("Individual Work Order Performance Cards", title))
    story.append(Paragraph("Each card gives the board-ready position for one Work Order only.", subtitle))

    for i, (_, row) in enumerate(summary.iterrows()):
        card_data = [
            [
                Paragraph(f"<b>WO {row['WorkOrderNum']}</b><br/><font color='{MUTED}'>Foreman: {row['Civil Foreman']}<br/>Inspector: {row['Inspector']}</font>", normal),
                Paragraph(f"<b>Total</b><br/><font color='{PURPLE}' size='15'><b>{row['Total Deviation']}</b></font>", normal),
                Paragraph(f"<b>Penalty</b><br/><font color='{PINK}' size='15'><b>{row['Penalty applied']}</b></font>", normal),
                Paragraph(f"<b>No Penalty</b><br/><font color='{GREEN}' size='15'><b>{row['No penalty applied']}</b></font>", normal),
                Paragraph(f"<b>Expected Waived</b><br/><font color='{ORANGE}' size='15'><b>{row['Expected Penalty Waived']}</b></font>", normal),
                Paragraph(f"<b>C / F / S</b><br/><font color='{YELLOW}' size='15'><b>{row['Civil Deviation']} / {row['Fiber Deviation']} / {row['Safety Deviation']}</b></font>", normal),
                Paragraph(f"<b>% of Total</b><br/><font color='{PURPLE}' size='15'><b>{row['% OF Total']}</b></font>", normal),
            ],
            [
                Paragraph(f"<b>Top 3 Deviations:</b> {row['Top 3 Deviations']}", normal),
                "", "", "", "", "", ""
            ]
        ]
        ctbl = Table(card_data, colWidths=[5.0*cm, 3.0*cm, 3.0*cm, 3.0*cm, 3.3*cm, 3.3*cm, 2.7*cm], rowHeights=[1.7*cm, 0.7*cm])
        ctbl.setStyle(TableStyle([
            ("SPAN", (0,1), (-1,1)),
            ("BACKGROUND", (0,0), (-1,-1), HexColor("#FFFFFF")),
            ("BOX", (0,0), (-1,-1), 0.9, HexColor("#D8C6F1")),
            ("INNERGRID", (0,0), (-1,0), 0.6, HexColor("#E4D8F5")),
            ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
            ("LEFTPADDING", (0,0), (-1,-1), 8),
            ("RIGHTPADDING", (0,0), (-1,-1), 8),
            ("BACKGROUND", (1,0), (1,0), HexColor("#F1E9FB")),
            ("BACKGROUND", (2,0), (2,0), HexColor("#FFE5EE")),
            ("BACKGROUND", (3,0), (3,0), HexColor("#EAF8F2")),
            ("BACKGROUND", (4,0), (4,0), HexColor("#FFF1E8")),
            ("BACKGROUND", (5,0), (5,0), HexColor("#FFF6D7")),
            ("BACKGROUND", (6,0), (6,0), HexColor("#F1E9FB")),
        ]))
        story.append(KeepTogether([ctbl, Spacer(1, 0.25*cm)]))

    doc.build(story)
    buffer.seek(0)
    return buffer


# ============================================================
# APP UI
# ============================================================
st.markdown(
    f"""
    <style>
    .main .block-container {{padding-top: 1.2rem; padding-bottom: 2rem; max-width: 1500px;}}
    .hero {{
        background: linear-gradient(135deg, {PURPLE}, {PURPLE_2});
        padding: 26px 30px;
        border-radius: 0 0 22px 22px;
        color: white;
        box-shadow: 0 16px 32px rgba(90,10,162,.16);
        margin-bottom: 20px;
    }}
    .hero h1 {{margin: 0; font-size: 32px; font-weight: 800;}}
    .hero p {{margin: 12px 0 0 0; font-size: 14px; opacity: .95;}}
    .kpi-card {{
        border: 1px solid {LINE};
        background: {CARD_BG};
        border-radius: 16px;
        padding: 18px 20px;
        min-height: 112px;
        box-shadow: 0 8px 22px rgba(90,10,162,.06);
    }}
    .kpi-title {{font-weight: 800; color: #3A2B52; font-size: 15px;}}
    .kpi-value {{font-size: 31px; font-weight: 900; margin-top: 8px;}}
    .kpi-subtitle {{color: {MUTED}; font-size: 12px; margin-top: 8px;}}
    .section-title {{font-size: 23px; font-weight: 900; color: {DARK}; margin-top: 16px; margin-bottom: 10px;}}
    .wo-card {{
        border: 1px solid {LINE};
        background: white;
        border-radius: 18px;
        padding: 18px;
        margin: 8px 0 16px 0;
        box-shadow: 0 8px 22px rgba(90,10,162,.06);
    }}
    .wo-header {{
        display: flex;
        justify-content: space-between;
        align-items: center;
        border-bottom: 1px solid #EEE6FA;
        padding-bottom: 10px;
        margin-bottom: 12px;
    }}
    .wo-title {{font-size: 22px; font-weight: 900; color: {PURPLE};}}
    .wo-sub {{font-size: 12px; color: {MUTED}; margin-top: 4px;}}
    .wo-percent {{background: #F1E9FB; color: {PURPLE}; font-weight: 900; padding: 8px 12px; border-radius: 999px;}}
    .wo-metrics {{display: grid; grid-template-columns: repeat(7, minmax(115px, 1fr)); gap: 10px;}}
    .mini-card {{
        background: #FAF8FE;
        border: 1px solid #E6DBF7;
        border-radius: 14px;
        padding: 12px;
        text-align: center;
    }}
    .mini-card span {{display: block; color: {MUTED}; font-size: 12px; font-weight: 700;}}
    .mini-card b {{font-size: 22px;}}
    .wo-top {{margin-top: 12px; color: {DARK}; font-size: 13px;}}
    </style>
    """,
    unsafe_allow_html=True
)

df = load_data()
summary = build_workorder_summary(df)

with st.sidebar:
    st.header("Filters")
    wo_filter = st.multiselect("Work Order", options=summary["WorkOrderNum"].tolist())
    foreman_filter = st.multiselect("Civil Foreman", options=sorted([x for x in df["Civil Foreman"].unique() if x != "—"]))
    inspector_filter = st.multiselect("Inspector", options=sorted([x for x in df["Inspector"].unique() if x != "—"]))
    designation_filter = st.multiselect("Classification", options=["Civil", "Fiber", "Safety"])

filtered = df.copy()
if wo_filter:
    filtered = filtered[filtered["WorkOrderNum"].isin(wo_filter)]
if foreman_filter:
    filtered = filtered[filtered["Civil Foreman"].isin(foreman_filter)]
if inspector_filter:
    filtered = filtered[filtered["Inspector"].isin(inspector_filter)]
if designation_filter:
    filtered = filtered[filtered["Designation"].isin(designation_filter)]

summary_f = build_workorder_summary(filtered) if not filtered.empty else build_workorder_summary(df.iloc[0:0])

st.markdown(
    """
    <div class="hero">
        <h1>STC Quality Executive Dashboard</h1>
        <p>Deviation analysis • Work Order executive performance cards • Penalty applied vs waived/not applied • Civil / Fiber / Safety classification</p>
    </div>
    """,
    unsafe_allow_html=True
)

# KPI row
total = len(filtered)
unique_wos = summary_f["WorkOrderNum"].nunique() if not summary_f.empty else 0
penalty = int(filtered["Penalty Applied Flag"].sum()) if not filtered.empty else 0
no_penalty = int(filtered["No Penalty Flag"].sum()) if not filtered.empty else 0
waived = int(filtered["Expected Penalty Waived Flag"].sum()) if not filtered.empty else 0
civil = int(filtered["Civil Flag"].sum()) if not filtered.empty else 0
fiber = int(filtered["Fiber Flag"].sum()) if not filtered.empty else 0
safety = int(filtered["Safety Flag"].sum()) if not filtered.empty else 0

cols = st.columns(6)
with cols[0]: kpi_card("Total Deviations", f"{total:,}", "All records after filters", PURPLE, "●")
with cols[1]: kpi_card("Unique WOs", f"{unique_wos:,}", "Impacted Work Orders", TEAL, "◆")
with cols[2]: kpi_card("Penalty Applied", f"{penalty:,}", f"{penalty/max(total,1):.1%} of deviations", PINK, "■")
with cols[3]: kpi_card("No Penalty", f"{no_penalty:,}", f"{no_penalty/max(total,1):.1%} of deviations", GREEN, "■")
with cols[4]: kpi_card("Expected Waived", f"{waived:,}", "Expected penalty but waived/cancelled", ORANGE, "▲")
with cols[5]: kpi_card("Civil / Fiber / Safety", f"{civil} / {fiber} / {safety}", "Deviation classification", YELLOW, "●")

st.markdown('<div class="section-title">Executive Charts</div>', unsafe_allow_html=True)
c1, c2 = st.columns([1.1, 1])
with c1:
    if not summary_f.empty:
        st.plotly_chart(fig_top_workorders(summary_f), use_container_width=True)
with c2:
    st.plotly_chart(fig_penalty_donut(penalty, no_penalty, waived), use_container_width=True)

st.markdown('<div class="section-title">Work Order Performance Cards</div>', unsafe_allow_html=True)
if not summary_f.empty:
    for _, row in summary_f.iterrows():
        wo_score_card(row)

st.markdown('<div class="section-title">Executive Work Order Summary Table</div>', unsafe_allow_html=True)
if not summary_f.empty:
    display_cols = [
        "WorkOrderNum", "Total Deviation", "Penalty applied", "No penalty applied",
        "Expected Penalty Waived", "Civil Deviation", "Civil Penalty applied",
        "Fiber Deviation", "Fiber Penalty applied", "Safety Deviation",
        "Safety Penalty applied", "% OF Total", "Civil Foreman", "Inspector"
    ]
    st.dataframe(summary_f[display_cols], use_container_width=True, hide_index=True)
else:
    st.warning("No data after current filters.")

st.markdown('<div class="section-title">Deviation Details</div>', unsafe_allow_html=True)
detail_cols = [
    "WorkOrderNum", "District", "Civil Foreman", "Inspector", "Designation",
    "Category", "SubCategory", "Short Deviation", "IsPenalty", "Expected Penalties"
]
available_detail_cols = [c for c in detail_cols if c in filtered.columns]
st.dataframe(filtered[available_detail_cols], use_container_width=True, hide_index=True)

st.markdown('<div class="section-title">Export Board-ready PDF</div>', unsafe_allow_html=True)
pdf_buffer = make_pdf(filtered, summary_f) if not filtered.empty and not summary_f.empty else None
if pdf_buffer:
    st.download_button(
        "Download Executive PDF Report",
        data=pdf_buffer,
        file_name=f"STC_Quality_Executive_Board_Report_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
        mime="application/pdf",
        use_container_width=True,
    )
else:
    st.info("No PDF available because current filters returned no data.")
