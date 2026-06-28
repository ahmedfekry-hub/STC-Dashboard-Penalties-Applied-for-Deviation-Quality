
import os
from io import BytesIO
import pandas as pd
import numpy as np
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go

# PDF export dependencies
import matplotlib.pyplot as plt
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image, PageBreak
)

st.set_page_config(
    page_title="STC Deviation Quality Dashboard",
    layout="wide",
    initial_sidebar_state="expanded"
)

PURPLE = "#5A0AA2"
DARK_PURPLE = "#2F184B"
TEAL = "#24C6D4"
ORANGE = "#FF7A45"
PINK = "#EF476F"
GREEN = "#10B981"
GRAY = "#6B7280"
SOFT_BG = "#F6F4F9"

st.markdown(f"""
<style>
    .stApp {{background-color: {SOFT_BG};}}
    .block-container {{padding-top: 1.1rem; padding-bottom: 1rem;}}
    .hero {{
        background: linear-gradient(90deg, {PURPLE}, {DARK_PURPLE});
        color: white;
        padding: 18px 22px;
        border-radius: 18px;
        box-shadow: 0 5px 20px rgba(90,10,162,0.18);
    }}
    .hero h1 {{font-size: 30px; margin: 0; font-weight: 800;}}
    .hero p {{font-size: 15px; margin: 6px 0 0 0; opacity: .92;}}
    div[data-testid="stMetric"] {{
        background: white;
        border: 1px solid #E7E2F0;
        padding: 14px 16px;
        border-radius: 16px;
        box-shadow: 0 3px 14px rgba(47,33,68,0.07);
    }}
    div[data-testid="stMetric"] label {{color: #6F6A7A !important; font-size: 13px !important;}}
    div[data-testid="stMetric"] [data-testid="stMetricValue"] {{color: {DARK_PURPLE}; font-weight: 800;}}
    .section-title {{
        font-size: 22px;
        color: {DARK_PURPLE};
        font-weight: 800;
        margin: 12px 0 2px 0;
    }}
    .note-box {{
        background: white;
        border-left: 7px solid {PURPLE};
        padding: 14px 18px;
        border-radius: 14px;
        box-shadow: 0 3px 14px rgba(47,33,68,0.06);
    }}
</style>
""", unsafe_allow_html=True)

@st.cache_data
def load_data():
    base = os.path.dirname(__file__)
    xlsx_path = os.path.join(base, "Deviation.xlsx")
    csv_path = os.path.join(base, "deviation_clean_data.csv")
    if os.path.exists(xlsx_path):
        df = pd.read_excel(xlsx_path)
    else:
        df = pd.read_csv(csv_path)

    df["District"] = df["District"].astype(str).str.strip().str.upper()
    df["WorkOrderNum"] = df["WorkOrderNum"].astype(str).str.strip()
    for col in ["DeviationName", "Category", "DeviationStatus", "SubCategory", "Designation"]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()
    df["IsPenalty"] = df["IsPenalty"].astype(str).str.strip().str.upper().map(lambda x: "Y" if x == "Y" else "N")
    df["ServiceAffecting"] = df["ServiceAffecting"].astype(str).str.strip().str.upper().map(lambda x: "YES" if x == "YES" else "NO")
    df["DateOfDeviation"] = pd.to_datetime(df["DateOfDeviation"], errors="coerce")
    df["Month"] = df["DateOfDeviation"].dt.to_period("M").astype(str)

    short_map = {
        "WORKER NOT HOLDING STC ID BADGE": "Worker missing STC ID badge",
        "NO DEBRIS IS REMAINING ON THE SITE?": "No debris remaining on site",
        "HAVE ALL STC SAFETY MEASUREMENTS BEEN FOLLOWED?": "STC safety measures not followed",
        "UNREGISTERED WORKER": "Unregistered worker",
        "UNCERTIFIED WORKER": "Uncertified worker",
        "PEDESTRIAN PASSES PLACED EVERY 100M IN FRONT OF CUSTOMER HOUSES AS REQUIRED": "Pedestrian passes not placed every 100m",
        "ALL DAMAGE TO PROPERTY (TILES, CURBS, WALLS, ASPHALT) HAS BEEN REPAIRED TO ORIGINAL STATE OR AS PER CUSTOMER REQUEST AT WORK COMPLETION ?": "Property damage not restored",
        "ALL INSTALLED PROTECTIVE U-GUARDS MEET STC SPECIFICATIONS?": "U-guards not meeting STC specs",
        "IS THE TRENCH ROUTE STRAIGHT?": "Trench route not straight",
        "ALL REQUIRED WORKER UNIFORM AND SAFETY ITEMS ARE PRESENT?": "Worker PPE/uniform items missing",
        "MUNICIPALITY PERMITS ARE VALID AND RENEWED AS REQUIRED?": "Municipality permits not valid/renewed",
        "NUMBER OF SIGNBOARDS AVAILABLE IN THE SITE ARE SUFFICIENT?": "Insufficient site signboards",
        "ALL HANDHOLES / MANHOLES ARE INSTALLED AS PER STC REQUIREMENTS?": "HH/MH not as per STC requirements",
        "MUNCIPALITY PERMIT IS AVIALABLE ON THE SIGNBOARD?": "Municipality permit not on signboard",
        "HAS THE TRENCH BEEN CLOSED WITHIN 24 HOURS OF OPENING?": "Trench not closed within 24 hours",
    }
    df["DeviationShort"] = df["DeviationName"].map(short_map).fillna(df["DeviationName"].astype(str).str.slice(0, 85))
    return df

df = load_data()

def filter_df(data):
    with st.sidebar:
        st.markdown("## Executive Filters")
        districts = st.multiselect("District", sorted(data["District"].dropna().unique()))
        wos = st.multiselect("Work Order", sorted(data["WorkOrderNum"].dropna().unique()))
        cats = st.multiselect("Category", sorted(data["Category"].dropna().unique()))
        subs = st.multiselect("Sub Category", sorted(data["SubCategory"].dropna().unique()))
        status = st.multiselect("Deviation Status", sorted(data["DeviationStatus"].dropna().unique()))
        penalty = st.multiselect("Penalty", ["Y", "N"])
        service = st.multiselect("Service Affecting", ["YES", "NO"])
        date_min = data["DateOfDeviation"].min()
        date_max = data["DateOfDeviation"].max()
        if pd.notna(date_min) and pd.notna(date_max):
            date_range = st.date_input("Deviation Date", value=(date_min.date(), date_max.date()))
        else:
            date_range = None
        top_n = st.slider("Top N display", min_value=5, max_value=25, value=12, step=1)

    f = data.copy()
    if districts:
        f = f[f["District"].isin(districts)]
    if wos:
        f = f[f["WorkOrderNum"].isin(wos)]
    if cats:
        f = f[f["Category"].isin(cats)]
    if subs:
        f = f[f["SubCategory"].isin(subs)]
    if status:
        f = f[f["DeviationStatus"].isin(status)]
    if penalty:
        f = f[f["IsPenalty"].isin(penalty)]
    if service:
        f = f[f["ServiceAffecting"].isin(service)]
    if date_range and len(date_range) == 2:
        start, end = pd.to_datetime(date_range[0]), pd.to_datetime(date_range[1])
        f = f[(f["DateOfDeviation"] >= start) & (f["DateOfDeviation"] <= end + pd.Timedelta(days=1))]
    return f, top_n

def wo_summary(data):
    if data.empty:
        return pd.DataFrame()
    out = data.groupby(["WorkOrderNum", "District"]).agg(
        Total_Deviations=("DeviationName", "size"),
        Penalty_Count=("IsPenalty", lambda s: int((s == "Y").sum())),
        Service_Affecting_Count=("ServiceAffecting", lambda s: int((s == "YES").sum())),
        Deviation_Types=("DeviationShort", "nunique"),
        Categories=("Category", "nunique"),
        First_Deviation_Date=("DateOfDeviation", "min"),
        Last_Deviation_Date=("DateOfDeviation", "max"),
    ).reset_index()
    out["Share_%"] = (out["Total_Deviations"] / max(len(data), 1) * 100).round(2)
    out["Penalty_%"] = (out["Penalty_Count"] / out["Total_Deviations"] * 100).round(2)
    out["Service_Affecting_%"] = (out["Service_Affecting_Count"] / out["Total_Deviations"] * 100).round(2)
    return out.sort_values(["Total_Deviations", "Penalty_Count", "Service_Affecting_Count"], ascending=False)

def _matplotlib_img(fig):
    bio = BytesIO()
    fig.savefig(bio, format="png", dpi=180, bbox_inches="tight")
    plt.close(fig)
    bio.seek(0)
    return bio

def chart_top_wo(summary, top_n=12):
    fig, ax = plt.subplots(figsize=(10, 4.6))
    if summary.empty:
        ax.text(0.5, 0.5, "No data", ha="center", va="center")
    else:
        d = summary.head(top_n).sort_values("Total_Deviations")
        ax.barh(d["WorkOrderNum"], d["Total_Deviations"], color=PURPLE)
        for y, (cnt, sh) in enumerate(zip(d["Total_Deviations"], d["Share_%"])):
            ax.text(cnt + max(d["Total_Deviations"]) * 0.01, y, f"{cnt} | {sh:.1f}%", va="center", fontsize=9)
        ax.set_title("Top Work Orders by Deviation Count and Share %", fontsize=13, fontweight="bold")
        ax.set_xlabel("Deviation Count")
        ax.spines[['top','right','left','bottom']].set_visible(False)
        ax.grid(axis="x", color="#E5E7EB")
    return _matplotlib_img(fig)

def chart_top_deviation(data, top_n=12):
    fig, ax = plt.subplots(figsize=(10, 4.6))
    d = data["DeviationShort"].value_counts().head(top_n).sort_values()
    if d.empty:
        ax.text(0.5, 0.5, "No data", ha="center", va="center")
    else:
        ax.barh(d.index, d.values, color=PURPLE)
        for y, v in enumerate(d.values):
            ax.text(v + max(d.values) * 0.01, y, str(int(v)), va="center", fontsize=9)
        ax.set_title("Most Common Deviation Types", fontsize=13, fontweight="bold")
        ax.set_xlabel("Deviation Count")
        ax.spines[['top','right','left','bottom']].set_visible(False)
        ax.grid(axis="x", color="#E5E7EB")
    return _matplotlib_img(fig)

def chart_heatmap(data):
    fig, ax = plt.subplots(figsize=(10.5, 4.8))
    heat = pd.crosstab(data["District"], data["DeviationShort"])
    if heat.empty:
        ax.text(0.5, 0.5, "No data", ha="center", va="center")
    else:
        im = ax.imshow(heat.values, cmap="Purples", aspect="auto")
        ax.set_xticks(range(len(heat.columns)))
        ax.set_xticklabels([c[:22] for c in heat.columns], rotation=35, ha="right", fontsize=8)
        ax.set_yticks(range(len(heat.index)))
        ax.set_yticklabels(heat.index, fontsize=9)
        for i in range(heat.shape[0]):
            for j in range(heat.shape[1]):
                val = int(heat.iloc[i, j])
                ax.text(j, i, str(val), ha="center", va="center", color="white" if val > heat.values.max()*0.45 else "#333333", fontsize=8)
        ax.set_title("District Heatmap by Deviation Nature", fontsize=13, fontweight="bold")
        for spine in ax.spines.values():
            spine.set_visible(False)
    return _matplotlib_img(fig)

def build_pdf_report(data, summary, top_n=12):
    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=landscape(A4),
        rightMargin=0.8*cm,
        leftMargin=0.8*cm,
        topMargin=0.8*cm,
        bottomMargin=0.8*cm,
    )
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle("TitlePurple", parent=styles["Title"], fontSize=22, textColor=colors.HexColor(PURPLE), spaceAfter=8)
    h_style = ParagraphStyle("Header", parent=styles["Heading2"], fontSize=14, textColor=colors.HexColor(DARK_PURPLE), spaceBefore=8, spaceAfter=6)
    small = ParagraphStyle("Small", parent=styles["BodyText"], fontSize=8, leading=10)
    normal = ParagraphStyle("NormalCustom", parent=styles["BodyText"], fontSize=10, leading=12)

    story = []
    story.append(Paragraph("STC Deviation Quality Director Dashboard", title_style))
    story.append(Paragraph("Work Order ranking • deviation percentages • penalty/service impact • violation nature drilldown", normal))
    story.append(Spacer(1, 0.25*cm))

    total = len(data)
    unique_wo = data["WorkOrderNum"].nunique()
    penalty_count = int((data["IsPenalty"] == "Y").sum()) if total else 0
    service_count = int((data["ServiceAffecting"] == "YES").sum()) if total else 0
    top_wo_count = int(summary.iloc[0]["Total_Deviations"]) if not summary.empty else 0
    top_wo_share = float(summary.iloc[0]["Share_%"]) if not summary.empty else 0
    top_wo_name = str(summary.iloc[0]["WorkOrderNum"]) if not summary.empty else "-"
    metrics = [
        ["Total Deviations", f"{total:,}", "Work Orders", f"{unique_wo:,}", "Highest WO", f"{top_wo_name}"],
        ["Penalty Deviations", f"{penalty_count:,} ({penalty_count/total*100:.1f}%)" if total else "0", "Service Affecting", f"{service_count:,} ({service_count/total*100:.1f}%)" if total else "0", "Highest WO Share", f"{top_wo_count:,} / {top_wo_share:.1f}%"],
    ]
    mt = Table(metrics, colWidths=[3.1*cm,2.7*cm,3.1*cm,2.7*cm,3.1*cm,3.5*cm])
    mt.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,-1),colors.HexColor("F3E8FF")),
        ("TEXTCOLOR",(0,0),(-1,-1),colors.HexColor(DARK_PURPLE)),
        ("FONTNAME",(0,0),(-1,-1),"Helvetica-Bold"),
        ("GRID",(0,0),(-1,-1),0.5,colors.HexColor("D8B4FE")),
        ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
        ("LEFTPADDING",(0,0),(-1,-1),8),
        ("RIGHTPADDING",(0,0),(-1,-1),8),
    ]))
    story.append(mt)
    story.append(Spacer(1, 0.35*cm))

    # Charts page 1
    chart_table = Table([
        [Image(chart_top_wo(summary, top_n), width=12.3*cm, height=6.0*cm),
         Image(chart_top_deviation(data, top_n), width=12.3*cm, height=6.0*cm)]
    ], colWidths=[13*cm,13*cm])
    story.append(chart_table)

    story.append(PageBreak())
    story.append(Paragraph("WO Ranking and Percentages", h_style))
    if not summary.empty:
        table_df = summary.head(20).copy()
        table_df["First_Deviation_Date"] = table_df["First_Deviation_Date"].dt.strftime("%Y-%m-%d")
        table_df["Last_Deviation_Date"] = table_df["Last_Deviation_Date"].dt.strftime("%Y-%m-%d")
        cols = ["WorkOrderNum","District","Total_Deviations","Share_%","Penalty_Count","Penalty_%","Service_Affecting_Count","Service_Affecting_%","Deviation_Types","First_Deviation_Date","Last_Deviation_Date"]
        display_names = ["WO","District","Dev.","Share %","Penalty","Penalty %","Service","Service %","Types","First","Last"]
        body = [display_names] + table_df[cols].astype(str).values.tolist()
        t = Table(body, repeatRows=1, colWidths=[2.5*cm,1.8*cm,1.4*cm,1.6*cm,1.5*cm,1.7*cm,1.5*cm,1.7*cm,1.4*cm,1.9*cm,1.9*cm])
        t.setStyle(TableStyle([
            ("BACKGROUND",(0,0),(-1,0),colors.HexColor(PURPLE)),
            ("TEXTCOLOR",(0,0),(-1,0),colors.white),
            ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
            ("FONTSIZE",(0,0),(-1,-1),7),
            ("GRID",(0,0),(-1,-1),0.25,colors.HexColor("E7E2F0")),
            ("ROWBACKGROUNDS",(0,1),(-1,-1),[colors.white, colors.HexColor("F8F5FC")]),
            ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
        ]))
        story.append(t)
    else:
        story.append(Paragraph("No WO data available for current filters.", normal))

    story.append(PageBreak())
    story.append(Paragraph("District Heatmap and Deviation Nature", h_style))
    story.append(Image(chart_heatmap(data), width=25.5*cm, height=8.0*cm))
    story.append(Spacer(1, 0.3*cm))

    top_dev = data["DeviationShort"].value_counts().head(15).rename_axis("Deviation").reset_index(name="Count")
    top_dev["Share %"] = (top_dev["Count"]/max(len(data),1)*100).round(1)
    body = [["Deviation Nature","Count","Share %"]] + top_dev.astype(str).values.tolist()
    t2 = Table(body, repeatRows=1, colWidths=[17*cm,3*cm,3*cm])
    t2.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,0),colors.HexColor(PURPLE)),
        ("TEXTCOLOR",(0,0),(-1,0),colors.white),
        ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
        ("FONTSIZE",(0,0),(-1,-1),8),
        ("GRID",(0,0),(-1,-1),0.25,colors.HexColor("E7E2F0")),
        ("ROWBACKGROUNDS",(0,1),(-1,-1),[colors.white, colors.HexColor("F8F5FC")]),
    ]))
    story.append(t2)

    story.append(PageBreak())
    story.append(Paragraph("Recommended Management Action Tracker", h_style))
    actions = [
        ["Issue", "Focus Area", "Action", "Timeline", "Owner"],
        ["High-repeat WOs", "Top WO ranking", "Daily review for WOs with highest deviation count/share; no closure without QA signoff", "Immediate", "Project Managers + QA/QC"],
        ["Penalty deviations", "Penalty = Y records", "Separate penalty list by WO; validate root cause and commercial exposure", "0-7 days", "Quality + PMO"],
        ["Service-affecting items", "Service Affecting = YES", "Escalate to site manager same day; evidence-based closure before next activity", "0-7 days", "Site Supervisors"],
        ["Repeated deviation nature", "Top 5 deviation types", "Toolbox talk and visual checklist for repeated categories", "0-14 days", "QA/QC + HSE"],
        ["District concentration", "District heatmap", "Assign district owner and publish weekly district ranking", "Weekly", "Area Managers"],
    ]
    at = Table(actions, repeatRows=1, colWidths=[4*cm,4.5*cm,10.5*cm,3*cm,4*cm])
    at.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,0),colors.HexColor(PURPLE)),
        ("TEXTCOLOR",(0,0),(-1,0),colors.white),
        ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
        ("FONTSIZE",(0,0),(-1,-1),8),
        ("GRID",(0,0),(-1,-1),0.25,colors.HexColor("E7E2F0")),
        ("ROWBACKGROUNDS",(0,1),(-1,-1),[colors.white, colors.HexColor("F8F5FC")]),
        ("VALIGN",(0,0),(-1,-1),"TOP"),
    ]))
    story.append(at)

    story.append(Spacer(1, 0.4*cm))
    story.append(Paragraph("Source: user-provided Deviation.xlsx. This report intentionally excludes Risk Score and ranks WOs by actual deviation count/share, penalty %, and service-affecting %.", small))

    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()

filtered, top_n = filter_df(df)
summary = wo_summary(filtered)

st.markdown("""
<div class="hero">
    <h1>STC Deviation Quality Director Dashboard</h1>
    <p>Work Order ranking • deviation percentages • penalty and service impact • violation nature drilldown • PDF executive export</p>
</div>
""", unsafe_allow_html=True)

total = len(filtered)
unique_wo = filtered["WorkOrderNum"].nunique()
penalty_count = int((filtered["IsPenalty"] == "Y").sum()) if total else 0
service_count = int((filtered["ServiceAffecting"] == "YES").sum()) if total else 0
top_wo_count = int(summary.iloc[0]["Total_Deviations"]) if not summary.empty else 0
top_wo_share = float(summary.iloc[0]["Share_%"]) if not summary.empty else 0
top_wo_name = str(summary.iloc[0]["WorkOrderNum"]) if not summary.empty else "-"
penalty_rate = penalty_count / total * 100 if total else 0
service_rate = service_count / total * 100 if total else 0

st.markdown("")
k1, k2, k3, k4, k5 = st.columns(5)
k1.metric("Total Deviations", f"{total:,}")
k2.metric("Work Orders", f"{unique_wo:,}")
k3.metric("Penalty Deviations", f"{penalty_count:,}", f"{penalty_rate:.1f}%")
k4.metric("Service Affecting", f"{service_count:,}", f"{service_rate:.1f}%")
k5.metric("Highest WO", top_wo_name, f"{top_wo_count:,} dev. / {top_wo_share:.1f}%")

st.download_button(
    "📄 Export Executive PDF Report",
    data=build_pdf_report(filtered, summary, top_n),
    file_name="STC_Deviation_Executive_Dashboard_Report.pdf",
    mime="application/pdf",
    help="Exports a management-style PDF report similar to the attached example, based on the current filters."
)

tabs = st.tabs([
    "Executive Overview",
    "WO Ranking & Percentages",
    "WO Drilldown",
    "Deviation Nature",
    "District Heatmap",
    "Penalty & Service Impact",
    "Raw Data"
])

with tabs[0]:
    st.markdown('<div class="section-title">Executive View</div>', unsafe_allow_html=True)
    c1, c2 = st.columns([1.1, 1])
    with c1:
        top_wo = summary.head(top_n).sort_values("Total_Deviations")
        fig = px.bar(
            top_wo,
            x="Total_Deviations",
            y="WorkOrderNum",
            orientation="h",
            color="District",
            text="Share_%",
            color_discrete_sequence=[PURPLE, TEAL, ORANGE, PINK, GREEN],
            hover_data=["Penalty_Count", "Penalty_%", "Service_Affecting_Count", "Service_Affecting_%", "Deviation_Types"]
        )
        fig.update_traces(texttemplate="%{text:.1f}%", textposition="outside")
        fig.update_layout(title="Top Work Orders by Deviation Count and Share %", height=430, margin=dict(l=10, r=10, t=50, b=10), xaxis_title="Deviation Count", yaxis_title="")
        st.plotly_chart(fig, use_container_width=True)

    with c2:
        top_dev = filtered["DeviationShort"].value_counts().head(top_n).reset_index()
        top_dev.columns = ["Deviation", "Count"]
        fig2 = px.bar(
            top_dev.sort_values("Count"),
            x="Count",
            y="Deviation",
            orientation="h",
            color_discrete_sequence=[PURPLE],
            text="Count"
        )
        fig2.update_layout(title="Most Common Deviation Types", height=430, margin=dict(l=10, r=10, t=50, b=10), yaxis_title="")
        st.plotly_chart(fig2, use_container_width=True)

    if not summary.empty:
        top3_share = summary.head(3)["Total_Deviations"].sum() / total * 100 if total else 0
        top5_share = summary.head(5)["Total_Deviations"].sum() / total * 100 if total else 0
        top_category = filtered["Category"].value_counts().idxmax() if total else "-"
        top_category_count = filtered["Category"].value_counts().max() if total else 0
        st.markdown(f"""
        <div class="note-box">
        <b>Quality Director Escalation Message:</b><br>
        The deviation pattern is concentrated in a limited number of work orders. The top 3 WOs represent <b>{top3_share:.1f}%</b> of all filtered deviations and the top 5 WOs represent <b>{top5_share:.1f}%</b>. 
        The highest repeated nature is linked to <b>{top_category}</b> with <b>{top_category_count:,}</b> records. Immediate WO-level accountability is required before the issue becomes a KPI drag on acceptance, certification, and H&S closure.
        </div>
        """, unsafe_allow_html=True)

with tabs[1]:
    st.markdown('<div class="section-title">WO Ranking, Share %, Penalty %, and Service-Affecting %</div>', unsafe_allow_html=True)
    if summary.empty:
        st.warning("No data available for selected filters.")
    else:
        fig = px.scatter(
            summary,
            x="Share_%",
            y="Penalty_%",
            size="Total_Deviations",
            color="District",
            hover_name="WorkOrderNum",
            hover_data=["District", "Total_Deviations", "Penalty_Count", "Penalty_%", "Service_Affecting_Count", "Service_Affecting_%"],
            color_discrete_sequence=[PURPLE, TEAL, ORANGE, PINK, GREEN]
        )
        fig.update_layout(title="WO Positioning: Share % vs Penalty %", height=480, xaxis_title="WO Share of Deviations (%)", yaxis_title="Penalty Rate (%)")
        st.plotly_chart(fig, use_container_width=True)
        display = summary.copy()
        display["First_Deviation_Date"] = display["First_Deviation_Date"].dt.strftime("%Y-%m-%d")
        display["Last_Deviation_Date"] = display["Last_Deviation_Date"].dt.strftime("%Y-%m-%d")
        st.dataframe(display, use_container_width=True, hide_index=True)

with tabs[2]:
    st.markdown('<div class="section-title">WO Drilldown</div>', unsafe_allow_html=True)
    if summary.empty:
        st.warning("No data available.")
    else:
        selected_wo = st.selectbox("Select Work Order", summary["WorkOrderNum"].tolist())
        wdf = filtered[filtered["WorkOrderNum"] == selected_wo]
        wsum = summary[summary["WorkOrderNum"] == selected_wo].iloc[0]
        a, b, c, d, e = st.columns(5)
        a.metric("WO Deviations", f"{int(wsum['Total_Deviations']):,}")
        b.metric("Share %", f"{wsum['Share_%']:.1f}%")
        c.metric("Penalty Count", f"{int(wsum['Penalty_Count']):,}", f"{wsum['Penalty_%']:.1f}%")
        d.metric("Service Affecting", f"{int(wsum['Service_Affecting_Count']):,}", f"{wsum['Service_Affecting_%']:.1f}%")
        e.metric("Deviation Types", f"{int(wsum['Deviation_Types']):,}")

        col1, col2 = st.columns([1.15, 1])
        with col1:
            d1 = wdf["DeviationShort"].value_counts().head(15).reset_index()
            d1.columns = ["Deviation", "Count"]
            fig = px.bar(d1.sort_values("Count"), x="Count", y="Deviation", orientation="h", color_discrete_sequence=[PURPLE], text="Count")
            fig.update_layout(title=f"Deviation Nature for WO {selected_wo}", height=440, margin=dict(l=10, r=10, t=50, b=10), yaxis_title="")
            st.plotly_chart(fig, use_container_width=True)
        with col2:
            d2 = wdf["Category"].value_counts().reset_index()
            d2.columns = ["Category", "Count"]
            fig2 = px.pie(d2, names="Category", values="Count", hole=.45, color_discrete_sequence=[PURPLE, TEAL, ORANGE, PINK, GREEN])
            fig2.update_layout(title="Category Mix", height=440)
            st.plotly_chart(fig2, use_container_width=True)

        timeline = wdf.groupby(wdf["DateOfDeviation"].dt.date).size().reset_index(name="Count")
        timeline.columns = ["Date", "Count"]
        fig3 = px.line(timeline, x="Date", y="Count", markers=True, color_discrete_sequence=[PURPLE])
        fig3.update_layout(title="Deviation Timeline for Selected WO", height=320, yaxis_title="Deviation Count")
        st.plotly_chart(fig3, use_container_width=True)

        st.markdown("#### Selected WO Raw Records")
        st.dataframe(wdf.sort_values("DateOfDeviation", ascending=False), use_container_width=True, hide_index=True)

with tabs[3]:
    st.markdown('<div class="section-title">Deviation Nature Analysis</div>', unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        cat = filtered["Category"].value_counts().reset_index()
        cat.columns = ["Category", "Count"]
        fig = px.bar(cat.sort_values("Count"), x="Count", y="Category", orientation="h", color_discrete_sequence=[TEAL], text="Count")
        fig.update_layout(title="Deviation Count by Category", height=420, yaxis_title="")
        st.plotly_chart(fig, use_container_width=True)
    with c2:
        sub = filtered["SubCategory"].value_counts().head(top_n).reset_index()
        sub.columns = ["SubCategory", "Count"]
        fig2 = px.bar(sub.sort_values("Count"), x="Count", y="SubCategory", orientation="h", color_discrete_sequence=[ORANGE], text="Count")
        fig2.update_layout(title="Deviation Count by SubCategory", height=420, yaxis_title="")
        st.plotly_chart(fig2, use_container_width=True)

    matrix = pd.crosstab(filtered["DeviationShort"], filtered["WorkOrderNum"])
    if not matrix.empty:
        fig3 = px.imshow(matrix, text_auto=True, aspect="auto", color_continuous_scale=[[0, "#F3E8FF"], [1, PURPLE]])
        fig3.update_layout(title="Deviation Type x Work Order Matrix", height=650, xaxis_title="Work Order", yaxis_title="Deviation Type", coloraxis_showscale=False)
        st.plotly_chart(fig3, use_container_width=True)

with tabs[4]:
    st.markdown('<div class="section-title">District Heatmap</div>', unsafe_allow_html=True)
    heat = pd.crosstab(filtered["District"], filtered["DeviationShort"])
    if heat.empty:
        st.warning("No data to display.")
    else:
        fig = px.imshow(heat, text_auto=True, aspect="auto", color_continuous_scale=[[0, "#F3E8FF"], [1, PURPLE]])
        fig.update_layout(title="Deviation Type Concentration by District", height=600, coloraxis_showscale=False)
        st.plotly_chart(fig, use_container_width=True)

    district_summary = filtered.groupby("District").agg(
        Total_Deviations=("DeviationName", "size"),
        Work_Orders=("WorkOrderNum", "nunique"),
        Penalty_Count=("IsPenalty", lambda s: int((s == "Y").sum())),
        Service_Affecting=("ServiceAffecting", lambda s: int((s == "YES").sum()))
    ).reset_index()
    if not district_summary.empty:
        district_summary["Share_%"] = (district_summary["Total_Deviations"]/len(filtered)*100).round(2)
        district_summary["Penalty_%"] = (district_summary["Penalty_Count"]/district_summary["Total_Deviations"]*100).round(2)
        st.dataframe(district_summary.sort_values("Total_Deviations", ascending=False), use_container_width=True, hide_index=True)

with tabs[5]:
    st.markdown('<div class="section-title">Penalty & Service Impact</div>', unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        penalty_by_wo = summary.sort_values("Penalty_Count", ascending=False).head(top_n)
        fig = px.bar(penalty_by_wo.sort_values("Penalty_Count"), x="Penalty_Count", y="WorkOrderNum", orientation="h", color="District", color_discrete_sequence=[PINK, ORANGE, PURPLE])
        fig.update_layout(title="Penalty Deviations by WO", height=430, yaxis_title="")
        st.plotly_chart(fig, use_container_width=True)
    with c2:
        service_by_wo = summary.sort_values("Service_Affecting_Count", ascending=False).head(top_n)
        fig2 = px.bar(service_by_wo.sort_values("Service_Affecting_Count"), x="Service_Affecting_Count", y="WorkOrderNum", orientation="h", color="District", color_discrete_sequence=[PINK, ORANGE, PURPLE])
        fig2.update_layout(title="Service-Affecting Deviations by WO", height=430, yaxis_title="")
        st.plotly_chart(fig2, use_container_width=True)

    impact_table = summary[["WorkOrderNum", "District", "Total_Deviations", "Share_%", "Penalty_Count", "Penalty_%", "Service_Affecting_Count", "Service_Affecting_%", "Deviation_Types"]]
    st.dataframe(impact_table, use_container_width=True, hide_index=True)

with tabs[6]:
    st.markdown('<div class="section-title">Raw Deviation Records</div>', unsafe_allow_html=True)
    st.download_button(
        "Download Filtered Data CSV",
        data=filtered.to_csv(index=False).encode("utf-8-sig"),
        file_name="filtered_deviation_records.csv",
        mime="text/csv"
    )
    st.dataframe(filtered.sort_values("DateOfDeviation", ascending=False), use_container_width=True, hide_index=True)

st.caption("Dashboard generated from the uploaded Deviation.xlsx file. WO priority is based on actual deviation count/share, penalty %, and service-affecting %, without Risk Score.")
