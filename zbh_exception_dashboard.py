"""
ZBH Japan · S&D Exception Dashboard
Reads zbh_sd_planning_template.xlsx (ABC_Master + Settings sheets).
Provides exception-based weekly review with DoS urgency colour coding.
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import io
import warnings
warnings.filterwarnings("ignore")

# ─────────────────────────────────────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="ZBH S&D Exception Dashboard",
    page_icon="⚠️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────────────────────────────────────
# STYLING
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Serif+Display&family=DM+Mono:wght@300;400;500&family=DM+Sans:wght@300;400;500;600&display=swap');

:root {
    --bg:      #0d1117;
    --surface: #161b22;
    --surface2:#1c2330;
    --border:  #30363d;
    --accent:  #00d4aa;
    --navy:    #1F3864;
    --red:     #f85149;
    --amber:   #f7931a;
    --yellow:  #e3b341;
    --green:   #3fb950;
    --blue:    #3b82f6;
    --text:    #e6edf3;
    --muted:   #c9d1d9;
    --dgray:   #8b949e;
}

html, body, [class*="css"] { font-family: 'DM Sans', sans-serif; color: var(--text); }
.stApp { background: var(--bg); }
#MainMenu, footer { visibility: hidden; }

section[data-testid="stSidebar"] {
    background: var(--surface) !important;
    border-right: 1px solid var(--border);
}
section[data-testid="stSidebar"] h3 {
    font-family: 'DM Mono', monospace; color: var(--accent);
    font-size: 0.72rem; letter-spacing: 0.12em; text-transform: uppercase;
    border-bottom: 1px solid var(--border); padding-bottom: 0.4rem; margin-top: 1.2rem;
}
section[data-testid="stSidebar"] label,
section[data-testid="stSidebar"] p,
section[data-testid="stSidebar"] span,
section[data-testid="stSidebar"] div[data-testid="stMarkdownContainer"] p {
    color: #e6edf3 !important;
}

.main-header {
    background: linear-gradient(135deg, #0d1117 0%, #161b22 100%);
    border: 1px solid var(--border); border-radius: 8px;
    padding: 1.25rem 2rem; margin-bottom: 1rem;
}
.main-header h1 {
    font-family: 'DM Serif Display', serif; font-size: 1.6rem;
    color: var(--text); margin: 0 0 0.2rem 0;
}
.main-header p { color: var(--muted); font-family: 'DM Mono', monospace; font-size: 0.78rem; margin: 0; }

.kpi-card {
    background: var(--surface); border: 1px solid var(--border);
    border-radius: 8px; padding: 1rem 1.25rem; text-align: center;
}
.kpi-card .val { font-family: 'DM Mono', monospace; font-size: 1.75rem; font-weight: 500; display: block; margin-bottom: 0.1rem; }
.kpi-card .lbl { font-size: 0.67rem; color: var(--muted); text-transform: uppercase; letter-spacing: 0.1em; }

.dos-critical { background: rgba(248,81,73,0.15);  color: #f85149; border-radius: 4px; padding: 2px 8px; font-family: 'DM Mono', monospace; font-size: 0.75rem; font-weight: 600; }
.dos-urgent   { background: rgba(247,147,26,0.15); color: #f7931a; border-radius: 4px; padding: 2px 8px; font-family: 'DM Mono', monospace; font-size: 0.75rem; font-weight: 600; }
.dos-monitor  { background: rgba(227,179,65,0.15); color: #e3b341; border-radius: 4px; padding: 2px 8px; font-family: 'DM Mono', monospace; font-size: 0.75rem; font-weight: 600; }
.dos-ok       { background: rgba(63,185,80,0.12);  color: #3fb950; border-radius: 4px; padding: 2px 8px; font-family: 'DM Mono', monospace; font-size: 0.75rem; font-weight: 600; }
.dos-over     { background: rgba(59,130,246,0.12); color: #3b82f6; border-radius: 4px; padding: 2px 8px; font-family: 'DM Mono', monospace; font-size: 0.75rem; font-weight: 600; }

.exc-badge-below { background: rgba(248,81,73,0.12); color:#f85149; border:1px solid rgba(248,81,73,0.25);
                   border-radius:4px; padding:2px 8px; font-family:'DM Mono',monospace; font-size:0.72rem; }
.exc-badge-over  { background: rgba(59,130,246,0.12); color:#3b82f6; border:1px solid rgba(59,130,246,0.25);
                   border-radius:4px; padding:2px 8px; font-family:'DM Mono',monospace; font-size:0.72rem; }
.exc-badge-ok    { background: rgba(63,185,80,0.10); color:#3fb950; border:1px solid rgba(63,185,80,0.2);
                   border-radius:4px; padding:2px 8px; font-family:'DM Mono',monospace; font-size:0.72rem; }

.info-box {
    background: var(--surface2); border-left: 3px solid var(--accent);
    padding: 0.6rem 1rem; border-radius: 0 4px 4px 0;
    font-size: 0.78rem; color: var(--muted); margin: 0.5rem 0;
    font-family: 'DM Mono', monospace;
}
.info-box strong { color: var(--accent); }

div[data-testid="stTabs"] button[role="tab"] {
    font-family: 'DM Mono', monospace; font-size: 0.8rem; color: var(--muted);
}
div[data-testid="stTabs"] button[role="tab"][aria-selected="true"] { color: var(--accent); }

h4, h5, h6,
div[data-testid="stMarkdownContainer"] h4,
div[data-testid="stMarkdownContainer"] h5,
div[data-testid="stMarkdownContainer"] p { color: #e6edf3 !important; }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# CONSTANTS
# ─────────────────────────────────────────────────────────────────────────────
PTLY = dict(
    paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
    font=dict(family="DM Mono, monospace", color="#e6edf3", size=11),
    margin=dict(l=45, r=20, t=40, b=40),
    legend=dict(bgcolor="rgba(22,27,34,0.85)", bordercolor="#30363d", borderwidth=1,
                font=dict(color="#e6edf3")),
)
AXIS = dict(gridcolor="#21262d", linecolor="#30363d", tickfont=dict(size=10), zeroline=False)

PL_COLORS = {
    "Hip":     "#00d4aa",
    "Knee":    "#3b82f6",
    "Shoulder":"#f7931a",
    "Trauma":  "#e05cb0",
    "CMF":     "#a78bfa",
}

ABC_COLORS = {"A": "#00d4aa", "B": "#3b82f6", "C": "#8b949e"}

DoS_COLORS = {
    "CRITICAL": "#f85149",
    "URGENT":   "#f7931a",
    "MONITOR":  "#e3b341",
    "OK":       "#3fb950",
    "OVERSTOCK":"#3b82f6",
}

# ─────────────────────────────────────────────────────────────────────────────
# DATA LOADING
# ─────────────────────────────────────────────────────────────────────────────
EXPECTED_COLS = [
    "SKU ID", "Description", "Product Line", "Component Type",
    "Size / Variant", "ABC Class", "Unit Cost (¥)", "Unit Cost ($)",
    "Current Stock (DC)", "W-4 Demand", "W-3 Demand", "W-2 Demand",
    "W-1 Demand (latest)", "Avg Weekly Demand", "Avg Daily Demand",
    "Annual Rev (¥)", "Revenue Rank", "SS Days (by class)", "SS Simple (units)",
    "SS DM (units)", "SS Active", "Days of Supply", "Inv Value (¥)",
    "Inv Value ($)", "Overstock Threshold (u)", "Exception Type",
    "Review Frequency", "In This Week's Review",
]

@st.cache_data(show_spinner=False)
def load_excel(file_bytes):
    """Load ABC_Master and Settings from the Excel workbook."""
    xl = pd.ExcelFile(io.BytesIO(file_bytes))

    # ABC_Master
    df = xl.parse("ABC_Master", header=1)
    df = df[df["SKU ID"].notna() & (df["SKU ID"] != "SKU ID")].copy()

    # Coerce numeric columns
    num_cols = [
        "Unit Cost (¥)", "Unit Cost ($)", "Current Stock (DC)",
        "W-4 Demand", "W-3 Demand", "W-2 Demand", "W-1 Demand (latest)",
        "Avg Weekly Demand", "Avg Daily Demand", "Annual Rev (¥)",
        "Revenue Rank", "SS Days (by class)", "SS Simple (units)",
        "SS DM (units)", "SS Active", "Days of Supply",
        "Inv Value (¥)", "Inv Value ($)", "Overstock Threshold (u)",
    ]
    for col in num_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    # Settings
    try:
        settings_raw = xl.parse("Settings", header=None)
        fx_rate  = float(settings_raw.iloc[3, 1])   # B4
        ss_method = str(settings_raw.iloc[6, 1])    # B7
    except Exception:
        fx_rate, ss_method = 150.0, "Simple"

    return df, fx_rate, ss_method


@st.cache_data(show_spinner=False)
def load_csv(file_bytes):
    """Load from CSV (same column structure as ABC_Master)."""
    df = pd.read_csv(io.BytesIO(file_bytes))
    num_cols = [
        "Unit Cost (¥)", "Unit Cost ($)", "Current Stock (DC)",
        "W-4 Demand", "W-3 Demand", "W-2 Demand", "W-1 Demand (latest)",
        "Avg Weekly Demand", "Avg Daily Demand", "Annual Rev (¥)",
        "Revenue Rank", "SS Days (by class)", "SS Simple (units)",
        "SS DM (units)", "SS Active", "Days of Supply",
        "Inv Value (¥)", "Inv Value ($)", "Overstock Threshold (u)",
    ]
    for col in num_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)
    return df, 150.0, "Simple"


def dos_tier(exc_type, dos):
    if exc_type == "OVERSTOCKED":   return "OVERSTOCK"
    if exc_type != "BELOW SS":      return "OK"
    if dos <= 7:                    return "CRITICAL"
    if dos <= 14:                   return "URGENT"
    return "MONITOR"


def dos_label_html(tier):
    cls = {"CRITICAL":"dos-critical","URGENT":"dos-urgent","MONITOR":"dos-monitor",
           "OK":"dos-ok","OVERSTOCK":"dos-over"}.get(tier, "dos-ok")
    labels = {"CRITICAL":"🔴 CRITICAL","URGENT":"🟡 URGENT","MONITOR":"🟠 MONITOR",
              "OK":"🟢 OK","OVERSTOCK":"🔵 OVERSTOCK"}
    return f'<span class="{cls}">{labels.get(tier, tier)}</span>'


# ─────────────────────────────────────────────────────────────────────────────
# HEADER
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="main-header">
  <h1>⚠️ ZBH Japan — S&D Exception Dashboard</h1>
  <p>DC-Level Inventory · Hiratsuka · Exception-based weekly review · Hip · Knee · Shoulder · Trauma · CMF</p>
</div>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# SIDEBAR — FILE UPLOAD
# ─────────────────────────────────────────────────────────────────────────────
DEFAULT_FILE = "zbh_sd_planning_template.xlsx"  # bundled in repo root

with st.sidebar:
    st.markdown("### 📂 Data Source")

    source_type = st.radio("Source", ["Excel workbook", "CSV file"], horizontal=True)

    uploaded = st.file_uploader(
        "Override with a different file (optional)" if source_type == "Excel workbook"
        else "Upload CSV",
        type=["xlsx"] if source_type == "Excel workbook" else ["csv"],
        key="main_upload",
    )

    st.markdown("### 🔍 Filters")
    # Filters populated after data load

# ─────────────────────────────────────────────────────────────────────────────
# LOAD DATA — bundled file auto-loads; uploader overrides if provided
# ─────────────────────────────────────────────────────────────────────────────
import os

with st.spinner("Loading data…"):
    try:
        if uploaded is not None:
            file_bytes = uploaded.read()
            if source_type == "Excel workbook":
                df_raw, fx_rate, ss_method = load_excel(file_bytes)
            else:
                df_raw, fx_rate, ss_method = load_csv(file_bytes)
        elif os.path.exists(DEFAULT_FILE):
            with open(DEFAULT_FILE, "rb") as f:
                df_raw, fx_rate, ss_method = load_excel(f.read())
        else:
            st.markdown("""
    <div class="info-box">
      <strong>No data file found.</strong> Either:<br>
      1. Add <code>zbh_sd_planning_template.xlsx</code> to the repo root, or<br>
      2. Upload a file using the sidebar uploader.
    </div>
            """, unsafe_allow_html=True)
            st.markdown("#### Expected columns (ABC_Master sheet)")
            sample = pd.DataFrame({
                "SKU ID": ["ZBH-HIP-00001", "ZBH-KNE-00001"],
                "Product Line": ["Hip", "Knee"],
                "ABC Class": ["A", "B"],
                "Current Stock (DC)": [25, 10],
                "SS Active": [20, 15],
                "Days of Supply": [12.5, 6.7],
                "Exception Type": ["OK", "BELOW SS"],
                "Inv Value (¥)": [4500000, 1800000],
            })
            st.dataframe(sample, hide_index=True, use_container_width=True)
            st.stop()
    except Exception as e:
        st.error(f"Error loading file: {e}")
        st.stop()

# Add DoS tier
# Exclude Spine (divested)
df_raw = df_raw[df_raw["Product Line"].str.strip() != "Spine"].copy()

df_raw["DoS Tier"] = df_raw.apply(
    lambda r: dos_tier(r.get("Exception Type", ""), r.get("Days of Supply", 999)), axis=1
)

# ─────────────────────────────────────────────────────────────────────────────
# SIDEBAR FILTERS
# ─────────────────────────────────────────────────────────────────────────────
with st.sidebar:
    all_pl  = sorted(df_raw["Product Line"].dropna().unique().tolist())
    all_abc = sorted(df_raw["ABC Class"].dropna().unique().tolist())
    all_exc = sorted(df_raw["Exception Type"].dropna().unique().tolist())

    sel_pl = st.multiselect("Product Line", all_pl, default=all_pl)
    sel_abc = st.multiselect("ABC Class", all_abc, default=all_abc)
    sel_exc = st.multiselect("Exception Type", all_exc,
                              default=[e for e in ["BELOW SS","OVERSTOCKED"] if e in all_exc])
    exc_only = st.toggle("Exceptions only", value=True,
                          help="Show only BELOW SS and OVERSTOCKED rows")

    st.markdown("### ⚙️ Settings")
    st.caption(f"FX rate: ¥{fx_rate:,.0f}/USD  |  SS method: {ss_method}")
    st.caption(f"Rows loaded: {len(df_raw):,}")

    st.markdown("### 🔤 ABC Class Definition")
    st.markdown("""
<div style="font-family:'DM Mono',monospace;font-size:0.72rem;line-height:1.8;">
  <span style="color:#00d4aa;font-weight:600;">A-Class</span>
  <span style="color:#c9d1d9;"> — Top 5% by annual revenue.<br>
  &nbsp;&nbsp;Highest priority. Tightest SS. Weekly review.</span><br>
  <span style="color:#3b82f6;font-weight:600;">B-Class</span>
  <span style="color:#c9d1d9;"> — Next 20% by annual revenue.<br>
  &nbsp;&nbsp;Medium priority. Standard SS. Bi-weekly review.</span><br>
  <span style="color:#8b949e;font-weight:600;">C-Class</span>
  <span style="color:#c9d1d9;"> — Remaining 75% by annual revenue.<br>
  &nbsp;&nbsp;Lower priority. Lean SS. Monthly review.</span>
</div>
""", unsafe_allow_html=True)

    st.markdown("### 📥 CSV Export")
    # Export button at bottom of sidebar

# ─────────────────────────────────────────────────────────────────────────────
# APPLY FILTERS
# ─────────────────────────────────────────────────────────────────────────────
df = df_raw.copy()
if sel_pl:  df = df[df["Product Line"].isin(sel_pl)]
if sel_abc: df = df[df["ABC Class"].isin(sel_abc)]
if sel_exc: df = df[df["Exception Type"].isin(sel_exc)]
if exc_only:
    df = df[df["Exception Type"].isin(["BELOW SS","OVERSTOCKED"])]

df_exc = df_raw[df_raw["Exception Type"].isin(["BELOW SS","OVERSTOCKED"])].copy()

# ─────────────────────────────────────────────────────────────────────────────
# KPI ROW
# ─────────────────────────────────────────────────────────────────────────────
total_exc   = len(df_exc)
below_ss    = len(df_exc[df_exc["Exception Type"] == "BELOW SS"])
overstocked = len(df_exc[df_exc["Exception Type"] == "OVERSTOCKED"])
total_val   = df_raw["Inv Value (¥)"].sum()
risk_val    = df_exc[df_exc["Exception Type"] == "BELOW SS"]["Inv Value (¥)"].sum()
exc_rate    = total_exc / max(len(df_raw), 1)

critical_ct = len(df_exc[df_exc["DoS Tier"] == "CRITICAL"])

k1, k2, k3, k4, k5, k6 = st.columns(6)
kpi_data = [
    (k1, f"{total_exc}",             "Total Exceptions",    "#f85149"),
    (k2, f"{below_ss}",             "Below Safety Stock",   "#f7931a"),
    (k3, f"{overstocked}",          "Overstocked",          "#3b82f6"),
    (k4, f"¥{total_val/1e6:.1f}M",  "Total Inv Value",      "#00d4aa"),
    (k5, f"¥{risk_val/1e6:.1f}M",   "At-Risk Value",        "#e05cb0"),
    (k6, f"{exc_rate:.1%}",         "Exception Rate",       "#c9d1d9"),
]
for col, val, lbl, color in kpi_data:
    with col:
        st.markdown(f"""
        <div class="kpi-card">
          <span class="val" style="color:{color};">{val}</span>
          <span class="lbl">{lbl}</span>
        </div>""", unsafe_allow_html=True)

if critical_ct > 0:
    st.markdown(f"""
    <div style="background:rgba(248,81,73,0.08);border:1px solid #f85149;border-radius:6px;
                padding:0.5rem 1rem;margin-top:0.75rem;font-family:'DM Mono',monospace;font-size:0.8rem;">
      🔴 <strong style="color:#f85149;">{critical_ct} CRITICAL</strong>
      SKU{"s" if critical_ct > 1 else ""} with ≤7 days of supply — immediate action required.
    </div>""", unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# TABS
# ─────────────────────────────────────────────────────────────────────────────
tab1, tab2, tab3, tab4 = st.tabs([
    "⚠️  Exception Table",
    "📊  Product Line & ABC",
    "📈  Demand Trends",
    "📋  Full Inventory",
])


# ════════════════════════════════════════════════════════════════════════════
# TAB 1 — EXCEPTION TABLE
# ════════════════════════════════════════════════════════════════════════════
with tab1:
    st.markdown("#### Exception Register")
    st.caption(
        "Sorted by urgency — CRITICAL (≤7d) first, then URGENT (8–14d), then MONITOR, then OVERSTOCKED. "
        "Use sidebar filters to narrow by product line, ABC class, or exception type."
    )

    # DoS urgency legend
    leg_cols = st.columns(5)
    for col, tier, label in zip(leg_cols,
        ["CRITICAL","URGENT","MONITOR","OVERSTOCK","OK"],
        ["≤7 days (CRITICAL)","8–14 days (URGENT)","15+ days (MONITOR)","Overstocked","OK"]):
        with col:
            color = DoS_COLORS[tier]
            st.markdown(
                f'<div style="background:rgba(22,27,34,1);border-left:3px solid {color};'
                f'padding:0.3rem 0.6rem;border-radius:0 3px 3px 0;font-family:DM Mono,monospace;'
                f'font-size:0.7rem;color:{color};">{label}</div>',
                unsafe_allow_html=True,
            )
    st.markdown("<br>", unsafe_allow_html=True)

    # Tier sort order
    tier_order = {"CRITICAL": 0, "URGENT": 1, "MONITOR": 2, "OVERSTOCK": 3, "OK": 4}
    df_sorted = df.copy()
    df_sorted["_tier_sort"] = df_sorted["DoS Tier"].map(tier_order).fillna(9)
    df_sorted = df_sorted.sort_values(["_tier_sort", "Days of Supply"]).drop(columns=["_tier_sort"])

    if df_sorted.empty:
        st.info("No rows match current filters.")
    else:
        # Display columns
        disp_cols = [
            "SKU ID", "Description", "Product Line", "ABC Class",
            "Component Type", "Size / Variant",
            "Current Stock (DC)", "SS Active", "Days of Supply",
            "Exception Type", "Inv Value (¥)", "Inv Value ($)", "DoS Tier",
        ]
        disp_cols = [c for c in disp_cols if c in df_sorted.columns]
        tbl = df_sorted[disp_cols].copy()

        # Format numbers
        for col in ["Inv Value (¥)"]:
            if col in tbl.columns:
                tbl[col] = tbl[col].map(lambda x: f"¥{x:,.0f}")
        for col in ["Inv Value ($)"]:
            if col in tbl.columns:
                tbl[col] = tbl[col].map(lambda x: f"${x:,.0f}")
        for col in ["Days of Supply"]:
            if col in tbl.columns:
                tbl[col] = tbl[col].map(lambda x: f"{x:.1f}")
        for col in ["Current Stock (DC)", "SS Active"]:
            if col in tbl.columns:
                tbl[col] = tbl[col].map(lambda x: f"{int(x):,}")

        # Colour rows by DoS tier
        def row_style(row):
            tier = row.get("DoS Tier", "OK")
            bg = {
                "CRITICAL": "background-color:rgba(248,81,73,0.10)",
                "URGENT":   "background-color:rgba(247,147,26,0.10)",
                "MONITOR":  "background-color:rgba(227,179,65,0.08)",
                "OVERSTOCK":"background-color:rgba(59,130,246,0.08)",
                "OK":       "",
            }.get(tier, "")
            return [bg] * len(row)

        styled = tbl.style.apply(row_style, axis=1)
        st.dataframe(styled, use_container_width=True, hide_index=True,
                     height=min(600, 38 + len(tbl) * 35))

        # Export
        csv_export = df_sorted[disp_cols].to_csv(index=False).encode("utf-8")
        st.download_button(
            "📥 Export to CSV",
            data=csv_export,
            file_name="zbh_exceptions_export.csv",
            mime="text/csv",
        )

        st.caption(f"Showing {len(tbl):,} rows after filters.")


# ════════════════════════════════════════════════════════════════════════════
# TAB 2 — PRODUCT LINE & ABC BREAKDOWN
# ════════════════════════════════════════════════════════════════════════════
with tab2:
    col_l, col_r = st.columns(2)

    # ── Exception count by product line ──────────────────────────────────────
    with col_l:
        st.markdown("##### Exceptions by Product Line")
        pl_exc = (df_exc.groupby(["Product Line","Exception Type"])
                  .size().reset_index(name="Count"))
        if pl_exc.empty:
            st.info("No exception data.")
        else:
            fig_pl = go.Figure()
            for exc_t, color in [("BELOW SS","#f85149"),("OVERSTOCKED","#3b82f6")]:
                sub = pl_exc[pl_exc["Exception Type"] == exc_t]
                fig_pl.add_trace(go.Bar(
                    x=sub["Product Line"], y=sub["Count"],
                    name=exc_t, marker_color=color,
                    text=sub["Count"], textposition="inside",
                ))
            fig_pl.update_layout(**PTLY, height=300, barmode="group",
                xaxis=dict(**AXIS), yaxis=dict(**AXIS, title="Exception Count"))
            fig_pl.update_layout(legend=dict(orientation="h", y=-0.25, x=0))
            st.plotly_chart(fig_pl, use_container_width=True)

    # ── Inventory value by product line ──────────────────────────────────────
    with col_r:
        st.markdown("##### Inventory Value by Product Line (¥M)")
        pl_val = (df_raw.groupby(["Product Line","ABC Class"])["Inv Value (¥)"]
                  .sum().reset_index())
        pl_val["Inv Value (¥M)"] = pl_val["Inv Value (¥)"] / 1e6
        if pl_val.empty:
            st.info("No data.")
        else:
            fig_val = go.Figure()
            for abc, color in ABC_COLORS.items():
                sub = pl_val[pl_val["ABC Class"] == abc]
                fig_val.add_trace(go.Bar(
                    x=sub["Product Line"], y=sub["Inv Value (¥M)"],
                    name=f"Class {abc}", marker_color=color,
                ))
            fig_val.update_layout(**PTLY, height=300, barmode="stack",
                xaxis=dict(**AXIS), yaxis=dict(**AXIS, title="¥M", tickprefix="¥", ticksuffix="M"))
            fig_val.update_layout(legend=dict(orientation="h", y=-0.25, x=0))
            st.plotly_chart(fig_val, use_container_width=True)

    st.markdown("---")
    col_l2, col_r2 = st.columns(2)

    # ── Exception rate by ABC class ───────────────────────────────────────────
    with col_l2:
        st.markdown("##### Exception Rate by ABC Class")
        abc_total = df_raw.groupby("ABC Class").size().reset_index(name="Total")
        abc_exc   = df_exc.groupby("ABC Class").size().reset_index(name="Exceptions")
        abc_merge = abc_total.merge(abc_exc, on="ABC Class", how="left").fillna(0)
        abc_merge["Rate"] = abc_merge["Exceptions"] / abc_merge["Total"]

        fig_abc = go.Figure(go.Bar(
            x=abc_merge["ABC Class"], y=abc_merge["Rate"],
            marker_color=[ABC_COLORS.get(c,"#8b949e") for c in abc_merge["ABC Class"]],
            text=[f"{r:.1%}" for r in abc_merge["Rate"]],
            textposition="outside",
        ))
        fig_abc.update_layout(**PTLY, height=280,
            xaxis=dict(**AXIS, title="ABC Class"),
            yaxis=dict(**AXIS, title="Exception Rate", tickformat=".0%"),
        )
        st.plotly_chart(fig_abc, use_container_width=True)

    # ── DoS urgency distribution ───────────────────────────────────────────────
    with col_r2:
        st.markdown("##### DoS Urgency Distribution (exceptions only)")
        tier_ct = df_exc.groupby("DoS Tier").size().reset_index(name="Count")
        tier_order_map = {"CRITICAL":0,"URGENT":1,"MONITOR":2,"OVERSTOCK":3}
        tier_ct["_ord"] = tier_ct["DoS Tier"].map(tier_order_map).fillna(9)
        tier_ct = tier_ct.sort_values("_ord")

        fig_tier = go.Figure(go.Bar(
            x=tier_ct["DoS Tier"], y=tier_ct["Count"],
            marker_color=[DoS_COLORS.get(t,"#8b949e") for t in tier_ct["DoS Tier"]],
            text=tier_ct["Count"], textposition="outside",
        ))
        fig_tier.update_layout(**PTLY, height=280,
            xaxis=dict(**AXIS), yaxis=dict(**AXIS, title="SKU Count"),
        )
        st.plotly_chart(fig_tier, use_container_width=True)

    # ── Exception matrix: Product Line × Type ────────────────────────────────
    st.markdown("---")
    st.markdown("##### Exception Count Matrix — Product Line × Exception Type")
    matrix = (df_exc.groupby(["Product Line","Exception Type"])
              .size().unstack(fill_value=0))
    if not matrix.empty:
        matrix["Total"] = matrix.sum(axis=1)
        matrix = matrix.sort_values("Total", ascending=False)
        st.dataframe(matrix.style.highlight_max(axis=None, color="rgba(248,81,73,0.3)"),
                     use_container_width=True)


# ════════════════════════════════════════════════════════════════════════════
# TAB 3 — DEMAND TRENDS
# ════════════════════════════════════════════════════════════════════════════
with tab3:
    st.markdown("#### 4-Week Demand Trend — Exception SKUs")
    st.caption("Focus on BELOW SS SKUs. Line shows W-4 → W-1 demand trajectory. "
               "Upward trend with low DoS = highest risk.")

    demand_cols = ["W-1 Demand (latest)","W-2 Demand","W-3 Demand","W-4 Demand"]
    week_labels = ["W-1 (latest)","W-2","W-3","W-4"]

    # Filter to below SS only for trend view
    df_trend = df_exc[df_exc["Exception Type"] == "BELOW SS"].copy()
    if "DoS Tier" in df_trend.columns:
        df_trend = df_trend.sort_values(
            df_trend["DoS Tier"].map({"CRITICAL":0,"URGENT":1,"MONITOR":2}).fillna(9).name
            if False else "Days of Supply"
        )

    if df_trend.empty:
        st.info("No BELOW SS SKUs to display.")
    else:
        # Product line selector for trend chart
        sel_trend_pl = st.selectbox(
            "Filter trend by product line",
            ["All"] + sorted(df_trend["Product Line"].dropna().unique().tolist()),
        )
        if sel_trend_pl != "All":
            df_trend = df_trend[df_trend["Product Line"] == sel_trend_pl]

        max_skus = st.slider("Max SKUs to display", 5, 50, 20, 5)
        df_trend = df_trend.head(max_skus)

        if df_trend.empty:
            st.info("No data for selected filter.")
        else:
            demand_avail = [c for c in demand_cols if c in df_trend.columns]
            if len(demand_avail) >= 2:
                fig_tr = go.Figure()
                for _, row in df_trend.iterrows():
                    tier  = row.get("DoS Tier","OK")
                    color = DoS_COLORS.get(tier,"#8b949e")
                    sku   = str(row.get("SKU ID",""))
                    pl    = str(row.get("Product Line",""))
                    vals  = [row[c] for c in demand_avail]
                    fig_tr.add_trace(go.Scatter(
                        x=week_labels[:len(demand_avail)], y=vals,
                        mode="lines+markers",
                        name=sku,
                        line=dict(color=color, width=1.5),
                        marker=dict(size=5),
                        hovertemplate=(
                            f"<b>{sku}</b><br>{pl}<br>"
                            f"DoS: {row.get('Days of Supply',0):.1f}d<br>"
                            "Week: %{x}<br>Demand: %{y}<extra></extra>"
                        ),
                    ))
                fig_tr.update_layout(**PTLY, height=420,
                    xaxis=dict(**AXIS, title="Week"),
                    yaxis=dict(**AXIS, title="Demand (units)"),
                    title=dict(text="Demand Trajectory — BELOW SS SKUs (sorted by urgency)",
                               font=dict(size=12)),
                    showlegend=False,
                )
                st.plotly_chart(fig_tr, use_container_width=True)
            else:
                st.info("Demand columns not found in data.")

        # Week-on-week change heatmap
        st.markdown("---")
        st.markdown("##### Week-on-Week Demand Change (BELOW SS SKUs)")
        if len(demand_avail) >= 2:
            wow = df_trend[["SKU ID"] + demand_avail].set_index("SKU ID")
            wow_pct = wow.pct_change(axis=1).iloc[:, 1:] * 100
            wow_pct.columns = ["W-1→W-2","W-2→W-3","W-3→W-4"][:len(wow_pct.columns)]
            if not wow_pct.empty:
                st.dataframe(
                    wow_pct.style.format("{:.1f}%")
                    .background_gradient(cmap="RdYlGn", axis=None, vmin=-50, vmax=50),
                    use_container_width=True,
                    height=min(400, 38 + len(wow_pct) * 35),
                )


# ════════════════════════════════════════════════════════════════════════════
# TAB 4 — FULL INVENTORY VIEW
# ════════════════════════════════════════════════════════════════════════════
with tab4:
    st.markdown("#### Full Inventory Register")
    st.caption("All SKUs including non-exception rows. Use sidebar filters to narrow.")

    all_disp = [
        "SKU ID","Description","Product Line","ABC Class","Component Type","Size / Variant",
        "Current Stock (DC)","SS Active","Days of Supply","Exception Type",
        "Inv Value (¥)","Inv Value ($)","Avg Daily Demand","Review Frequency",
        "In This Week's Review","DoS Tier",
    ]
    all_disp = [c for c in all_disp if c in df_raw.columns]
    full_tbl = df_raw[all_disp].copy()

    # Apply product line filter only (show all exceptions)
    if sel_pl:
        full_tbl = full_tbl[full_tbl["Product Line"].isin(sel_pl)]

    for col in ["Inv Value (¥)"]:
        if col in full_tbl.columns:
            full_tbl[col] = full_tbl[col].map(lambda x: f"¥{x:,.0f}")
    for col in ["Inv Value ($)"]:
        if col in full_tbl.columns:
            full_tbl[col] = full_tbl[col].map(lambda x: f"${x:,.0f}")

    st.dataframe(full_tbl, use_container_width=True, hide_index=True,
                 height=600)

    # Full CSV export
    full_csv = df_raw.to_csv(index=False).encode("utf-8")
    st.download_button(
        "📥 Export full inventory to CSV",
        data=full_csv,
        file_name="zbh_full_inventory_export.csv",
        mime="text/csv",
    )
    st.caption(f"{len(full_tbl):,} SKUs displayed.")


# ─────────────────────────────────────────────────────────────────────────────
# FOOTER
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("""
<div style="text-align:center;color:#8b949e;font-family:'DM Mono',monospace;
font-size:0.68rem;margin-top:2.5rem;padding:1rem;border-top:1px solid #21262d;">
ZBH Japan S&D Exception Dashboard · DC-Level Inventory · Hiratsuka · Exception-based weekly review
</div>
""", unsafe_allow_html=True)
