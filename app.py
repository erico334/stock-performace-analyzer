"""
app.py - Stock Performance Analyzer
Upload your stock Excel file and get a full analysis report.
"""
import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import date
import analyzer
import excel_builder

st.set_page_config(
    page_title="Stock Performance Analyzer",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
  [data-testid="stAppViewContainer"] { background:#f4f6f9; }
  [data-testid="stSidebar"]          { background:#1F3864; }
  [data-testid="stSidebar"] * { color:#ffffff !important; }
  .metric-card {
    background:#ffffff; border-radius:12px; padding:1.1rem 1.3rem;
    border:1px solid #e0e6f0; box-shadow:0 1px 4px rgba(0,0,0,.06);
    text-align:center;
  }
  .metric-label { font-size:12px; color:#888; font-weight:600;
                  text-transform:uppercase; letter-spacing:.05em; margin-bottom:6px; }
  .metric-value { font-size:26px; font-weight:800; color:#1F3864; }
  .metric-sub   { font-size:12px; color:#aaa; margin-top:4px; }
  .mv-red   { color:#A32D2D !important; }
  .mv-green { color:#375623 !important; }
  .mv-amber { color:#7D5A00 !important; }
  .section-title {
    font-size:13px; font-weight:700; text-transform:uppercase;
    letter-spacing:.07em; color:#444;
    border-bottom:2px solid #1F3864;
    padding-bottom:6px; margin:1.5rem 0 .8rem;
  }
  .upload-hint {
    background:#eef2fb; border-radius:10px; padding:1.2rem 1.5rem;
    border:1px dashed #2E75B6; color:#1F3864;
    font-size:14px; line-height:1.8;
  }
  .stDownloadButton > button {
    background:#1F3864 !important; color:white !important;
    border-radius:8px !important; font-weight:700 !important;
    font-size:15px !important; padding:.7rem 1.5rem !important;
    width:100%; border:none !important;
  }
  .stDownloadButton > button:hover { background:#2E75B6 !important; }
  footer { visibility:hidden; }
</style>
""", unsafe_allow_html=True)

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("## Stock Analyzer")
    st.markdown("---")
    st.markdown("### Settings")
    snap_date = st.date_input(
        "Snapshot date",
        value=date.today(),
        help="Reference date for calculating days since last sale. Default is today.",
    )
    st.markdown("---")
    st.markdown("### Required columns")
    st.markdown("""
Your Excel file needs:
- **ITEM NAME**
- **QTY AT HAND**
- **QTY SOLD**
- **LAST SALES DATE**

Optional: UNIT COST PRICE, BARCODE, WAREHOUSE

Column names are case-insensitive.
""")
    st.markdown("---")
    st.markdown("### Report includes")
    st.markdown("""
1. Summary Dashboard
2. Slow Moving (31-90 days)
3. Dormant (91-180 days)
4. Near Dead (181+ days)
5. Dead Stock (never sold)
6. All Stock by Idle Days
7. Top 50 Products
8. Monthly Trend
9. Negative Stock (if any)
10. Full Chronological Register
""")
    st.markdown("---")
    st.caption("Stock Performance Analyzer v1.0")

# ── Main ──────────────────────────────────────────────────────────────────────
st.markdown("# Stock Performance Analyzer")
st.markdown("Upload your stock Excel file and download a full multi-sheet analysis report.")
st.markdown("---")

uploaded = st.file_uploader(
    "Upload stock Excel file",
    type=["xlsx","xls"],
    label_visibility="collapsed",
)

if uploaded is None:
    st.markdown("""
<div class="upload-hint">
  <strong>Drop your Excel file above or click Browse</strong><br>
  Accepts .xlsx and .xls formats<br>
  Analysis runs instantly - no data is stored or shared<br>
  Download your full report as a colour-coded Excel workbook
</div>
""", unsafe_allow_html=True)
    st.stop()


@st.cache_data(show_spinner=False)
def load_file(file_bytes, snap):
    raw = pd.read_excel(BytesIO(file_bytes))
    df  = analyzer.load_and_prepare(raw, snapshot_date=snap)
    m   = analyzer.get_summary_metrics(df)
    return df, m

with st.spinner("Analysing your data..."):
    try:
        file_bytes   = uploaded.read()
        df, metrics  = load_file(file_bytes, snap_date)
    except ValueError as e:
        st.error(f"{e}")
        st.stop()
    except Exception as e:
        st.error(f"Unexpected error: {e}")
        st.stop()

snap_str = pd.Timestamp(snap_date).strftime("%d %b %Y")
st.success(f"**{uploaded.name}** analysed — {metrics['total_skus']:,} SKUs | Snapshot: {snap_str}")

# ── KPIs ──────────────────────────────────────────────────────────────────────
st.markdown('<div class="section-title">Portfolio Overview</div>', unsafe_allow_html=True)
cols = st.columns(6)
kpis = [
    ("Total SKUs",       f"{metrics['total_skus']:,}",                      "",        ""),
    ("Active (<=30d)",   f"{metrics['active']:,}",                          "mv-green", "Sold in last 30 days"),
    ("Never Sold",       f"{metrics['never_sold']:,}",                      "mv-red",   f"{metrics['never_sold']/metrics['total_skus']*100:.1f}% of catalogue"),
    ("Total Revenue",    f"N{metrics['total_revenue']/1e6:.1f}M",           "",         "Cost-basis"),
    ("Capital at Risk",  f"N{metrics['idle_capital']/1e6:.2f}M",            "mv-red",   "Idle stock"),
    ("Data Issues",      f"{metrics['negative_stock']:,}",                  "mv-amber", "Negative stock SKUs"),
]
for col, (label, val, cls, sub) in zip(cols, kpis):
    with col:
        st.markdown(f"""
<div class="metric-card">
  <div class="metric-label">{label}</div>
  <div class="metric-value {cls}">{val}</div>
  <div class="metric-sub">{sub}</div>
</div>""", unsafe_allow_html=True)

# ── Bucket table ──────────────────────────────────────────────────────────────
st.markdown('<div class="section-title">Stock Health by Age Bucket</div>', unsafe_allow_html=True)
bucket_df   = analyzer.get_bucket_summary(df)
display_cols= ["Age Bucket","Category","Total SKUs","With Stock","Units in Stock","Capital Tied (N)","Avg Days Idle","Risk Level","Action"]
disp        = bucket_df[display_cols].copy()
disp["Capital Tied (N)"] = disp["Capital Tied (N)"].apply(lambda x: f"N{x:,.0f}" if pd.notna(x) else "---")
disp["Avg Days Idle"]    = disp["Avg Days Idle"].apply(lambda x: f"{x:.1f}" if pd.notna(x) and x else "---")
disp["Units in Stock"]   = disp["Units in Stock"].apply(lambda x: f"{x:,}")
disp["Total SKUs"]       = disp["Total SKUs"].apply(lambda x: f"{x:,}")
disp["With Stock"]       = disp["With Stock"].apply(lambda x: f"{x:,}")

def color_risk(val):
    colors = {
        "Low":      "background-color:#E2EFDA;color:#375623",
        "Medium":   "background-color:#FFF2CC;color:#7D5A00",
        "High":     "background-color:#FCE4D6;color:#843C0C",
        "Critical": "background-color:#F4CCCC;color:#7F0000",
    }
    return colors.get(val,"")

st.dataframe(disp.style.map(color_risk, subset=["Risk Level"]), use_container_width=True, hide_index=True)

# ── Top 15 preview ────────────────────────────────────────────────────────────
st.markdown('<div class="section-title">Top 15 Products by Revenue</div>', unsafe_allow_html=True)
top15 = analyzer.get_top_products(df, n=15, by="REVENUE")
top_disp = top15[["ITEM NAME","QTY SOLD","QTY AT HAND","REVENUE","STATUS"]].copy()
top_disp.columns = ["Product","Qty Sold","Qty At Hand","Revenue (N)","Status"]
top_disp["Revenue (N)"] = top_disp["Revenue (N)"].apply(lambda x: f"N{x:,.0f}")
top_disp["Qty Sold"]    = top_disp["Qty Sold"].apply(lambda x: f"{x:,.0f}")
top_disp.index = range(1, len(top_disp)+1)
st.dataframe(top_disp, use_container_width=True)

# ── Monthly trend ─────────────────────────────────────────────────────────────
st.markdown('<div class="section-title">Monthly Sales Trend</div>', unsafe_allow_html=True)
trend = analyzer.get_monthly_trend(df)
if len(trend) > 0:
    td = trend[["MONTH_STR","skus","qty_sold","revenue"]].copy()
    td.columns = ["Month","Active SKUs","Units Sold","Revenue (N)"]
    td["Revenue (N)"] = td["Revenue (N)"].apply(lambda x: f"N{x:,.0f}")
    td["Units Sold"]  = td["Units Sold"].apply(lambda x: f"{x:,.0f}")
    td = td.sort_values("Month", ascending=False).reset_index(drop=True)
    td.index = range(1, len(td)+1)
    st.dataframe(td, use_container_width=True)

# ── Alerts ────────────────────────────────────────────────────────────────────
st.markdown('<div class="section-title">Alerts</div>', unsafe_allow_html=True)
col1, col2 = st.columns(2)

with col1:
    near_so = analyzer.get_near_stockout(df, max_stock=10, min_sold=50)
    if len(near_so) > 0:
        st.warning(f"{len(near_so)} near-stockout products — proven sellers almost out of stock")
        so_disp = near_so[["ITEM NAME","QTY AT HAND","QTY SOLD"]].head(8).copy()
        so_disp.columns = ["Product","Stock Left","Total Sold"]
        so_disp.index = range(1, len(so_disp)+1)
        st.dataframe(so_disp, use_container_width=True)
    else:
        st.success("No near-stockout alerts")

with col2:
    neg = analyzer.get_negative_stock(df)
    if len(neg) > 0:
        st.error(f"{len(neg)} negative-inventory products — data integrity issue")
        neg_disp = neg[["ITEM NAME","QTY AT HAND","QTY SOLD"]].head(8).copy()
        neg_disp.columns = ["Product","Qty At Hand","Qty Sold"]
        neg_disp.index = range(1, len(neg_disp)+1)
        st.dataframe(neg_disp, use_container_width=True)
    else:
        st.success("No negative inventory detected")

# ── Download ──────────────────────────────────────────────────────────────────
st.markdown("---")
st.markdown('<div class="section-title">Download Full Report</div>', unsafe_allow_html=True)

with st.spinner("Building your Excel report..."):
    report_bytes = excel_builder.build_report(df, metrics)

fname = f"Stock_Analysis_{pd.Timestamp(snap_date).strftime('%Y%m%d')}.xlsx"
col_dl, col_info = st.columns([1,2])
with col_dl:
    st.download_button(
        label="Download Excel Report",
        data=report_bytes,
        file_name=fname,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
with col_info:
    st.markdown(f"""
**10 sheets included:**
Summary · Slow Moving · Dormant · Near Dead · Dead Stock ·
All Stock by Idle Days · Top Products · Monthly Trend ·
Negative Stock · Full Chronological Register

File will be saved as: `{fname}`
""")
