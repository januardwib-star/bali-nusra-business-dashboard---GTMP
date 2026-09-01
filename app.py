import os
from datetime import datetime
import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go

# =========================================================
# BALI NUSRA BUSINESS COMMAND CENTER V4
# =========================================================
st.set_page_config(
    page_title="Bali Nusra | Business Command Center",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

DATA_DIR = os.path.join(os.path.dirname(__file__), "data")

FILES = {
    "Channel": "Channel Dashboard.xlsx",
    "Halo": "Halo Dashboard.xlsx",
    "Skul.id": "Skul.id.xlsx",
    "FB Youth": "fb_youth_16082026.xlsx",
    "Semarak": "Program Semarak.xlsx",
}

# -------------------------
# THEME
# -------------------------
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');

html, body, [class*="css"] {
    font-family: Inter, sans-serif;
}
.block-container {
    max-width: 1600px;
    padding-top: 1.1rem;
    padding-bottom: 2.5rem;
}
[data-testid="stSidebar"] {
    border-right: 1px solid #E7EAF0;
}
.hero {
    background: linear-gradient(135deg, #0B1020 0%, #172554 55%, #0F766E 100%);
    color: white;
    border-radius: 24px;
    padding: 26px 30px;
    margin-bottom: 18px;
    box-shadow: 0 12px 35px rgba(15,23,42,.12);
}
.hero-title {
    font-size: 30px;
    font-weight: 800;
    letter-spacing: -0.04em;
    margin: 0;
}
.hero-sub {
    color: #D9E2F2;
    font-size: 13px;
    margin-top: 6px;
}
.badge {
    display: inline-block;
    padding: 5px 10px;
    border-radius: 999px;
    background: rgba(255,255,255,.12);
    color: #E8EEF8;
    font-size: 11px;
    font-weight: 700;
    margin-top: 12px;
}
.section-title {
    font-size: 20px;
    font-weight: 800;
    color: #101828;
    margin: 22px 0 10px;
}
.section-sub {
    color: #667085;
    font-size: 12px;
    margin-top: -6px;
    margin-bottom: 12px;
}
.kpi {
    background: #FFFFFF;
    border: 1px solid #E7EAF0;
    border-radius: 18px;
    padding: 15px 17px;
    min-height: 118px;
    box-shadow: 0 5px 18px rgba(16,24,40,.05);
}
.kpi-label {
    color: #667085;
    font-size: 10px;
    font-weight: 800;
    letter-spacing: .10em;
    text-transform: uppercase;
}
.kpi-value {
    color: #101828;
    font-size: 25px;
    font-weight: 800;
    margin-top: 7px;
}
.kpi-note {
    color: #667085;
    font-size: 11px;
    margin-top: 4px;
}
.insight {
    background: #F8FAFC;
    border: 1px solid #E7EAF0;
    border-radius: 16px;
    padding: 14px 16px;
    margin-bottom: 9px;
}
.insight b { color:#101828; }
.small { color:#667085; font-size:11px; }
.good { color:#067647; font-weight:800; }
.warn { color:#B54708; font-weight:800; }
.bad { color:#B42318; font-weight:800; }
div[data-testid="stMetric"] {
    background: #fff;
    border: 1px solid #E7EAF0;
    border-radius: 16px;
    padding: 12px 14px;
}
button[kind="secondary"] {
    border-radius: 12px;
}
.stTabs [data-baseweb="tab-list"] {
    gap: 8px;
}
.stTabs [data-baseweb="tab"] {
    border-radius: 10px;
    padding: 8px 14px;
}
</style>
""", unsafe_allow_html=True)

# -------------------------
# HELPERS
# -------------------------
def clean(df):
    if df is None or df.empty:
        return pd.DataFrame()
    x = df.copy()
    x.columns = [str(c).strip() for c in x.columns]
    return x.dropna(how="all").reset_index(drop=True)

def num(series):
    if series is None:
        return pd.Series(dtype="float64")
    if isinstance(series, pd.DataFrame):
        series = series.iloc[:, 0]
    if pd.api.types.is_numeric_dtype(series):
        return pd.to_numeric(series, errors="coerce")
    return pd.to_numeric(
        series.astype(str)
        .str.replace(",", "", regex=False)
        .str.replace("%", "", regex=False)
        .str.replace("Rp", "", regex=False)
        .str.strip(),
        errors="coerce",
    )

def money(v):
    if v is None or pd.isna(v):
        return "-"
    v = float(v)
    if abs(v) >= 1e9:
        return f"Rp {v/1e9:.2f} B"
    if abs(v) >= 1e6:
        return f"Rp {v/1e6:.2f} M"
    if abs(v) >= 1e3:
        return f"Rp {v/1e3:.1f} K"
    return f"Rp {v:,.0f}"

def count(v):
    if v is None or pd.isna(v):
        return "-"
    v = float(v)
    if abs(v) >= 1e6:
        return f"{v/1e6:.2f} M"
    if abs(v) >= 1e3:
        return f"{v/1e3:.1f} K"
    return f"{v:,.0f}"

def pct(v):
    if v is None or pd.isna(v):
        return "-"
    return f"{float(v)*100:.1f}%"

def pct_raw(v):
    if v is None or pd.isna(v):
        return np.nan
    v = float(v)
    return v / 100 if abs(v) > 1.5 else v

def find_col(df, names):
    """Robust column finder. Avoids DataFrame-return problem on duplicate columns."""
    if df is None or df.empty:
        return None
    cols = list(df.columns)
    lower = {str(c).strip().lower(): c for c in cols}
    for name in names:
        if name.lower() in lower:
            return lower[name.lower()]
    for name in names:
        key = name.lower()
        for c in cols:
            if key in str(c).lower():
                return c
    return None

def promote_header(df):
    x = clean(df)
    if x.empty:
        return x
    first = x.iloc[0].astype(str).tolist()
    nonempty = sum(v not in ("nan", "None", "") for v in first)
    unnamed = sum(str(c).lower().startswith("unnamed") for c in x.columns)
    if unnamed and nonempty >= max(3, int(len(x.columns) * .25)):
        headers = []
        for i, v in enumerate(first):
            v = str(v).strip()
            headers.append(v if v and v.lower() != "nan" else f"col_{i}")
        y = x.iloc[1:].copy()
        y.columns = headers
        return y.reset_index(drop=True)
    return x

@st.cache_data(show_spinner=False)
def read_book(path):
    if not os.path.exists(path):
        return {}
    result = {}
    try:
        xl = pd.ExcelFile(path)
        for sheet in xl.sheet_names:
            try:
                result[sheet] = promote_header(pd.read_excel(path, sheet_name=sheet))
            except Exception:
                result[sheet] = pd.DataFrame()
    except Exception:
        return {}
    return result

@st.cache_data(show_spinner=False)
def load_all():
    books = {}
    for key, filename in FILES.items():
        path = os.path.join(DATA_DIR, filename)
        if os.path.exists(path):
            books[key] = read_book(path)
    return books

B = load_all()

def all_values(terms):
    vals = set()
    for src in B:
        for df in B[src].values():
            c = find_col(df, terms)
            if c:
                vals.update(
                    df[c].dropna().astype(str).str.strip().tolist()
                )
    return sorted(
        v for v in vals
        if v and v.lower() not in {"nan", "none", "null"}
    )

def filter_df(df, region="All", branch="All", cluster="All", city="All"):
    x = clean(df)
    filters = [
        (region, ["regional", "region"]),
        (branch, ["branch", "branch_lacci"]),
        (cluster, ["cluster", "cluster_name", "cluster_lacci", "cluster_sales"]),
        (city, ["city", "kota", "kabupaten", "city_lacci"]),
    ]
    for value, terms in filters:
        if value != "All":
            c = find_col(x, terms)
            if c:
                x = x[x[c].astype(str).str.strip() == value]
    return x

def first_matching_sheet(src, keywords):
    if src not in B:
        return None
    for sheet, df in B[src].items():
        name = sheet.lower()
        if all(k.lower() in name for k in keywords):
            return sheet
    for sheet, df in B[src].items():
        if any(k.lower() in sheet.lower() for k in keywords):
            return sheet
    return None

def latest_sheet(src, keyword="raw"):
    if src not in B:
        return None
    candidates = [
        s for s in B[src]
        if keyword.lower() in s.lower()
    ]
    if not candidates:
        return None
    return candidates[-1]

def chart(fig, height=390):
    fig.update_layout(
        height=height,
        margin=dict(l=10, r=10, t=48, b=10),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Inter", color="#344054"),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0),
        hovermode="x unified",
    )
    st.plotly_chart(fig, use_container_width=True, config={"displaylogo": False})

def kpi_card(container, label, value, note=""):
    container.markdown(
        f"""
        <div class="kpi">
            <div class="kpi-label">{label}</div>
            <div class="kpi-value">{value}</div>
            <div class="kpi-note">{note}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

def title(text, sub=None):
    st.markdown(f'<div class="section-title">{text}</div>', unsafe_allow_html=True)
    if sub:
        st.markdown(f'<div class="section-sub">{sub}</div>', unsafe_allow_html=True)

def insight(text):
    st.markdown(f'<div class="insight">💡 {text}</div>', unsafe_allow_html=True)

# -------------------------
# SIDEBAR
# -------------------------
with st.sidebar:
    st.markdown("## 📊 BALI NUSRA")
    st.caption("Business Command Center • V4")
    if st.button("🔄 Refresh data", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

    st.divider()

    page = st.radio(
        "NAVIGATION",
        [
            "🏠 Executive Overview",
            "📈 Channel Performance",
            "💰 Revenue & RGB",
            "📱 Halo & Device",
            "🎓 Skul.id",
            "⚔️ FB Youth",
            "🚀 Program Semarak",
            "🗂️ Data Explorer",
        ],
    )

    st.divider()
    st.caption("DATA SOURCES")
    for key in FILES:
        exists = os.path.exists(os.path.join(DATA_DIR, FILES[key]))
        st.write(("🟢 " if exists else "🔴 ") + key)

# -------------------------
# GLOBAL FILTERS
# -------------------------
st.markdown("""
<div class="hero">
    <div class="hero-title">Bali Nusra Business Command Center</div>
    <div class="hero-sub">Executive performance • revenue • channel • digital adoption • competitive share • program effectiveness</div>
    <div class="badge">LIVE FROM EXCEL DATA • INTERACTIVE ANALYTICS</div>
</div>
""", unsafe_allow_html=True)

f1, f2, f3, f4 = st.columns(4)
regions = ["All"] + all_values(["regional", "region"])[:300]
branches = ["All"] + all_values(["branch", "branch_lacci"])[:300]
clusters = ["All"] + all_values(["cluster", "cluster_name", "cluster_lacci", "cluster_sales"])[:400]
cities = ["All"] + all_values(["city", "kota", "kabupaten", "city_lacci"])[:500]

region = f1.selectbox("Region", regions)
branch = f2.selectbox("Branch", branches)
cluster = f3.selectbox("Cluster", clusters)
city = f4.selectbox("City / Kabupaten", cities)

# -------------------------
# EXECUTIVE OVERVIEW
# -------------------------
if page == "🏠 Executive Overview":
    title("Executive Scorecard", "Satu halaman untuk melihat kesehatan bisnis Bali Nusra.")

    # Revenue
    revenue = 0
    revenue_rows = 0
    rev_sheet = first_matching_sheet("Channel", ["rev"])
    if rev_sheet:
        x = filter_df(B["Channel"][rev_sheet], region, branch, cluster, city)
        rev_col = find_col(x, ["REV_ALL_BB", "REV_ALL", "Total Omzet", "Total_Omzet"])
        if rev_col:
            revenue = num(x[rev_col]).sum()
            revenue_rows = len(x)

    # RGB
    rgb = 0
    rgb_sheet = first_matching_sheet("Channel", ["rgb"])
    if rgb_sheet:
        x = filter_df(B["Channel"][rgb_sheet], region, branch, cluster, city)
        c = find_col(x, ["rgb_all", "rgb"])
        if c:
            rgb = num(x[c]).sum()

    # Halo
    halo = 0
    halo_sheet = first_matching_sheet("Halo", ["rev", "halo"])
    if halo_sheet:
        x = filter_df(B["Halo"][halo_sheet], region, branch, cluster, city)
        c = find_col(x, ["subs", "subscription", "subscriber"])
        if c:
            halo = num(x[c]).sum()

    # Skul
    active = productive = schools = 0
    skul_sheet = latest_sheet("Skul.id", "raw")
    if skul_sheet:
        x = filter_df(B["Skul.id"][skul_sheet], region, branch, cluster, city)
        ca = find_col(x, ["User Active"])
        cp = find_col(x, ["User Productive"])
        ci = find_col(x, ["ID Sekolah", "School ID"])
        if ca:
            active = num(x[ca]).sum()
        if cp:
            productive = num(x[cp]).sum()
        if ci:
            schools = x[ci].nunique()

    # Program Semarak
    so = fr = 0
    sem_sheet = None
    if "Semarak" in B:
        for s, df in B["Semarak"].items():
            if find_col(df, ["uniq_SO", "unique SO"]) and find_col(df, ["uniq_FR", "unique FR"]):
                sem_sheet = s
                break
    if sem_sheet:
        x = filter_df(B["Semarak"][sem_sheet], region, branch, cluster, city)
        cso = find_col(x, ["uniq_SO", "unique SO"])
        cfr = find_col(x, ["uniq_FR", "unique FR"])
        if cso:
            so = num(x[cso]).sum()
        if cfr:
            fr = num(x[cfr]).sum()

    cs = st.columns(6)
    kpi_card(cs[0], "REVENUE", money(revenue), f"{revenue_rows:,} filtered rows")
    kpi_card(cs[1], "RGB", count(rgb), "registered/base metric")
    kpi_card(cs[2], "HALO SUBS", count(halo), "subscription dataset")
    kpi_card(cs[3], "SKUL ACTIVE", count(active), "latest raw period")
    kpi_card(cs[4], "SKUL PRODUCTIVE", count(productive), "latest raw period")
    kpi_card(cs[5], "FR RATE", pct(fr / so if so else np.nan), f"{count(fr)} FR / {count(so)} SO")

    title("Management Pulse", "Fokus bukan hanya angka, tetapi apa yang perlu diperhatikan manajemen.")

    a, b = st.columns([1.35, 1])

    with a:
        if rev_sheet:
            x = filter_df(B["Channel"][rev_sheet], region, branch, cluster, city)
            cc = find_col(x, ["CITY", "City", "Kabupaten", "Kota"])
            mc = find_col(x, ["REV_ALL_BB", "REV_ALL", "Total Omzet"])
            if cc and mc:
                g = x.groupby(cc, dropna=False)[mc].sum().reset_index()
                g = g.sort_values(mc, ascending=False).head(15)
                fig = px.bar(
                    g, x=mc, y=cc, orientation="h",
                    text=mc, title="Revenue by City — Top 15",
                    template="plotly_white"
                )
                fig.update_traces(texttemplate="%{text:.3s}", textposition="outside")
                chart(fig)

    with b:
        if "FB Youth" in B:
            s = first_matching_sheet("FB Youth", ["fb"])
            if s:
                x = filter_df(B["FB Youth"][s], region, branch, cluster, city)
                terr = find_col(x, ["TERITORI", "Territory"])
                op_cols = [
                    c for c in [
                        find_col(x, ["TSEL"]),
                        find_col(x, ["IOH"]),
                        find_col(x, ["XL"]),
                    ] if c
                ]
                if terr and op_cols:
                    g = x.groupby(terr)[op_cols].mean().reset_index()
                    melted = g.melt(id_vars=[terr], var_name="Operator", value_name="Share")
                    melted["Share"] = melted["Share"].apply(pct_raw)
                    fig = px.bar(
                        melted, x="Share", y=terr, color="Operator",
                        orientation="h", barmode="group",
                        title="Competitive Share by Territory",
                        template="plotly_white"
                    )
                    fig.update_xaxes(tickformat=".0%")
                    chart(fig)

    title("Management Signals")

    if revenue > 0:
        insight(f"<b>Revenue</b> saat ini {money(revenue)} berdasarkan filter aktif.")
    if productive > 0 and active > 0:
        insight(f"<b>Skul productivity</b> = <b>{pct(productive/active)}</b>. Ini menunjukkan proporsi active users yang sudah productive.")
    if so > 0:
        rate = fr / so
        cls = "good" if rate >= .8 else ("warn" if rate >= .5 else "bad")
        insight(f"<b>Semarak FR Rate</b> = <span class='{cls}'>{pct(rate)}</span>. Gunakan filter Cluster/City untuk menemukan area yang paling lemah.")
    if revenue == 0 and rgb == 0 and halo == 0:
        insight("Data KPI belum dapat dipetakan. Cek nama kolom Excel atau gunakan menu Data Explorer.")

# -------------------------
# CHANNEL PERFORMANCE
# -------------------------
elif page == "📈 Channel Performance":
    title("Channel Performance", "Analisis produktivitas channel, revenue, RGB dan outlet.")

    tabs = st.tabs(["Revenue", "RGB", "Outlet", "Channel Mix"])

    with tabs[0]:
        s = first_matching_sheet("Channel", ["rev"])
        if s:
            x = filter_df(B["Channel"][s], region, branch, cluster, city)
            m = find_col(x, ["REV_ALL_BB", "REV_ALL", "Total Omzet"])
            d = find_col(x, ["TGL", "Tanggal", "Date", "Period"])
            c = find_col(x, ["CITY", "City", "Kota", "Kabupaten"])

            k1, k2, k3 = st.columns(3)
            total = num(x[m]).sum() if m else 0
            k1.metric("Revenue", money(total))
            k2.metric("Rows", f"{len(x):,}")
            k3.metric("Cities", f"{x[c].nunique():,}" if c else "-")

            left, right = st.columns(2)
            with left:
                if c and m:
                    g = x.groupby(c)[m].sum().reset_index().sort_values(m, ascending=False).head(15)
                    chart(px.bar(g, x=m, y=c, orientation="h", title="Top Cities by Revenue", template="plotly_white"))
            with right:
                if d and m:
                    y = x.copy()
                    y[d] = pd.to_datetime(y[d], errors="coerce")
                    y = y.dropna(subset=[d])
                    g = y.groupby(d)[m].sum().reset_index().sort_values(d)
                    chart(px.area(g, x=d, y=m, title="Revenue Trend", template="plotly_white"))

            if m:
                st.download_button(
                    "⬇️ Download filtered Revenue CSV",
                    x.to_csv(index=False).encode("utf-8"),
                    "channel_revenue_filtered.csv",
                    "text/csv",
                )

    with tabs[1]:
        s = first_matching_sheet("Channel", ["rgb"])
        if s:
            x = filter_df(B["Channel"][s], region, branch, cluster, city)
            m = find_col(x, ["rgb_all", "rgb"])
            d = find_col(x, ["period", "tanggal", "date"])
            c = find_col(x, ["kabupaten", "kota", "city"])
            total = num(x[m]).sum() if m else 0
            st.metric("Total RGB", count(total))

            left, right = st.columns(2)
            with left:
                if c and m:
                    g = x.groupby(c)[m].sum().reset_index().sort_values(m, ascending=False).head(15)
                    chart(px.bar(g, x=m, y=c, orientation="h", title="RGB by City", template="plotly_white"))
            with right:
                if d and m:
                    y = x.copy()
                    y[d] = pd.to_datetime(y[d], errors="coerce")
                    y = y.dropna(subset=[d])
                    g = y.groupby(d)[m].sum().reset_index().sort_values(d)
                    chart(px.line(g, x=d, y=m, markers=True, title="RGB Trend", template="plotly_white"))

    with tabs[2]:
        s = first_matching_sheet("Channel", ["omzet", "outlet"])
        if s:
            x = filter_df(B["Channel"][s], region, branch, cluster, city)
            m = find_col(x, ["Total Omzet", "Omzet"])
            c = find_col(x, ["cluster", "cluster_name"])
            outlet = find_col(x, ["fisik", "outlet", "channel"])
            total = num(x[m]).sum() if m else 0
            st.metric("Outlet Revenue", money(total))

            if c and m:
                g = x.groupby(c)[m].sum().reset_index().sort_values(m, ascending=False)
                chart(px.bar(g, x=c, y=m, text=m, title="Revenue by Cluster", template="plotly_white"))
            if outlet and m:
                g = x.groupby(outlet)[m].sum().reset_index().sort_values(m, ascending=False).head(20)
                chart(px.bar(g, x=m, y=outlet, orientation="h", title="Top Outlet / Physical Channel", template="plotly_white"))

    with tabs[3]:
        # Cross-source channel mix
        rows = []
        if "Halo" in B:
            s = first_matching_sheet("Halo", ["rev"])
            if s:
                x = filter_df(B["Halo"][s], region, branch, cluster, city)
                c = find_col(x, ["channel", "Channel (standar dashboard)"])
                if c:
                    rows.append(x[c].value_counts().rename_axis("Channel").reset_index(name="Count"))
        if rows:
            g = pd.concat(rows).groupby("Channel")["Count"].sum().reset_index().sort_values("Count", ascending=False)
            chart(px.pie(g, names="Channel", values="Count", hole=.58, title="Activation Mix by Channel", template="plotly_white"))

# -------------------------
# REVENUE & RGB
# -------------------------
elif page == "💰 Revenue & RGB":
    title("Revenue & RGB", "Trend, contribution dan konsentrasi performance.")

    srev = first_matching_sheet("Channel", ["rev"])
    srgb = first_matching_sheet("Channel", ["rgb"])

    if srev:
        x = filter_df(B["Channel"][srev], region, branch, cluster, city)
        m = find_col(x, ["REV_ALL_BB", "REV_ALL", "Total Omzet"])
        d = find_col(x, ["TGL", "Tanggal", "Date", "Period"])
        c = find_col(x, ["CITY", "City", "Kota", "Kabupaten"])

        k1, k2, k3 = st.columns(3)
        total = num(x[m]).sum() if m else 0
        k1.metric("Revenue", money(total))
        if c and m:
            g = x.groupby(c)[m].sum().reset_index()
            k2.metric("Top City", str(g.loc[g[m].idxmax(), c]) if not g.empty else "-")
            k3.metric("Top City Revenue", money(g[m].max()) if not g.empty else "-")

        left, right = st.columns(2)
        with left:
            if c and m:
                g = x.groupby(c)[m].sum().reset_index().sort_values(m, ascending=False).head(12)
                chart(px.bar(g, x=m, y=c, orientation="h", title="Revenue Concentration", template="plotly_white"))
        with right:
            if d and m:
                y = x.copy()
                y[d] = pd.to_datetime(y[d], errors="coerce")
                y = y.dropna(subset=[d])
                g = y.groupby(d)[m].sum().reset_index().sort_values(d)
                chart(px.line(g, x=d, y=m, markers=True, title="Revenue Daily / Period Trend", template="plotly_white"))

    if srgb:
        x = filter_df(B["Channel"][srgb], region, branch, cluster, city)
        m = find_col(x, ["rgb_all", "rgb"])
        c = find_col(x, ["brand", "channel", "operator"])
        if m:
            st.markdown("---")
            st.subheader("RGB Contribution")
            left, right = st.columns(2)
            with left:
                if c:
                    g = x.groupby(c)[m].sum().reset_index().sort_values(m, ascending=False)
                    chart(px.pie(g, names=c, values=m, hole=.58, title="RGB Mix", template="plotly_white"))
            with right:
                if c:
                    g = x.groupby(c)[m].sum().reset_index().sort_values(m, ascending=False)
                    g["Share"] = g[m] / g[m].sum()
                    chart(px.bar(g, x="Share", y=c, orientation="h", title="RGB Share by Segment", template="plotly_white"))
                    st.dataframe(g, use_container_width=True, hide_index=True)

# -------------------------
# HALO & DEVICE
# -------------------------
elif page == "📱 Halo & Device":
    title("Halo & Device", "Activation, revenue, channel contribution dan status.")

    s = first_matching_sheet("Halo", ["rev"])
    if s:
        x = filter_df(B["Halo"][s], region, branch, cluster, city)
        subs = find_col(x, ["subs", "subscription", "subscriber"])
        price = find_col(x, ["price", "revenue", "amount"])
        channel = find_col(x, ["channel", "Channel (standar dashboard)"])
        flag = find_col(x, ["flag", "status"])
        date = find_col(x, ["date", "tanggal", "period"])

        total_sub = num(x[subs]).sum() if subs else len(x)
        total_rev = num(x[price]).sum() if price else 0

        a, b, c, d = st.columns(4)
        a.metric("Subscriptions", count(total_sub))
        b.metric("Revenue / Price", money(total_rev))
        c.metric("Channels", f"{x[channel].nunique():,}" if channel else "-")
        d.metric("Rows", f"{len(x):,}")

        left, right = st.columns(2)
        with left:
            if channel:
                g = x.groupby(channel).size().reset_index(name="Activations").sort_values("Activations", ascending=False).head(15)
                chart(px.bar(g, x="Activations", y=channel, orientation="h", title="Activation by Channel", template="plotly_white"))
        with right:
            if flag:
                g = x.groupby(flag).size().reset_index(name="Count")
                chart(px.pie(g, names=flag, values="Count", hole=.58, title="Status / Flag Mix", template="plotly_white"))

        if date and subs:
            y = x.copy()
            y[date] = pd.to_datetime(y[date], errors="coerce")
            y["_subs"] = num(y[subs])
            y = y.dropna(subset=[date])
            g = y.groupby(date)["_subs"].sum().reset_index()
            chart(px.area(g, x=date, y="_subs", title="Subscription Trend", template="plotly_white"))

        with st.expander("🔎 Detail data"):
            st.dataframe(x, use_container_width=True, height=500, hide_index=True)

# -------------------------
# SKUL.ID
# -------------------------
elif page == "🎓 Skul.id":
    title("Skul.id — Digital Adoption Funnel", "Pantau school adoption dari total users → active → productive.")

    raws = []
    if "Skul.id" in B:
        for s, df in B["Skul.id"].items():
            if "raw" in s.lower():
                raws.append((s, df))

    if raws:
        labels = [s for s, _ in raws]
        pick = st.selectbox("Period / Raw Sheet", labels, index=len(labels)-1)
        x = filter_df(dict(raws)[pick], region, branch, cluster, city)

        school = find_col(x, ["ID Sekolah", "School ID"])
        total_users = find_col(x, ["Total Users", "Total User"])
        active = find_col(x, ["User Active", "Active Users"])
        productive = find_col(x, ["User Productive", "Productive Users"])
        cityc = find_col(x, ["Kota", "City", "Kabupaten"])

        total_school = x[school].nunique() if school else 0
        total_u = num(x[total_users]).sum() if total_users else 0
        active_u = num(x[active]).sum() if active else 0
        productive_u = num(x[productive]).sum() if productive else 0

        a, b, c, d = st.columns(4)
        a.metric("Schools", count(total_school))
        b.metric("Total Users", count(total_u))
        c.metric("Active Users", count(active_u))
        d.metric("Productive Users", count(productive_u))

        if active_u:
            st.progress(min(active_u / total_u, 1.0) if total_u else 0, text=f"Active conversion: {pct(active_u/total_u) if total_u else '-'}")
        if productive_u:
            st.progress(min(productive_u / active_u, 1.0) if active_u else 0, text=f"Productive conversion: {pct(productive_u/active_u) if active_u else '-'}")

        left, right = st.columns(2)

        with left:
            stages = pd.DataFrame({
                "Stage": ["Total Users", "Active", "Productive"],
                "Users": [total_u, active_u, productive_u]
            })
            chart(px.funnel(stages, x="Users", y="Stage", title="User Conversion Funnel", template="plotly_white"))

        with right:
            if cityc and active:
                g = x.groupby(cityc)[active].sum().reset_index().sort_values(active, ascending=False).head(15)
                chart(px.bar(g, x=active, y=cityc, orientation="h", title="Active Users by City", template="plotly_white"))

        if cityc and active and productive:
            g = x.groupby(cityc).agg(
                Active=(active, "sum"),
                Productive=(productive, "sum")
            ).reset_index()
            g["Productive Rate"] = np.where(g["Active"] > 0, g["Productive"]/g["Active"], 0)
            st.subheader("City Productivity Ranking")
            st.dataframe(
                g.sort_values("Productive Rate", ascending=False),
                use_container_width=True,
                hide_index=True,
                column_config={
                    "Productive Rate": st.column_config.ProgressColumn(
                        "Productive Rate", min_value=0, max_value=1, format="%.1%"
                    )
                }
            )
    else:
        st.warning("Sheet Raw Data Skul.id belum ditemukan.")

# -------------------------
# FB YOUTH
# -------------------------
elif page == "⚔️ FB Youth":
    title("FB Youth — Competitive Intelligence", "Bandingkan TSEL vs IOH vs XL per territory dan periode.")

    s = first_matching_sheet("FB Youth", ["fb"])
    if s:
        x = filter_df(B["FB Youth"][s], region, branch, cluster, city)
        terr = find_col(x, ["TERITORI", "Territory"])
        dt = find_col(x, ["TANGGAL", "Tanggal", "Date"])
        tsel = find_col(x, ["TSEL"])
        ioh = find_col(x, ["IOH"])
        xl = find_col(x, ["XL"])

        ops = [c for c in [tsel, ioh, xl] if c]
        if dt:
            x[dt] = pd.to_datetime(x[dt], errors="coerce")

        if ops:
            for c in ops:
                x[c] = x[c].apply(pct_raw)

        left, right = st.columns(2)
        with left:
            if terr:
                g = x.groupby(terr)[ops].mean().reset_index()
                melted = g.melt(id_vars=[terr], var_name="Operator", value_name="Share")
                chart(px.bar(melted, x="Share", y=terr, color="Operator", orientation="h", barmode="group",
                             title="Average Share by Territory", template="plotly_white"))
        with right:
            if dt:
                g = x.groupby(dt)[ops].mean().reset_index().sort_values(dt)
                melted = g.melt(id_vars=[dt], var_name="Operator", value_name="Share")
                fig = px.line(melted, x=dt, y="Share", color="Operator", markers=True,
                              title="Competitive Share Trend", template="plotly_white")
                fig.update_yaxes(tickformat=".0%")
                chart(fig)

        if terr and ops:
            g = x.groupby(terr)[ops].mean().reset_index()
            g["Winner"] = g[ops].idxmax(axis=1)
            winner_counts = g["Winner"].value_counts().reset_index()
            winner_counts.columns = ["Operator", "Territories Won"]
            st.subheader("Territory Win Count")
            chart(px.bar(winner_counts, x="Operator", y="Territories Won", text="Territories Won",
                         title="Territory Win Count", template="plotly_white"))

        with st.expander("🔎 Detail competitive data"):
            st.dataframe(x, use_container_width=True, height=500, hide_index=True)

# -------------------------
# SEMARAK
# -------------------------
elif page == "🚀 Program Semarak":
    title("Program Semarak — SO → FR Conversion", "Cari cluster dan branch dengan conversion terbaik / terlemah.")

    s = None
    if "Semarak" in B:
        for sheet, df in B["Semarak"].items():
            if find_col(df, ["uniq_SO", "unique SO"]) and find_col(df, ["uniq_FR", "unique FR"]):
                s = sheet
                break

    if s:
        x = filter_df(B["Semarak"][s], region, branch, cluster, city)
        so_col = find_col(x, ["uniq_SO", "unique SO"])
        fr_col = find_col(x, ["uniq_FR", "unique FR"])
        cluster_col = find_col(x, ["cluster_name", "cluster"])
        branch_col = find_col(x, ["branch"])
        status_col = find_col(x, ["status_fr", "status"])

        so = num(x[so_col]).sum() if so_col else 0
        fr = num(x[fr_col]).sum() if fr_col else 0
        rate = fr / so if so else 0

        a, b, c, d = st.columns(4)
        a.metric("Unique SO", count(so))
        b.metric("Successful FR", count(fr))
        c.metric("FR Rate", pct(rate))
        d.metric("Data Rows", f"{len(x):,}")

        if cluster_col:
            g = x.groupby(cluster_col).agg(
                SO=(so_col, "sum"),
                FR=(fr_col, "sum")
            ).reset_index()
            g["FR Rate"] = np.where(g["SO"] > 0, g["FR"] / g["SO"], 0)

            left, right = st.columns(2)
            with left:
                chart(px.bar(g.sort_values("SO", ascending=False), x=cluster_col, y=["SO", "FR"],
                             barmode="group", title="SO vs FR by Cluster", template="plotly_white"))
            with right:
                chart(px.bar(g.sort_values("FR Rate", ascending=False), x="FR Rate", y=cluster_col,
                             orientation="h", title="FR Rate Ranking", template="plotly_white"))

            st.subheader("Cluster Performance Table")
            st.dataframe(
                g.sort_values("FR Rate", ascending=False),
                use_container_width=True,
                hide_index=True,
                column_config={
                    "FR Rate": st.column_config.ProgressColumn(
                        "FR Rate", min_value=0, max_value=1, format="%.1%"
                    )
                }
            )

        if branch_col and fr_col:
            g = x.groupby(branch_col)[fr_col].sum().reset_index().sort_values(fr_col, ascending=False).head(15)
            chart(px.bar(g, x=fr_col, y=branch_col, orientation="h", title="Top Branch by Successful FR", template="plotly_white"))

        if status_col:
            g = x[status_col].value_counts().reset_index()
            g.columns = [status_col, "Count"]
            chart(px.pie(g, names=status_col, values="Count", hole=.58, title="FR Status Mix", template="plotly_white"))
    else:
        st.warning("Raw data Program Semarak belum bisa dipetakan.")

# -------------------------
# DATA EXPLORER
# -------------------------
else:
    title("Data Explorer", "Telusuri data mentah tanpa keluar dari dashboard.")

    if B:
        src = st.selectbox("Source", list(B.keys()))
        sheet = st.selectbox("Sheet", list(B[src].keys()))
        x = filter_df(B[src][sheet], region, branch, cluster, city)

        c1, c2, c3 = st.columns(3)
        c1.metric("Rows", f"{len(x):,}")
        c2.metric("Columns", f"{len(x.columns):,}")
        c3.metric("Missing Cells", f"{int(x.isna().sum().sum()):,}")

        search = st.text_input("🔎 Search text in all columns")
        if search:
            mask = x.astype(str).apply(
                lambda row: row.str.contains(search, case=False, na=False).any(),
                axis=1
            )
            x = x[mask]

        st.dataframe(x, use_container_width=True, height=620, hide_index=True)

        st.download_button(
            "⬇️ Download filtered data",
            x.to_csv(index=False).encode("utf-8"),
            "bali_nusra_filtered.csv",
            "text/csv",
        )
    else:
        st.error("Tidak ada file Excel di folder data/.")

# Footer
st.markdown("---")
st.caption(
    "Bali Nusra Business Command Center V4 • "
    f"Last refresh: {datetime.now().strftime('%d %b %Y %H:%M')} • "
    "Data source: Excel files in /data"
)
