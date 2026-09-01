import os
from datetime import datetime
import re
import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px

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

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
html,body,[class*="css"]{font-family:Inter,sans-serif}
.block-container{max-width:1600px;padding-top:1rem}
[data-testid="stSidebar"]{border-right:1px solid #e5e7eb}
.hero{background:linear-gradient(135deg,#0f172a 0%,#172554 48%,#0f766e 100%);color:#fff;border-radius:24px;padding:28px 32px;margin-bottom:18px;box-shadow:0 14px 40px rgba(15,23,42,.14)}
.hero h1{font-size:32px;font-weight:800;margin:0;letter-spacing:-.04em}.hero p{margin:7px 0 0;color:#dbeafe;font-size:13px}
.badge{display:inline-block;margin-top:13px;background:rgba(255,255,255,.13);padding:6px 11px;border-radius:999px;font-size:11px;font-weight:700}
.kpi{background:#fff;border:1px solid #e5e7eb;border-radius:18px;padding:15px 17px;min-height:112px;box-shadow:0 6px 22px rgba(15,23,42,.05)}
.kpi .label{font-size:10px;font-weight:800;letter-spacing:.1em;color:#667085;text-transform:uppercase}.kpi .value{font-size:25px;font-weight:800;color:#111827;margin-top:7px}.kpi .note{font-size:11px;color:#667085;margin-top:4px}
.section{font-size:20px;font-weight:800;color:#101828;margin:22px 0 8px}.sub{font-size:12px;color:#667085;margin-bottom:12px}
.insight{background:#f8fafc;border:1px solid #e5e7eb;border-radius:15px;padding:13px 15px;margin:7px 0}
</style>
""", unsafe_allow_html=True)


def clean(df):
    if df is None or not isinstance(df, pd.DataFrame) or df.empty:
        return pd.DataFrame()
    x = df.copy()
    seen = {}
    new = []
    for i, c in enumerate(x.columns):
        base = str(c).strip()
        if not base or base.lower() in {"nan", "none"}:
            base = f"Column_{i+1}"
        base = re.sub(r"\s+", " ", base)
        n = seen.get(base.casefold(), 0) + 1
        seen[base.casefold()] = n
        new.append(base if n == 1 else f"{base}__{n}")
    x.columns = new
    return x.dropna(how="all").reset_index(drop=True)


def norm(value):
    return re.sub(r"[^a-z0-9]+", "", str(value).strip().casefold())


def safe_series(df, col):
    if df is None or not isinstance(df, pd.DataFrame) or col is None or col not in df.columns:
        return pd.Series(index=df.index if isinstance(df, pd.DataFrame) else [], dtype="object")
    s = df[col]
    return s.iloc[:, 0] if isinstance(s, pd.DataFrame) else s


def num(s):
    if isinstance(s, pd.DataFrame):
        s = s.iloc[:, 0]
    if s is None:
        return pd.Series(dtype=float)
    if pd.api.types.is_numeric_dtype(s):
        return pd.to_numeric(s, errors="coerce")
    z = s.astype(str)
    z = z.str.replace(r"\((.*?)\)", r"-\1", regex=True)
    z = z.str.replace(",", "", regex=False)
    z = z.str.replace("%", "", regex=False)
    z = z.str.replace("Rp", "", regex=False)
    z = z.str.replace(" ", "", regex=False)
    z = z.str.replace(r"[^0-9eE\.\-\+]", "", regex=True)
    return pd.to_numeric(z, errors="coerce")


def money(v):
    if v is None or pd.isna(v): return "-"
    v = float(v)
    if abs(v) >= 1e9: return f"Rp {v/1e9:.2f} B"
    if abs(v) >= 1e6: return f"Rp {v/1e6:.2f} M"
    if abs(v) >= 1e3: return f"Rp {v/1e3:.1f} K"
    return f"Rp {v:,.0f}"


def count(v):
    if v is None or pd.isna(v): return "-"
    v = float(v)
    if abs(v) >= 1e6: return f"{v/1e6:.2f} M"
    if abs(v) >= 1e3: return f"{v/1e3:.1f} K"
    return f"{v:,.0f}"


def pct(v):
    if v is None or pd.isna(v): return "-"
    return f"{float(v)*100:.1f}%"


def pct_value(v):
    if v is None or pd.isna(v): return np.nan
    v = float(v)
    return v / 100 if abs(v) > 1.5 else v


def find_col(df, aliases):
    if df is None or df.empty:
        return None
    cols = list(df.columns)
    exact = {norm(c): c for c in cols}
    for alias in aliases:
        if norm(alias) in exact:
            return exact[norm(alias)]
    for alias in sorted(aliases, key=lambda x: len(norm(x)), reverse=True):
        key = norm(alias)
        if not key:
            continue
        for c in cols:
            nc = norm(c)
            if key in nc or nc in key:
                return c
    return None


def unique_values(df, col):
    if df is None or df.empty or col is None:
        return []
    s = safe_series(df, col).dropna().astype(str).str.strip()
    bad = {"", "nan", "none", "null", "-", "all"}
    return sorted([v for v in s.unique() if v.casefold() not in bad], key=lambda x: x.casefold())


def likely_header(raw):
    if raw.empty:
        return 0
    keywords = [
        "region", "regional", "branch", "cluster", "city", "kota", "kabupaten",
        "revenue", "rev", "rgb", "halo", "user", "active", "productive", "school",
        "operator", "share", "territory", "tanggal", "date", "period", "so", "fr",
        "channel", "tsel", "ioh", "xl", "student", "siswa", "id sekolah",
    ]
    best_row, best_score = 0, -10**9
    for r in range(min(12, len(raw))):
        vals = [str(v).strip() for v in raw.iloc[r].tolist()]
        nonempty = [v for v in vals if v.lower() not in {"nan", "none", ""}]
        score = sum(any(k in norm(v) for k in keywords) for v in nonempty)
        score += min(len(nonempty), 25) * 0.05
        score -= sum(v.lower().startswith("unnamed") for v in vals) * 0.5
        if score > best_score:
            best_score, best_row = score, r
    return best_row


@st.cache_data(show_spinner=False)
def read_book(path):
    out = {}
    if not os.path.exists(path):
        return out
    try:
        xl = pd.ExcelFile(path)
    except Exception:
        return out
    for sheet in xl.sheet_names:
        try:
            raw = pd.read_excel(path, sheet_name=sheet, header=None)
            if raw.empty:
                continue
            h = likely_header(raw)
            x = pd.read_excel(path, sheet_name=sheet, header=h)
            x = clean(x)
            if x.empty:
                continue
            out[str(sheet).strip()] = x
        except Exception:
            continue
    return out


@st.cache_data(show_spinner=False)
def load_all():
    return {key: read_book(os.path.join(DATA_DIR, filename)) for key, filename in FILES.items()}


B = load_all()

DIM = {
    "region": ["regional", "region", "region_name", "region_lacci", "regional_name"],
    "branch": ["branch", "branch_name", "branch_lacci", "branch_sales"],
    "cluster": ["cluster", "cluster_name", "cluster_lacci", "cluster_sales"],
    "city": ["city", "city_name", "kota", "kabupaten", "city_lacci", "city_name_lacci"],
    "date": ["date", "tanggal", "tgl", "period", "periode", "month", "bulan", "week", "minggu"],
}

METRICS = {
    "revenue": ["rev_all_bb", "rev_all", "revenue", "total revenue", "total_omzet", "total omzet", "omzet", "sales revenue", "revenue_all"],
    "rgb": ["rgb_all", "rgb", "total rgb", "rgb users"],
    "halo": ["subs", "subscription", "subscriber", "halo subs", "halo subscriber", "halo"],
    "price": ["price", "amount", "revenue", "nominal", "value"],
    "active": ["user active", "active users", "active_user", "active", "users active"],
    "productive": ["user productive", "productive users", "productive_user", "productive"],
    "users": ["total users", "total user", "users", "user count"],
    "school": ["school", "schools", "sekolah", "school count", "jumlah sekolah"],
    "school_id": ["id sekolah", "school id", "school_id"],
    "student": ["student", "students", "siswa", "student count"],
    "so": ["uniq_so", "unique so", "total so", "sales order", "so"],
    "fr": ["uniq_fr", "unique fr", "total fr", "first recharge", "fr"],
}


def detect_dim(df, dim):
    return find_col(df, DIM.get(dim, [dim]))


def detect_metric(df, metric):
    return find_col(df, METRICS.get(metric, [metric]))


def all_frames(source=None):
    sources = [source] if source else list(B.keys())
    for src in sources:
        for sheet, df in B.get(src, {}).items():
            if not df.empty:
                yield src, sheet, df


def best_sheet(source, metrics=None, dims=None, keywords=None):
    metrics = metrics or []
    dims = dims or []
    keywords = [norm(k) for k in (keywords or [])]
    best, best_score = None, -1
    for sheet, df in B.get(source, {}).items():
        score = 0
        for metric in metrics:
            if detect_metric(df, metric): score += 5
        for dim in dims:
            if detect_dim(df, dim): score += 2
        sheet_norm = norm(sheet)
        score += sum(4 for k in keywords if k and k in sheet_norm)
        score += min(len(df.columns), 30) * 0.03
        score += min(len(df), 50000) / 100000
        if score > best_score:
            best_score, best = score, sheet
    return best


def apply_filters(df, filters):
    x = clean(df)
    if x.empty:
        return x
    for dim, value in filters.items():
        if value in (None, "All"):
            continue
        c = detect_dim(x, dim)
        if c:
            mask = safe_series(x, c).astype(str).str.strip().str.casefold().eq(str(value).strip().casefold())
            x = x.loc[mask]
    return x


def dimension_values(dim, filters=None):
    filters = filters or {}
    vals = set()
    for _, _, df in all_frames():
        other = {k: v for k, v in filters.items() if k != dim}
        x = apply_filters(df, other)
        c = detect_dim(x, dim)
        if c:
            vals.update(unique_values(x, c))
    return sorted(vals, key=lambda x: x.casefold())


def source_data(source, metrics=None, dims=None, keywords=None, filters=None):
    sheet = best_sheet(source, metrics=metrics, dims=dims, keywords=keywords)
    if not sheet:
        return None, None
    return sheet, apply_filters(B[source][sheet], filters or {})


def sum_metric(source, metric, filters):
    # Prefer one best matching sheet to avoid double counting summary + raw sheets.
    sheet = best_sheet(source, metrics=[metric], dims=["region", "branch", "cluster", "city"])
    if not sheet:
        return np.nan
    x = apply_filters(B[source][sheet], filters)
    c = detect_metric(x, metric)
    if not c:
        return np.nan
    v = num(safe_series(x, c)).sum(min_count=1)
    return float(v) if pd.notna(v) else np.nan


def show_chart(fig, height=390):
    fig.update_layout(
        height=height,
        margin=dict(l=10, r=10, t=55, b=10),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="Inter", color="#344054"),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0),
        hovermode="x unified",
    )
    fig.update_xaxes(showgrid=True, gridcolor="#eef2f6")
    fig.update_yaxes(showgrid=False)
    st.plotly_chart(fig, use_container_width=True, config={"displaylogo": False})


def kpi(c, label, value, note=""):
    c.markdown(
        f'<div class="kpi"><div class="label">{label}</div>'
        f'<div class="value">{value}</div><div class="note">{note}</div></div>',
        unsafe_allow_html=True,
    )


def title(text, subtitle=None):
    st.markdown(f'<div class="section">{text}</div>', unsafe_allow_html=True)
    if subtitle:
        st.markdown(f'<div class="sub">{subtitle}</div>', unsafe_allow_html=True)


def insight(text):
    st.markdown(f'<div class="insight">💡 {text}</div>', unsafe_allow_html=True)


# -----------------------------
# Sidebar
# -----------------------------
with st.sidebar:
    st.markdown("## 📊 BALI NUSRA")
    st.caption("Business Command Center • V6")
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
    for key, filename in FILES.items():
        path = os.path.join(DATA_DIR, filename)
        sheets = B.get(key, {})
        if os.path.exists(path) and sheets:
            st.write(f"🟢 {key} · {len(sheets)} sheet")
        elif os.path.exists(path):
            st.write(f"🟡 {key} · file ada, sheet gagal dibaca")
        else:
            st.write(f"🔴 {key} · file tidak ada")

# -----------------------------
# Header + cascading filters
# -----------------------------
st.markdown(
    "<div class='hero'><h1>Bali Nusra Business Command Center</h1>"
    "<p>Executive performance • revenue • channel • digital adoption • competitive share • program effectiveness</p>"
    "<div class='badge'>LIVE EXCEL • SMART SHEET/COLUMN DETECTION • INTERACTIVE FILTERS</div></div>",
    unsafe_allow_html=True,
)

filters = {}
f1, f2, f3, f4 = st.columns(4)
with f1:
    filters["region"] = st.selectbox("Region", ["All"] + dimension_values("region"), key="filter_region")
with f2:
    filters["branch"] = st.selectbox("Branch", ["All"] + dimension_values("branch", {"region": filters["region"]}), key="filter_branch")
with f3:
    filters["cluster"] = st.selectbox("Cluster", ["All"] + dimension_values("cluster", {"region": filters["region"], "branch": filters["branch"]}), key="filter_cluster")
with f4:
    filters["city"] = st.selectbox("City / Kabupaten", ["All"] + dimension_values("city", {"region": filters["region"], "branch": filters["branch"], "cluster": filters["cluster"]}), key="filter_city")

# -----------------------------
# Executive Overview
# -----------------------------
if page == "🏠 Executive Overview":
    title("Executive Scorecard", "KPI lintas Excel dengan sheet dan nama kolom yang dicari otomatis.")
    revenue = sum_metric("Channel", "revenue", filters)
    rgb = sum_metric("Channel", "rgb", filters)
    halo = sum_metric("Halo", "halo", filters)
    active = sum_metric("Skul.id", "active", filters)
    productive = sum_metric("Skul.id", "productive", filters)
    so = sum_metric("Semarak", "so", filters)
    fr = sum_metric("Semarak", "fr", filters)

    cards = st.columns(7)
    kpi(cards[0], "REVENUE", money(revenue), "Channel")
    kpi(cards[1], "RGB", count(rgb), "Channel")
    kpi(cards[2], "HALO SUBS", count(halo), "Halo")
    kpi(cards[3], "SKUL ACTIVE", count(active), "Skul.id")
    kpi(cards[4], "SKUL PRODUCTIVE", count(productive), "Skul.id")
    kpi(cards[5], "FR", count(fr), "Program Semarak")
    kpi(cards[6], "FR RATE", pct(fr / so if pd.notna(fr) and pd.notna(so) and so else np.nan), f"{count(fr)} / {count(so)} SO")

    title("Business Pulse", "Gunakan filter di atas untuk drill-down sampai level kota.")
    left, right = st.columns([1.35, 1])
    with left:
        sheet, x = source_data("Channel", metrics=["revenue"], dims=["city"], filters=filters)
        if sheet is not None:
            rev = detect_metric(x, "revenue")
            city = detect_dim(x, "city")
            if rev and city:
                g = x.assign(_value=num(safe_series(x, rev))).groupby(city, dropna=False)["_value"].sum().reset_index().sort_values("_value", ascending=False).head(15)
                show_chart(px.bar(g, x="_value", y=city, orientation="h", text="_value", title="Revenue by City • Top 15", template="plotly_white"))
            else:
                st.info("Revenue atau City belum terdeteksi pada sheet Channel terpilih.")
    with right:
        sheet, x = source_data("FB Youth", dims=["region", "branch", "cluster", "city"], keywords=["raw", "youth", "share"], filters=filters)
        if sheet is not None:
            terr = find_col(x, ["TERITORI", "Territory", "territory_name", "city", "kota", "kabupaten"])
            ops = list(dict.fromkeys([c for c in [find_col(x, ["TSEL"]), find_col(x, ["IOH"]), find_col(x, ["XL", "XLSMART", "XLSmart"])] if c]))
            if terr and ops:
                g = x.groupby(terr)[ops].mean().reset_index()
                for c in ops:
                    g[c] = g[c].apply(pct_value)
                melted = g.melt(id_vars=[terr], var_name="Operator", value_name="Share")
                show_chart(px.bar(melted, x="Share", y=terr, color="Operator", orientation="h", barmode="group", title="Competitive Share", template="plotly_white"))
            else:
                st.info("Operator TSEL/IOH/XL atau territory belum terdeteksi pada FB Youth.")

    title("Management Signals")
    if pd.notna(revenue): insight(f"<b>Revenue</b> = {money(revenue)} pada filter aktif.")
    if pd.notna(active) and pd.notna(productive) and active: insight(f"<b>Skul productivity</b> = <b>{pct(productive/active)}</b> dari active users.")
    if pd.notna(so) and so: insight(f"<b>Semarak FR rate</b> = <b>{pct(fr/so)}</b>.")

# -----------------------------
# Channel Performance
# -----------------------------
elif page == "📈 Channel Performance":
    title("Channel Performance", "Revenue, RGB, outlet/channel mix dan trend.")
    tabs = st.tabs(["Revenue", "RGB", "Channel Mix", "Raw Data"])
    with tabs[0]:
        sheet, x = source_data("Channel", metrics=["revenue"], dims=["city", "date"], filters=filters)
        if sheet is not None:
            rev, city, date = detect_metric(x, "revenue"), detect_dim(x, "city"), detect_dim(x, "date")
            a,b,c = st.columns(3)
            kpi(a,"REVENUE",money(num(safe_series(x,rev)).sum()) if rev else "-",f"Sheet: {sheet}")
            kpi(b,"ROWS",f"{len(x):,}","filtered")
            kpi(c,"CITIES",f"{x[city].nunique():,}" if city else "-","unique")
            if rev and city:
                g=x.assign(_v=num(safe_series(x,rev))).groupby(city)['_v'].sum().reset_index().sort_values('_v',ascending=False).head(20)
                show_chart(px.bar(g,x='_v',y=city,orientation='h',title='Revenue by City',template='plotly_white'))
            if rev and date:
                z=x.copy(); z['_date']=pd.to_datetime(safe_series(z,date),errors='coerce'); z=z.dropna(subset=['_date'])
                if not z.empty:
                    z['_v']=num(safe_series(z,rev)); g=z.groupby(z['_date'].dt.to_period('M').astype(str))['_v'].sum().reset_index(name='Revenue')
                    show_chart(px.line(g,x='_date',y='Revenue',markers=True,title='Revenue Trend',template='plotly_white'))
    with tabs[1]:
        sheet, x = source_data("Channel", metrics=["rgb"], dims=["city", "date"], filters=filters)
        if sheet is not None:
            rgb, city, date = detect_metric(x,"rgb"), detect_dim(x,"city"), detect_dim(x,"date")
            kpi(st.columns(3)[0],"RGB",count(num(safe_series(x,rgb)).sum()) if rgb else "-",f"Sheet: {sheet}")
            if rgb and city:
                g=x.assign(_v=num(safe_series(x,rgb))).groupby(city)['_v'].sum().reset_index().sort_values('_v',ascending=False).head(20)
                show_chart(px.bar(g,x='_v',y=city,orientation='h',title='RGB by City',template='plotly_white'))
            if rgb and date:
                z=x.copy(); z['_date']=pd.to_datetime(safe_series(z,date),errors='coerce'); z=z.dropna(subset=['_date']); z['_v']=num(safe_series(z,rgb))
                if not z.empty:
                    g=z.groupby(z['_date'].dt.to_period('M').astype(str))['_v'].sum().reset_index(name='RGB')
                    show_chart(px.line(g,x='_date',y='RGB',markers=True,title='RGB Trend',template='plotly_white'))
    with tabs[2]:
        sheet, x = source_data("Halo", metrics=["halo"], dims=["city"], filters=filters)
        if sheet is not None:
            channel=find_col(x,["channel","Channel (standar dashboard)","channel standard"])
            if channel:
                g=x[channel].astype(str).value_counts().reset_index(); g.columns=['Channel','Count']
                show_chart(px.pie(g,names='Channel',values='Count',hole=.58,title='Activation Mix by Channel',template='plotly_white'))
            else: st.info("Kolom channel belum terdeteksi pada Halo.")
    with tabs[3]:
        sheet, x = source_data("Channel", metrics=["revenue"], dims=["city"], filters=filters)
        if sheet is not None:
            st.caption(f"Sheet terpilih otomatis: {sheet}")
            st.dataframe(x.head(1000),use_container_width=True,hide_index=True)
            st.download_button("⬇️ Download filtered CSV",x.to_csv(index=False).encode('utf-8'),'channel_filtered.csv','text/csv')

# -----------------------------
# Revenue & RGB
# -----------------------------
elif page == "💰 Revenue & RGB":
    title("Revenue & RGB", "Trend, concentration dan contribution.")
    sheet, x = source_data("Channel", metrics=["revenue", "rgb"], dims=["city", "date", "cluster"], filters=filters)
    if sheet is not None:
        rev, rgb, city, date, cluster = detect_metric(x,"revenue"), detect_metric(x,"rgb"), detect_dim(x,"city"), detect_dim(x,"date"), detect_dim(x,"cluster")
        a,b,c,d=st.columns(4)
        kpi(a,"REVENUE",money(num(safe_series(x,rev)).sum()) if rev else "-",f"{sheet}")
        kpi(b,"RGB",count(num(safe_series(x,rgb)).sum()) if rgb else "-","Channel")
        kpi(c,"TOP CITY",str(x.groupby(city)[rev].sum().idxmax()) if city and rev and not x.empty else "-","revenue")
        kpi(d,"ROWS",f"{len(x):,}","filtered")
        l,r=st.columns(2)
        if city and rev:
            with l:
                g=x.assign(_v=num(safe_series(x,rev))).groupby(city)['_v'].sum().reset_index().sort_values('_v',ascending=False).head(15)
                show_chart(px.bar(g,x='_v',y=city,orientation='h',title='Revenue Concentration',template='plotly_white'))
        if city and rgb:
            with r:
                g=x.assign(_v=num(safe_series(x,rgb))).groupby(city)['_v'].sum().reset_index().sort_values('_v',ascending=False).head(15)
                show_chart(px.bar(g,x='_v',y=city,orientation='h',title='RGB by City',template='plotly_white'))
        if date and rev:
            z=x.copy(); z['_date']=pd.to_datetime(safe_series(z,date),errors='coerce'); z=z.dropna(subset=['_date']); z['_v']=num(safe_series(z,rev))
            if not z.empty:
                g=z.groupby(z['_date'].dt.to_period('M').astype(str))['_v'].sum().reset_index(name='Revenue')
                show_chart(px.line(g,x='_date',y='Revenue',markers=True,title='Revenue Trend',template='plotly_white'))
    else: st.warning('Channel file/sheet tidak terbaca.')

# -----------------------------
# Halo & Device
# -----------------------------
elif page == "📱 Halo & Device":
    title("Halo & Device", "Activation, subscription, price, channel dan status.")
    sheet, x = source_data("Halo", metrics=["halo"], dims=["city", "date"], keywords=["halo", "raw"], filters=filters)
    if sheet is not None:
        subs=detect_metric(x,"halo"); price=detect_metric(x,"price"); city=detect_dim(x,"city"); date=detect_dim(x,"date")
        channel=find_col(x,["channel","Channel (standar dashboard)","channel standard"]); flag=find_col(x,["flag","status"])
        a,b,c,d=st.columns(4)
        kpi(a,"SUBSCRIPTIONS",count(num(safe_series(x,subs)).sum()) if subs else count(len(x)),f"{sheet}")
        kpi(b,"VALUE / PRICE",money(num(safe_series(x,price)).sum()) if price else "-","auto-detected")
        kpi(c,"CHANNELS",f"{x[channel].nunique():,}" if channel else "-","unique")
        kpi(d,"ROWS",f"{len(x):,}","filtered")
        l,r=st.columns(2)
        if city and subs:
            with l:
                g=x.assign(_v=num(safe_series(x,subs))).groupby(city)['_v'].sum().reset_index().sort_values('_v',ascending=False).head(20)
                show_chart(px.bar(g,x='_v',y=city,orientation='h',title='Halo Base by City',template='plotly_white'))
        if channel:
            with r:
                g=x[channel].astype(str).value_counts().reset_index();g.columns=['Channel','Count']
                show_chart(px.pie(g,names='Channel',values='Count',hole=.58,title='Activation by Channel',template='plotly_white'))
        st.dataframe(pd.DataFrame({'Metric':['Subscription','Price/Value','City','Date','Channel','Flag'],'Detected Column':[subs,price,city,date,channel,flag]}),use_container_width=True,hide_index=True)
    else: st.warning('File Halo belum terbaca.')

# -----------------------------
# Skul.id
# -----------------------------
elif page == "🎓 Skul.id":
    title("Skul.id — Digital Adoption Funnel", "Tidak lagi bergantung pada nama sheet 'Raw Skul.id'; sistem mencari sheet terbaik berdasarkan kolom yang benar-benar ada.")
    # Show every plausible sheet as a period selector when available.
    candidates=[]
    for s,df in B.get('Skul.id',{}).items():
        score=sum(bool(detect_metric(df,m)) for m in ['users','active','productive','school_id','student'])
        if score>=1: candidates.append((s,score,len(df)))
    candidates=sorted(candidates,key=lambda z:(z[1],z[2]),reverse=True)
    if candidates:
        labels=[z[0] for z in candidates]
        selected=st.selectbox('Sheet / Period Skul.id',labels,index=0,key='skul_sheet')
        x=apply_filters(B['Skul.id'][selected],filters)
        school_id=detect_metric(x,'school_id'); users=detect_metric(x,'users'); active=detect_metric(x,'active'); productive=detect_metric(x,'productive'); student=detect_metric(x,'student'); city=detect_dim(x,'city'); date=detect_dim(x,'date')
        schools=x[school_id].nunique() if school_id else (x[city].nunique() if city else 0)
        total_users=num(safe_series(x,users)).sum() if users else np.nan
        active_u=num(safe_series(x,active)).sum() if active else np.nan
        productive_u=num(safe_series(x,productive)).sum() if productive else np.nan
        student_u=num(safe_series(x,student)).sum() if student else np.nan
        cards=st.columns(5)
        kpi(cards[0],'SCHOOLS',count(schools),'unique school')
        kpi(cards[1],'STUDENTS',count(student_u),'detected')
        kpi(cards[2],'TOTAL USERS',count(total_users),'detected')
        kpi(cards[3],'ACTIVE USERS',count(active_u),'detected')
        kpi(cards[4],'PRODUCTIVE',count(productive_u),'detected')
        if pd.notna(total_users) and total_users:
            st.progress(min(max(active_u/total_users,0),1),text=f'Active conversion: {pct(active_u/total_users)}')
        if pd.notna(active_u) and active_u:
            st.progress(min(max(productive_u/active_u,0),1),text=f'Productive conversion: {pct(productive_u/active_u)}')
        l,r=st.columns(2)
        if city and active:
            with l:
                g=x.assign(_v=num(safe_series(x,active))).groupby(city)['_v'].sum().reset_index().sort_values('_v',ascending=False).head(20)
                show_chart(px.bar(g,x='_v',y=city,orientation='h',title='Active Users by City',template='plotly_white'))
        if city and active and productive:
            with r:
                g=x.assign(Active=num(safe_series(x,active)),Productive=num(safe_series(x,productive))).groupby(city)[['Active','Productive']].sum().reset_index()
                g['Productive Rate']=np.where(g['Active']>0,g['Productive']/g['Active'],0)
                show_chart(px.bar(g.sort_values('Productive Rate',ascending=False).head(20),x='Productive Rate',y=city,orientation='h',title='Productive Rate by City',template='plotly_white'))
        if date and active:
            z=x.copy();z['_date']=pd.to_datetime(safe_series(z,date),errors='coerce');z=z.dropna(subset=['_date']);z['_v']=num(safe_series(z,active))
            if not z.empty:
                g=z.groupby(z['_date'].dt.to_period('M').astype(str))['_v'].sum().reset_index(name='Active')
                show_chart(px.line(g,x='_date',y='Active',markers=True,title='Active Users Trend',template='plotly_white'))
        st.success(f'Sheet terdeteksi: {selected}')
        st.dataframe(pd.DataFrame({'Metric':['School ID','Student','Total Users','Active','Productive','City','Date'],'Detected Column':[school_id,student,users,active,productive,city,date]}),use_container_width=True,hide_index=True)
    else:
        st.warning('Tidak ada sheet Skul.id yang mempunyai kolom users/active/productive/school yang bisa dipetakan.')
        st.info('Buka Data Explorer untuk melihat nama sheet dan semua kolom yang benar-benar terbaca.')

# -----------------------------
# FB Youth
# -----------------------------
elif page == "⚔️ FB Youth":
    title("FB Youth — Competitive Intelligence", "TSEL vs IOH vs XL per territory/city dan filter wilayah.")
    sheet, x = source_data('FB Youth',dims=['region','branch','cluster','city'],keywords=['raw','fb','youth','share'],filters=filters)
    if sheet is not None:
        terr=find_col(x,['TERITORI','Territory','territory_name','city','kota','kabupaten'])
        ops=list(dict.fromkeys([c for c in [find_col(x,['TSEL']),find_col(x,['IOH']),find_col(x,['XL','XLSMART','XLSmart'])] if c]))
        if terr and ops:
            g=x.groupby(terr)[ops].mean().reset_index()
            for c in ops:g[c]=g[c].apply(pct_value)
            melt=g.melt(id_vars=[terr],var_name='Operator',value_name='Share')
            show_chart(px.bar(melt,x='Share',y=terr,color='Operator',orientation='h',barmode='group',title='Competitive Share by Territory',template='plotly_white'))
            st.dataframe(g,use_container_width=True,hide_index=True)
        else: st.warning('Kolom territory/TSEL/IOH/XL belum terdeteksi.')
        st.caption(f'Sheet terdeteksi: {sheet}')
    else: st.warning('File FB Youth belum terbaca.')

# -----------------------------
# Program Semarak
# -----------------------------
elif page == "🚀 Program Semarak":
    title("Program Semarak — SO to FR", "Funnel SO → FR, conversion rate, city dan trend.")
    sheet, x = source_data('Semarak',metrics=['so','fr'],dims=['region','branch','cluster','city','date'],keywords=['raw','semarak','program'],filters=filters)
    if sheet is not None:
        soc,frc,city,date=detect_metric(x,'so'),detect_metric(x,'fr'),detect_dim(x,'city'),detect_dim(x,'date')
        so_v=num(safe_series(x,soc)).sum() if soc else np.nan;fr_v=num(safe_series(x,frc)).sum() if frc else np.nan
        k=st.columns(4);kpi(k[0],'UNIQUE SO',count(so_v),sheet);kpi(k[1],'UNIQUE FR',count(fr_v),'detected');kpi(k[2],'FR RATE',pct(fr_v/so_v) if pd.notna(so_v) and so_v else '-','FR / SO');kpi(k[3],'ROWS',f'{len(x):,}','filtered')
        if city and soc:
            g=x.assign(_v=num(safe_series(x,soc))).groupby(city)['_v'].sum().reset_index().sort_values('_v',ascending=False).head(20)
            show_chart(px.bar(g,x='_v',y=city,orientation='h',title='SO by City',template='plotly_white'))
        if date and (soc or frc):
            z=x.copy();z['_date']=pd.to_datetime(safe_series(z,date),errors='coerce');z=z.dropna(subset=['_date'])
            if not z.empty:
                if soc:z['_so']=num(safe_series(z,soc))
                if frc:z['_fr']=num(safe_series(z,frc))
                g=z.groupby(z['_date'].dt.to_period('M').astype(str))[[c for c in ['_so','_fr'] if c in z.columns]].sum().reset_index()
                g=g.rename(columns={'_date':'Period','_so':'SO','_fr':'FR'})
                show_chart(px.line(g,x=g.columns[0],y=[c for c in ['SO','FR'] if c in g.columns],markers=True,title='SO vs FR Trend',template='plotly_white'))
        st.dataframe(pd.DataFrame({'Metric':['SO','FR','City','Date'],'Detected Column':[soc,frc,city,date]}),use_container_width=True,hide_index=True)
    else: st.warning('File Program Semarak belum terbaca.')

# -----------------------------
# Data Explorer / diagnostic
# -----------------------------
else:
    title('Data Explorer','Diagnostic lengkap: sheet, ukuran data, dimensi dan metric yang berhasil dideteksi.')
    rows=[]
    for src,sheet,df in all_frames():
        row={'Source':src,'Sheet':sheet,'Rows':len(df),'Columns':len(df.columns)}
        for d in DIM: row[d.title()]=detect_dim(df,d)
        for m in ['revenue','rgb','halo','price','users','active','productive','school_id','student','so','fr']: row[m.title()]=detect_metric(df,m)
        rows.append(row)
    if rows:
        diag=pd.DataFrame(rows)
        st.dataframe(diag,use_container_width=True,hide_index=True)
        st.download_button('⬇️ Download detection report',diag.to_csv(index=False).encode('utf-8'),'data_detection_report.csv','text/csv')
        st.divider()
        st.markdown('### Inspect one sheet')
        options=[f"{r['Source']} / {r['Sheet']}" for r in rows]
        pick=st.selectbox('Sheet',options)
        src,sh=pick.split(' / ',1);df=B[src][sh]
        st.write(f'**{len(df):,} rows × {len(df.columns):,} columns**')
        st.code('\n'.join(map(str,df.columns.tolist())))
        st.dataframe(df.head(200),use_container_width=True,hide_index=True)
    else:
        st.error('Tidak ada file Excel yang berhasil dibaca dari folder data/.')

st.divider()
st.caption(f"Bali Nusra Business Command Center V6 • Last refresh: {datetime.now().strftime('%d %b %Y %H:%M')} • Excel sources: data/")
