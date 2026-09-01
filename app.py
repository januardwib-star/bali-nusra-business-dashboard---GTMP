import os
import pandas as pd
import numpy as np
import streamlit as st
import plotly.express as px

st.set_page_config(
    page_title="Bali Nusra | Business Command Center",
    page_icon="📊",
    layout="wide",
)

DATA = os.path.join(os.path.dirname(__file__), "data")

FILES = {
    "Channel": "Channel Dashboard.xlsx",
    "Halo": "Halo Dashboard.xlsx",
    "Skul.id": "Skul.id.xlsx",
    "FB Youth": "fb_youth_16082026.xlsx",
    "Semarak": "Program Semarak.xlsx",
}

st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
    html, body, [class*="css"] {font-family: Inter, sans-serif;}
    .block-container {max-width:1500px; padding-top:1rem; padding-bottom:2rem;}
    .hero {
        background:linear-gradient(120deg,#111827,#1f2937 65%,#374151);
        color:white; border-radius:24px; padding:24px 28px; margin-bottom:18px;
    }
    .hero h1 {margin:0; font-size:30px; letter-spacing:-.04em;}
    .hero p {margin:6px 0 0; color:#D0D5DD; font-size:13px;}
    .kpi {
        background:#fff; border:1px solid #EAECF0; border-radius:18px;
        padding:15px 17px; min-height:112px;
        box-shadow:0 2px 12px rgba(16,24,40,.045);
    }
    .kpi .l {font-size:10px;color:#667085;font-weight:800;letter-spacing:.10em;text-transform:uppercase;}
    .kpi .v {font-size:26px;color:#101828;font-weight:800;margin-top:7px;}
    .kpi .s {font-size:11px;color:#667085;margin-top:4px;}
    .section {font-size:20px;font-weight:800;color:#101828;margin:22px 0 10px;}
    div[data-testid="stMetric"] {
        background:#fff;border:1px solid #EAECF0;padding:12px 14px;border-radius:16px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

def clean(df):
    if df is None or df.empty:
        return pd.DataFrame()
    x = df.copy()
    x.columns = [str(c).strip() for c in x.columns]
    return x.dropna(how="all")

def num(s):
    if pd.api.types.is_numeric_dtype(s):
        return pd.to_numeric(s, errors="coerce")
    return pd.to_numeric(
        s.astype(str)
        .str.replace(",", "", regex=False)
        .str.replace("%", "", regex=False)
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

def find_col(df, names):
    if df is None or df.empty:
        return None
    cols = list(df.columns)
    lower = {str(c).lower().strip(): c for c in cols}

    for name in names:
        key = str(name).lower().strip()
        if key in lower:
            return lower[key]

    for name in names:
        key = str(name).lower().strip()
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
    return {
        sheet: pd.read_excel(path, sheet_name=sheet)
        for sheet in pd.ExcelFile(path).sheet_names
    }

@st.cache_data(show_spinner=False)
def load_all():
    result = {}
    for source, filename in FILES.items():
        path = os.path.join(DATA, filename)
        if os.path.exists(path):
            result[source] = read_book(path)
    return result

B = load_all()

def raw(source, sheet):
    return promote_header(B[source][sheet])

def source_rows(source):
    return [(sheet, raw(source, sheet)) for sheet in B.get(source, {})]

def values(terms):
    vals = set()

    for source in B:
        for _, df0 in source_rows(source):
            df = clean(df0)
            if df.empty:
                continue

            c = find_col(df, terms)
            if c is None:
                continue

            series = df[c].dropna().astype(str).str.strip()
            for value in series:
                if value and value.lower() not in {"nan", "none"}:
                    vals.add(value)

    return sorted(vals)

def apply_filters(df):
    x = clean(df)

    filters = [
        (region, ["regional", "region"]),
        (branch, ["branch", "branch_lacci"]),
        (cluster, ["cluster", "cluster_name", "cluster_lacci", "cluster_sales"]),
        (city, ["city", "kota", "kabupaten", "city_lacci"]),
    ]

    for selected, terms in filters:
        if selected != "All":
            c = find_col(x, terms)
            if c:
                x = x[x[c].astype(str).str.strip() == selected]

    return x

def kpi(container, label, value, subtitle=""):
    container.markdown(
        f"""
        <div class="kpi">
            <div class="l">{label}</div>
            <div class="v">{value}</div>
            <div class="s">{subtitle}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

# Sidebar
with st.sidebar:
    st.markdown("## 📊 BALI NUSRA")
    st.caption("Business Command Center • V3")

    if st.button("🔄 Refresh data", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

    st.divider()

    page = st.radio(
        "MENU",
        [
            "Executive Overview",
            "Channel Performance",
            "Revenue & RGB",
            "Halo & Device",
            "Skul.id",
            "FB Youth",
            "Program Semarak",
            "Data Explorer",
        ],
    )

    st.divider()
    st.caption("DATA UPDATE")
    st.caption("Replace Excel in `data/` → keep filename → Refresh.")

# Global filters
f1, f2, f3, f4 = st.columns(4)

region = f1.selectbox(
    "Region",
    ["All"] + values(["regional", "region"])[:300],
)

branch = f2.selectbox(
    "Branch",
    ["All"] + values(["branch", "branch_lacci"])[:300],
)

cluster = f3.selectbox(
    "Cluster",
    ["All"] + values(
        ["cluster", "cluster_name", "cluster_lacci", "cluster_sales"]
    )[:400],
)

city = f4.selectbox(
    "City / Kabupaten",
    ["All"] + values(
        ["city", "kota", "kabupaten", "city_lacci"]
    )[:500],
)

st.markdown(
    """
    <div class="hero">
        <h1>Bali Nusra Business Command Center</h1>
        <p>Executive performance • channel productivity • revenue • digital adoption • competitive share</p>
    </div>
    """,
    unsafe_allow_html=True,
)

# Executive Overview
if page == "Executive Overview":
    st.markdown(
        '<div class="section">Executive Scorecard</div>',
        unsafe_allow_html=True,
    )

    cards = st.columns(6)

    revenue = 0
    revenue_source = 0

    if "Channel" in B:
        for sheet in B["Channel"]:
            x = apply_filters(raw("Channel", sheet))
            c = find_col(
                x,
                ["REV_ALL_BB", "REV_ALL", "Total Omzet", "Total_Omzet"],
            )
            if c:
                revenue += num(x[c]).sum()
                revenue_source += 1

    kpi(
        cards[0],
        "Revenue",
        money(revenue) if revenue_source else "-",
        "Revenue / broadband source",
    )

    rgb = 0
    if "Channel" in B and "RGB" in B["Channel"]:
        x = apply_filters(raw("Channel", "RGB"))
        c = find_col(x, ["rgb_all", "rgb"])
        if c:
            rgb = num(x[c]).sum()

    kpi(cards[1], "RGB", count(rgb), "Registered / base metric")

    halo = 0
    if "Halo" in B and "Rev Halo All" in B["Halo"]:
        x = apply_filters(raw("Halo", "Rev Halo All"))
        c = find_col(x, ["subs"])
        if c:
            halo = num(x[c]).sum()

    kpi(cards[2], "Halo Subs", count(halo), "Activation dataset")

    active = 0
    productive = 0
    schools = 0

    if "Skul.id" in B:
        raw_sheets = [
            raw("Skul.id", sheet)
            for sheet in B["Skul.id"]
            if "raw data" in sheet.lower()
        ]

        if raw_sheets:
            x = apply_filters(raw_sheets[-1])
            ca = find_col(x, ["User Active"])
            cp = find_col(x, ["User Productive"])
            ci = find_col(x, ["ID Sekolah"])

            if ca:
                active = num(x[ca]).sum()
            if cp:
                productive = num(x[cp]).sum()
            if ci:
                schools = x[ci].nunique()

    kpi(cards[3], "Skul Active", count(active), "Latest raw period")
    kpi(cards[4], "Skul Productive", count(productive), "Latest raw period")
    kpi(cards[5], "Schools", count(schools), "Unique school IDs")

    st.markdown(
        '<div class="section">Management View</div>',
        unsafe_allow_html=True,
    )

    left, right = st.columns([1.35, 1])

    with left:
        if "Channel" in B and "RevAll_BB" in B["Channel"]:
            x = apply_filters(raw("Channel", "RevAll_BB"))
            city_col = find_col(x, ["CITY"])
            money_col = find_col(x, ["REV_ALL_BB", "REV_ALL"])

            if city_col and money_col:
                g = (
                    x.groupby(city_col, dropna=False)[money_col]
                    .sum()
                    .reset_index()
                    .sort_values(money_col, ascending=False)
                    .head(15)
                )

                st.plotly_chart(
                    px.bar(
                        g,
                        x=money_col,
                        y=city_col,
                        orientation="h",
                        text=money_col,
                        template="plotly_white",
                        title="Revenue by City — Top 15",
                    ),
                    use_container_width=True,
                )

    with right:
        if "FB Youth" in B and "FB YOUTH" in B["FB Youth"]:
            x = apply_filters(raw("FB Youth", "FB YOUTH"))
            territory = find_col(x, ["TERITORI"])
            tsel = find_col(x, ["TSEL"])
            ioh = find_col(x, ["IOH"])
            xl = find_col(x, ["XL"])

            if territory and tsel:
                cols = [tsel]
                if ioh:
                    cols.append(ioh)
                if xl:
                    cols.append(xl)

                g = (
                    x.groupby(territory)[cols]
                    .mean()
                    .reset_index()
                )

                melted = g.melt(
                    id_vars=[territory],
                    var_name="Operator",
                    value_name="Share",
                )

                st.plotly_chart(
                    px.bar(
                        melted,
                        x="Share",
                        y=territory,
                        color="Operator",
                        orientation="h",
                        template="plotly_white",
                        title="FB Youth — Avg Share by Territory",
                    ),
                    use_container_width=True,
                )

# Channel Performance
elif page == "Channel Performance":
    st.markdown(
        '<div class="section">Channel Performance</div>',
        unsafe_allow_html=True,
    )

    tabs = st.tabs(["Revenue", "RGB", "Outlet"])

    with tabs[0]:
        if "Channel" in B and "RevAll_BB" in B["Channel"]:
            x = apply_filters(raw("Channel", "RevAll_BB"))
            m = find_col(x, ["REV_ALL_BB"])
            city_col = find_col(x, ["CITY"])
            date_col = find_col(x, ["TGL"])

            if m:
                st.metric("Revenue BB", money(num(x[m]).sum()))

            if city_col and m:
                g = (
                    x.groupby(city_col)[m]
                    .sum()
                    .reset_index()
                    .sort_values(m, ascending=False)
                    .head(15)
                )

                st.plotly_chart(
                    px.bar(
                        g,
                        x=m,
                        y=city_col,
                        orientation="h",
                        text=m,
                        template="plotly_white",
                        title="Revenue by City",
                    ),
                    use_container_width=True,
                )

            if date_col and m:
                x[date_col] = pd.to_datetime(
                    x[date_col], errors="coerce"
                )
                g = x.groupby(date_col)[m].sum().reset_index()

                st.plotly_chart(
                    px.line(
                        g,
                        x=date_col,
                        y=m,
                        markers=True,
                        template="plotly_white",
                        title="Revenue Trend",
                    ),
                    use_container_width=True,
                )

    with tabs[1]:
        if "Channel" in B and "RGB" in B["Channel"]:
            x = apply_filters(raw("Channel", "RGB"))
            m = find_col(x, ["rgb_all", "rgb"])
            period = find_col(x, ["period"])
            kab = find_col(x, ["kabupaten"])

            if m:
                st.metric("RGB", count(num(x[m]).sum()))

            if kab and m:
                g = (
                    x.groupby(kab)[m]
                    .sum()
                    .reset_index()
                    .sort_values(m, ascending=False)
                    .head(15)
                )

                st.plotly_chart(
                    px.bar(
                        g,
                        x=m,
                        y=kab,
                        orientation="h",
                        text=m,
                        template="plotly_white",
                        title="RGB by Kabupaten",
                    ),
                    use_container_width=True,
                )

            if period and m:
                x[period] = pd.to_datetime(
                    x[period], errors="coerce"
                )
                g = x.groupby(period)[m].sum().reset_index()

                st.plotly_chart(
                    px.line(
                        g,
                        x=period,
                        y=m,
                        markers=True,
                        template="plotly_white",
                        title="RGB Trend",
                    ),
                    use_container_width=True,
                )

    with tabs[2]:
        if "Channel" in B and "Omzet Outlet" in B["Channel"]:
            x = apply_filters(raw("Channel", "Omzet Outlet"))
            m = find_col(x, ["Total Omzet"])
            group_col = find_col(x, ["cluster"])

            if m:
                st.metric(
                    "Total Outlet Omzet",
                    money(num(x[m]).sum()),
                )

                if group_col:
                    g = (
                        x.groupby(group_col)[m]
                        .sum()
                        .reset_index()
                        .sort_values(m, ascending=False)
                    )

                    st.plotly_chart(
                        px.bar(
                            g,
                            x=group_col,
                            y=m,
                            text=m,
                            template="plotly_white",
                            title="Outlet Revenue by Cluster",
                        ),
                        use_container_width=True,
                    )

                st.dataframe(
                    x.head(1000),
                    use_container_width=True,
                )

# Revenue & RGB
elif page == "Revenue & RGB":
    st.markdown(
        '<div class="section">Revenue & RGB</div>',
        unsafe_allow_html=True,
    )

    left, right = st.columns(2)

    if "Channel" in B and "RevAll_BB" in B["Channel"]:
        x = apply_filters(raw("Channel", "RevAll_BB"))
        m = find_col(x, ["REV_ALL_BB"])
        date_col = find_col(x, ["TGL"])

        with left:
            st.subheader("Revenue BB")

            if m:
                st.metric("Revenue", money(num(x[m]).sum()))

            if date_col and m:
                x[date_col] = pd.to_datetime(
                    x[date_col], errors="coerce"
                )
                g = x.groupby(date_col)[m].sum().reset_index()

                st.plotly_chart(
                    px.line(
                        g,
                        x=date_col,
                        y=m,
                        markers=True,
                        template="plotly_white",
                    ),
                    use_container_width=True,
                )

    if "Channel" in B and "RGB" in B["Channel"]:
        x = apply_filters(raw("Channel", "RGB"))
        m = find_col(x, ["rgb_all", "rgb"])
        brand = find_col(x, ["brand"])

        with right:
            st.subheader("RGB")

            if m:
                st.metric("RGB", count(num(x[m]).sum()))

            if brand and m:
                g = (
                    x.groupby(brand)[m]
                    .sum()
                    .reset_index()
                    .sort_values(m, ascending=False)
                )

                st.plotly_chart(
                    px.pie(
                        g,
                        names=brand,
                        values=m,
                        hole=.55,
                        template="plotly_white",
                    ),
                    use_container_width=True,
                )

# Halo & Device
elif page == "Halo & Device":
    st.markdown(
        '<div class="section">Halo & Device</div>',
        unsafe_allow_html=True,
    )

    if "Halo" in B and "Rev Halo All" in B["Halo"]:
        x = apply_filters(raw("Halo", "Rev Halo All"))

        subs = find_col(x, ["subs"])
        price = find_col(x, ["price"])
        channel = find_col(
            x,
            ["channel", "Channel (standar dashboard)"],
        )
        flag = find_col(x, ["flag"])

        a, b, c, d = st.columns(4)

        a.metric(
            "Subscriptions",
            count(num(x[subs]).sum()) if subs else "-",
        )

        b.metric(
            "Revenue / Price",
            money(num(x[price]).sum()) if price else "-",
        )

        if channel:
            g = (
                x.groupby(channel)
                .size()
                .reset_index(name="Activations")
                .sort_values("Activations", ascending=False)
                .head(12)
            )

            c.plotly_chart(
                px.bar(
                    g,
                    x="Activations",
                    y=channel,
                    orientation="h",
                    template="plotly_white",
                ),
                use_container_width=True,
            )

        if flag:
            g = (
                x.groupby(flag)
                .size()
                .reset_index(name="Count")
            )

            d.plotly_chart(
                px.pie(
                    g,
                    names=flag,
                    values="Count",
                    hole=.55,
                    template="plotly_white",
                ),
                use_container_width=True,
            )

        st.dataframe(
            x.head(1000),
            use_container_width=True,
        )

# Skul.id
elif page == "Skul.id":
    st.markdown(
        '<div class="section">Skul.id — School & User Funnel</div>',
        unsafe_allow_html=True,
    )

    if "Skul.id" in B:
        sheets = [
            sheet
            for sheet in B["Skul.id"]
            if "raw data" in sheet.lower()
        ]

        if sheets:
            pick = st.selectbox(
                "Period",
                sheets,
                index=len(sheets) - 1,
            )

            x = apply_filters(raw("Skul.id", pick))

            school = find_col(x, ["ID Sekolah"])
            users = find_col(x, ["Total Users"])
            active = find_col(x, ["User Active"])
            productive = find_col(x, ["User Productive"])

            a, b, c, d = st.columns(4)

            a.metric(
                "Schools",
                count(x[school].nunique()) if school else "-",
            )
            b.metric(
                "Total Users",
                count(num(x[users]).sum()) if users else "-",
            )
            c.metric(
                "Active Users",
                count(num(x[active]).sum()) if active else "-",
            )
            d.metric(
                "Productive Users",
                count(num(x[productive]).sum())
                if productive
                else "-",
            )

            if active and productive:
                funnel = pd.DataFrame(
                    {
                        "Stage": ["Active", "Productive"],
                        "Users": [
                            num(x[active]).sum(),
                            num(x[productive]).sum(),
                        ],
                    }
                )

                st.plotly_chart(
                    px.funnel(
                        funnel,
                        x="Users",
                        y="Stage",
                        template="plotly_white",
                        title="User Funnel",
                    ),
                    use_container_width=True,
                )

            city_col = find_col(x, ["Kota", "City"])

            if city_col and active:
                g = (
                    x.groupby(city_col)[active]
                    .sum()
                    .reset_index()
                    .sort_values(active, ascending=False)
                    .head(15)
                )

                st.plotly_chart(
                    px.bar(
                        g,
                        x=active,
                        y=city_col,
                        orientation="h",
                        text=active,
                        template="plotly_white",
                        title="Active Users by City",
                    ),
                    use_container_width=True,
                )
        else:
            st.warning("Raw Skul.id sheet tidak ditemukan.")

# FB Youth
elif page == "FB Youth":
    st.markdown(
        '<div class="section">FB Youth — Competitive Share</div>',
        unsafe_allow_html=True,
    )

    if "FB Youth" in B and "FB YOUTH" in B["FB Youth"]:
        x = apply_filters(raw("FB Youth", "FB YOUTH"))

        territory = find_col(x, ["TERITORI"])
        date_col = find_col(x, ["TANGGAL"])

        operators = [
            c
            for c in [
                find_col(x, ["TSEL"]),
                find_col(x, ["IOH"]),
                find_col(x, ["XL"]),
            ]
            if c
        ]

        if date_col:
            x[date_col] = pd.to_datetime(
                x[date_col], errors="coerce"
            )

        if territory and operators:
            g = (
                x.groupby(territory)[operators]
                .mean()
                .reset_index()
            )

            melted = g.melt(
                id_vars=[territory],
                var_name="Operator",
                value_name="Share",
            )

            st.plotly_chart(
                px.bar(
                    melted,
                    x="Share",
                    y=territory,
                    color="Operator",
                    orientation="h",
                    barmode="group",
                    template="plotly_white",
                    title="Average Share by Territory",
                ),
                use_container_width=True,
            )

        if date_col and operators:
            g = (
                x.groupby(date_col)[operators]
                .mean()
                .reset_index()
            )

            melted = g.melt(
                id_vars=[date_col],
                var_name="Operator",
                value_name="Share",
            )

            st.plotly_chart(
                px.line(
                    melted,
                    x=date_col,
                    y="Share",
                    color="Operator",
                    markers=True,
                    template="plotly_white",
                    title="Share Trend",
                ),
                use_container_width=True,
            )

        st.dataframe(
            x.head(1500),
            use_container_width=True,
        )

# Program Semarak
elif page == "Program Semarak":
    st.markdown(
        '<div class="section">Program Semarak — SO to FR</div>',
        unsafe_allow_html=True,
    )

    x = pd.DataFrame()

    if "Semarak" in B:
        candidates = []

        for sheet in B["Semarak"]:
            candidate = raw("Semarak", sheet)
            required = {
                "cluster_name",
                "branch",
                "status_fr",
            }

            if required.issubset(
                {str(c).lower() for c in candidate.columns}
            ):
                candidates.append(candidate)

        if candidates:
            x = candidates[0].copy()

            mapping = {
                str(c).lower(): c
                for c in x.columns
            }

            rename = {}

            for target in [
                "cluster_name",
                "branch",
                "status_fr",
                "channel",
                "period",
                "created_at",
                "code_semarak",
                "uniq_SO",
                "uniq_FR",
                "price",
            ]:
                if target in mapping:
                    rename[mapping[target]] = target

            x = x.rename(columns=rename)

            if "period" in x:
                x["period"] = pd.to_datetime(
                    x["period"], errors="coerce"
                )

            if "uniq_SO" in x:
                x["uniq_SO"] = num(x["uniq_SO"]).fillna(0)

            if "uniq_FR" in x:
                x["uniq_FR"] = num(x["uniq_FR"]).fillna(0)

    if not x.empty:
        x = apply_filters(x)

        a, b, c, d = st.columns(4)

        so = (
            num(x["uniq_SO"]).sum()
            if "uniq_SO" in x
            else 0
        )

        fr = (
            num(x["uniq_FR"]).sum()
            if "uniq_FR" in x
            else 0
        )

        rate = fr / so if so else np.nan

        a.metric("SO Unique", count(so))
        b.metric("FR Sukses", count(fr))
        c.metric("FR Rate", pct(rate))

        if "status_fr" in x:
            d.metric("Rows", count(len(x)))

        if "cluster_name" in x:
            g = (
                x.groupby("cluster_name")
                .agg(
                    SO=("uniq_SO", "sum"),
                    FR=("uniq_FR", "sum"),
                )
                .reset_index()
            )

            g["FR Rate"] = np.where(
                g["SO"] > 0,
                g["FR"] / g["SO"],
                0,
            )

            st.plotly_chart(
                px.bar(
                    g.sort_values("SO", ascending=False),
                    x="cluster_name",
                    y=["SO", "FR"],
                    barmode="group",
                    template="plotly_white",
                    title="SO vs FR by Cluster",
                ),
                use_container_width=True,
            )

            st.dataframe(
                g.sort_values(
                    "FR Rate",
                    ascending=False,
                ),
                use_container_width=True,
            )

        if "branch" in x and "uniq_FR" in x:
            g = (
                x.groupby("branch")["uniq_FR"]
                .sum()
                .reset_index()
                .sort_values(
                    "uniq_FR",
                    ascending=False,
                )
                .head(15)
            )

            st.plotly_chart(
                px.bar(
                    g,
                    x="uniq_FR",
                    y="branch",
                    orientation="h",
                    text="uniq_FR",
                    template="plotly_white",
                    title="Top Branch — FR Sukses",
                ),
                use_container_width=True,
            )
    else:
        st.warning("Raw Semarak belum bisa dipetakan.")

# Data Explorer
else:
    st.markdown(
        '<div class="section">Data Explorer</div>',
        unsafe_allow_html=True,
    )

    if B:
        source = st.selectbox(
            "Source",
            list(B.keys()),
        )

        sheet = st.selectbox(
            "Sheet",
            list(B[source].keys()),
        )

        x = apply_filters(raw(source, sheet))

        st.caption(
            f"{len(x):,} rows × {len(x.columns):,} columns"
        )

        st.dataframe(
            x,
            use_container_width=True,
            height=680,
        )
    else:
        st.error(
            "Tidak ada file Excel di folder data/. "
            "Pastikan semua filename sesuai."
        )

st.caption(
    "Bali Nusra Business Command Center • "
    "Data source: Excel files in /data"
)
