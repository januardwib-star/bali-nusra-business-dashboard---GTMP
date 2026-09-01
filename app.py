import os
import re
from pathlib import Path

import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px

# ============================================================
# BALI NUSRA BUSINESS COMMAND CENTER V3
# Robust version:
# - Handles duplicate Excel column names
# - Handles missing/changed columns
# - Does not use Series.str on a DataFrame
# - Reads all Excel files from data/
# ============================================================

st.set_page_config(
    page_title="Bali Nusra | Business Command Center",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

DATA_DIR = Path(__file__).parent / "data"

# ------------------------------------------------------------
# STYLE
# ------------------------------------------------------------
st.markdown(
    """
    <style>
    .block-container {padding-top: 1.2rem; padding-bottom: 2rem;}
    .section {
        background: linear-gradient(90deg,#eaf5ff,#ffffff);
        padding: 12px 18px;
        border-radius: 12px;
        margin: 12px 0;
        border-left: 5px solid #e31b23;
        font-weight: 700;
        font-size: 20px;
    }
    [data-testid="stMetricValue"] {font-size: 28px;}
    </style>
    """,
    unsafe_allow_html=True,
)

# ------------------------------------------------------------
# HELPERS
# ------------------------------------------------------------
def clean_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Make Excel headers safe and unique."""
    if df is None:
        return pd.DataFrame()

    out = df.copy()
    new_cols = []
    seen = {}

    for i, col in enumerate(out.columns):
        name = str(col).strip()
        if not name or name.lower() == "nan":
            name = f"column_{i+1}"

        # Remove line breaks / excessive spaces
        name = re.sub(r"\s+", " ", name).strip()

        # Make duplicate headers unique
        if name in seen:
            seen[name] += 1
            name = f"{name}_{seen[name]}"
        else:
            seen[name] = 0

        new_cols.append(name)

    out.columns = new_cols
    return out


def read_excel_file(path: Path) -> pd.DataFrame:
    """Read the first useful sheet from an Excel file."""
    try:
        book = pd.ExcelFile(path)
        best = None
        best_score = -1

        for sheet in book.sheet_names:
            try:
                df = pd.read_excel(path, sheet_name=sheet)
                df = clean_columns(df)

                # Prefer sheets with more rows and columns
                score = len(df) * max(len(df.columns), 1)
                if score > best_score:
                    best = df
                    best_score = score
            except Exception:
                continue

        if best is None:
            return pd.DataFrame()

        return best.dropna(how="all").reset_index(drop=True)

    except Exception:
        return pd.DataFrame()


@st.cache_data(ttl=300, show_spinner=False)
def load_all_data():
    files = {}
    if not DATA_DIR.exists():
        return files

    for path in sorted(DATA_DIR.glob("*.xlsx")):
        df = read_excel_file(path)
        files[path.stem] = {
            "path": str(path),
            "file": path.name,
            "df": df,
        }

    return files


def normalize_text(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def find_col(df, candidates):
    """Return the first matching column using exact or fuzzy matching."""
    if df is None or df.empty:
        return None

    cols = list(df.columns)
    low = {str(c).strip().lower(): c for c in cols}

    # Exact
    for candidate in candidates:
        key = str(candidate).strip().lower()
        if key in low:
            return low[key]

    # Fuzzy
    for candidate in candidates:
        key = str(candidate).strip().lower()
        for col in cols:
            c = str(col).strip().lower()
            if key in c or c in key:
                return col

    return None


def numeric_series(df, col):
    if col is None or col not in df.columns:
        return pd.Series(dtype="float64", index=df.index)

    s = df[col]

    # Extra protection if a source somehow still has duplicate labels
    if isinstance(s, pd.DataFrame):
        s = s.iloc[:, 0]

    if pd.api.types.is_numeric_dtype(s):
        return pd.to_numeric(s, errors="coerce").fillna(0)

    text = (
        s.astype("string")
        .str.replace(",", "", regex=False)
        .str.replace("%", "", regex=False)
        .str.replace("IDR", "", regex=False)
        .str.replace("Rp", "", regex=False)
        .str.strip()
    )

    return pd.to_numeric(text, errors="coerce").fillna(0)


def text_series(df, col):
    if col is None or col not in df.columns:
        return pd.Series("", index=df.index, dtype="string")

    s = df[col]

    # IMPORTANT: if duplicate headers make df[col] a DataFrame,
    # take only the first column before using .str.
    if isinstance(s, pd.DataFrame):
        s = s.iloc[:, 0]

    return s.astype("string").fillna("").str.strip()


def parse_date_col(df, candidates):
    col = find_col(df, candidates)
    if col is None:
        return None, pd.Series(pd.NaT, index=df.index)

    s = df[col]
    if isinstance(s, pd.DataFrame):
        s = s.iloc[:, 0]

    dates = pd.to_datetime(s, errors="coerce", dayfirst=True)
    return col, dates


def fmt_number(x):
    try:
        return f"{x:,.0f}".replace(",", ".")
    except Exception:
        return "0"


def fmt_compact(x):
    try:
        x = float(x)
    except Exception:
        return "0"

    if abs(x) >= 1_000_000_000:
        return f"{x/1_000_000_000:.1f} B"
    if abs(x) >= 1_000_000:
        return f"{x/1_000_000:.1f} M"
    if abs(x) >= 1_000:
        return f"{x/1_000:.1f} K"
    return f"{x:,.0f}"


def section(title):
    st.markdown(
        f'<div class="section">{title}</div>',
        unsafe_allow_html=True,
    )


def safe_chart(fig, height=420):
    fig.update_layout(
        template="plotly_white",
        margin=dict(l=20, r=20, t=55, b=20),
        height=height,
    )
    st.plotly_chart(fig, use_container_width=True)


# ------------------------------------------------------------
# LOAD DATA
# ------------------------------------------------------------
DATA = load_all_data()

# ------------------------------------------------------------
# SIDEBAR
# ------------------------------------------------------------
with st.sidebar:
    st.markdown("## 📊 BALI NUSRA")
    st.caption("Business Command Center • V3")

    if st.button("🔄 Refresh data", use_container_width=True):
        load_all_data.clear()
        st.rerun()

    st.divider()

    pages = [
        "Executive Overview",
        "Channel Performance",
        "Revenue & RGB",
        "Halo & Device",
        "Skul.id",
        "FB Youth",
        "Program Semarak",
        "Data Explorer",
    ]

    page = st.radio("MENU", pages)

    st.divider()
    st.caption("DATA UPDATE")
    st.caption("Replace Excel files inside `data/` and keep the same filenames.")
    st.caption("Then click **Refresh data**.")


# ------------------------------------------------------------
# EMPTY DATA CHECK
# ------------------------------------------------------------
if not DATA:
    st.error("Folder `data/` tidak ditemukan atau belum berisi file Excel.")
    st.info("Pastikan struktur repository seperti ini:\n\n"
            "app.py\nrequirements.txt\nREADME.md\nDEPLOYMENT.md\ndata/*.xlsx")
    st.stop()


# ------------------------------------------------------------
# DATASETS
# ------------------------------------------------------------
def get_df(name):
    item = DATA.get(name)
    if not item:
        return pd.DataFrame()
    return item["df"].copy()


channel = get_df("Channel Dashboard")
halo = get_df("Halo Dashboard")
semarak = get_df("Program Semarak")
skul = get_df("Skul.id")
fb = get_df("fb_youth_16082026")


# ============================================================
# EXECUTIVE OVERVIEW
# ============================================================
if page == "Executive Overview":
    section("Executive Overview")

    total_files = len(DATA)
    total_rows = sum(len(v["df"]) for v in DATA.values())

    # Find common business columns across all datasets
    revenue = 0
    users = 0
    so = 0

    for item in DATA.values():
        df = item["df"]

        rev_col = find_col(
            df,
            ["REVENUE", "REV", "REVENUE_TOTAL", "SALES", "VALUE", "AMOUNT"],
        )
        user_col = find_col(
            df,
            ["USER", "USERS", "ACTIVE USER", "ACTIVE_USERS", "RGB", "CUSTOMER"],
        )
        so_col = find_col(
            df,
            ["SO", "SALES ORDER", "SALES_ORDER", "ORDER", "TRANSACTION"],
        )

        if rev_col:
            revenue += numeric_series(df, rev_col).sum()
        if user_col:
            users += numeric_series(df, user_col).sum()
        if so_col:
            so += numeric_series(df, so_col).sum()

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Data Files", fmt_number(total_files))
    c2.metric("Total Rows", fmt_number(total_rows))
    c3.metric("Detected SO", fmt_compact(so))
    c4.metric("Detected Value", fmt_compact(revenue))

    st.divider()

    st.write("### Dataset Status")

    rows = []
    for name, item in DATA.items():
        df = item["df"]
        rows.append(
            {
                "Dataset": name,
                "File": item["file"],
                "Rows": len(df),
                "Columns": len(df.columns),
                "Status": "OK" if not df.empty else "EMPTY",
            }
        )

    st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

    st.info(
        "Dashboard ini membaca struktur kolom Excel secara otomatis. "
        "Jika nama kolom berbeda, data tetap ditampilkan di Data Explorer."
    )


# ============================================================
# CHANNEL PERFORMANCE
# ============================================================
elif page == "Channel Performance":
    section("Channel Performance")

    if channel.empty:
        st.warning("Channel Dashboard.xlsx tidak tersedia atau kosong.")
        st.stop()

    df = channel.copy()

    region_col = find_col(df, ["REGION", "REGIONAL", "AREA", "WITEL"])
    channel_col = find_col(
        df,
        ["CHANNEL", "CHANNEL_NAME", "CHANNEL NAME", "TYPE", "CATEGORY"],
    )
    date_col, dates = parse_date_col(
        df,
        ["TANGGAL", "DATE", "PERIOD", "MONTH", "BULAN"],
    )

    if date_col:
        df["_DATE_"] = dates

    # Filters
    f1, f2, f3 = st.columns(3)

    with f1:
        if region_col:
            regions = sorted(
                [x for x in text_series(df, region_col).unique() if x]
            )
            selected_region = st.selectbox("Region", ["ALL"] + regions)
            if selected_region != "ALL":
                df = df[text_series(df, region_col) == selected_region]

    with f2:
        if channel_col:
            channels = sorted(
                [x for x in text_series(df, channel_col).unique() if x]
            )
            selected_channel = st.selectbox("Channel", ["ALL"] + channels)
            if selected_channel != "ALL":
                df = df[text_series(df, channel_col) == selected_channel]

    with f3:
        if date_col and df["_DATE_"].notna().any():
            min_d = df["_DATE_"].min().date()
            max_d = df["_DATE_"].max().date()
            date_range = st.date_input(
                "Date Range",
                value=(min_d, max_d),
            )

            if isinstance(date_range, tuple) and len(date_range) == 2:
                start, end = date_range
                df = df[
                    (df["_DATE_"].dt.date >= start)
                    & (df["_DATE_"].dt.date <= end)
                ]

    value_col = find_col(
        df,
        ["SO", "SALES ORDER", "SALES_ORDER", "REVENUE", "SALES", "VALUE"],
    )

    if value_col:
        total_value = numeric_series(df, value_col).sum()
    else:
        total_value = len(df)

    c1, c2, c3 = st.columns(3)
    c1.metric("Records", fmt_number(len(df)))
    c2.metric("Detected Value", fmt_compact(total_value))
    c3.metric("Columns", fmt_number(len(df.columns)))

    if channel_col and value_col and not df.empty:
        g = (
            pd.DataFrame(
                {
                    "Channel": text_series(df, channel_col),
                    "Value": numeric_series(df, value_col),
                }
            )
            .groupby("Channel", as_index=False)["Value"]
            .sum()
            .sort_values("Value", ascending=False)
            .head(20)
        )

        if not g.empty:
            safe_chart(
                px.bar(
                    g,
                    x="Value",
                    y="Channel",
                    orientation="h",
                    text_auto=".2s",
                    title="Performance by Channel",
                )
            )

    st.dataframe(df.head(500), use_container_width=True, hide_index=True)


# ============================================================
# REVENUE & RGB
# ============================================================
elif page == "Revenue & RGB":
    section("Revenue & RGB")

    frames = []
    for name, item in DATA.items():
        df = item["df"].copy()

        rev_col = find_col(
            df,
            ["REVENUE", "REV", "REVENUE_TOTAL", "SALES", "VALUE", "AMOUNT"],
        )
        rgb_col = find_col(df, ["RGB", "ACTIVE USER", "ACTIVE_USERS", "USER"])
        date_col, dates = parse_date_col(
            df,
            ["TANGGAL", "DATE", "PERIOD", "MONTH", "BULAN"],
        )

        if rev_col or rgb_col:
            temp = pd.DataFrame(index=df.index)
            temp["Dataset"] = name
            temp["Revenue"] = numeric_series(df, rev_col) if rev_col else 0
            temp["RGB"] = numeric_series(df, rgb_col) if rgb_col else 0
            temp["Date"] = dates if date_col else pd.NaT
            frames.append(temp)

    if not frames:
        st.warning("Tidak ditemukan kolom Revenue / RGB pada data.")
        st.stop()

    all_rev = pd.concat(frames, ignore_index=True)

    c1, c2, c3 = st.columns(3)
    c1.metric("Revenue", fmt_compact(all_rev["Revenue"].sum()))
    c2.metric("RGB", fmt_compact(all_rev["RGB"].sum()))
    c3.metric("Datasets", fmt_number(all_rev["Dataset"].nunique()))

    if all_rev["Revenue"].sum() != 0:
        g = (
            all_rev.groupby("Dataset", as_index=False)["Revenue"]
            .sum()
            .sort_values("Revenue", ascending=False)
        )
        safe_chart(
            px.bar(
                g,
                x="Dataset",
                y="Revenue",
                text_auto=".2s",
                title="Revenue by Dataset",
            )
        )

    if all_rev["RGB"].sum() != 0:
        g = (
            all_rev.groupby("Dataset", as_index=False)["RGB"]
            .sum()
            .sort_values("RGB", ascending=False)
        )
        safe_chart(
            px.bar(
                g,
                x="Dataset",
                y="RGB",
                text_auto=".2s",
                title="RGB / Users by Dataset",
            )
        )


# ============================================================
# HALO & DEVICE
# ============================================================
elif page == "Halo & Device":
    section("Halo & Device")

    if halo.empty:
        st.warning("Halo Dashboard.xlsx tidak tersedia atau kosong.")
        st.stop()

    df = halo.copy()

    region_col = find_col(df, ["REGION", "REGIONAL", "AREA"])
    product_col = find_col(
        df,
        ["PRODUCT", "PRODUCT_NAME", "PRODUCT NAME", "DEVICE", "PACKAGE", "ITEM"],
    )
    value_col = find_col(
        df,
        ["SO", "SALES ORDER", "SALES_ORDER", "REVENUE", "VALUE", "AMOUNT"],
    )

    if region_col:
        regions = sorted([x for x in text_series(df, region_col).unique() if x])
        selected = st.selectbox("Region", ["ALL"] + regions)
        if selected != "ALL":
            df = df[text_series(df, region_col) == selected]

    total = numeric_series(df, value_col).sum() if value_col else len(df)

    c1, c2, c3 = st.columns(3)
    c1.metric("Records", fmt_number(len(df)))
    c2.metric("Value / SO", fmt_compact(total))
    c3.metric("Products", fmt_number(text_series(df, product_col).nunique() if product_col else 0))

    if product_col and value_col:
        g = (
            pd.DataFrame(
                {
                    "Product": text_series(df, product_col),
                    "Value": numeric_series(df, value_col),
                }
            )
            .groupby("Product", as_index=False)["Value"]
            .sum()
            .sort_values("Value", ascending=False)
            .head(20)
        )

        if not g.empty:
            safe_chart(
                px.bar(
                    g,
                    x="Value",
                    y="Product",
                    orientation="h",
                    text_auto=".2s",
                    title="Halo / Device Performance",
                )
            )

    st.dataframe(df.head(500), use_container_width=True, hide_index=True)


# ============================================================
# SKUL.ID
# ============================================================
elif page == "Skul.id":
    section("Skul.id")

    if skul.empty:
        st.warning("Skul.id.xlsx tidak tersedia atau kosong.")
        st.stop()

    df = skul.copy()

    region_col = find_col(df, ["REGION", "REGIONAL", "PROVINCE", "PROVINSI"])
    city_col = find_col(
        df,
        ["CITY", "KOTA", "KABUPATEN", "CITY_NAME", "KOTA/KABUPATEN"],
    )
    user_col = find_col(
        df,
        ["ACTIVE USER", "ACTIVE_USERS", "USER", "USERS", "TOTAL USER"],
    )
    date_col, dates = parse_date_col(
        df,
        ["TANGGAL", "DATE", "PERIOD", "MONTH", "BULAN"],
    )

    if region_col:
        regions = sorted([x for x in text_series(df, region_col).unique() if x])
        selected = st.selectbox("Region / Province", ["ALL"] + regions)
        if selected != "ALL":
            df = df[text_series(df, region_col) == selected]

    active_users = numeric_series(df, user_col).sum() if user_col else len(df)

    c1, c2, c3 = st.columns(3)
    c1.metric("Records", fmt_number(len(df)))
    c2.metric("Active Users", fmt_compact(active_users))
    c3.metric("Cities", fmt_number(text_series(df, city_col).nunique() if city_col else 0))

    if city_col and user_col:
        g = (
            pd.DataFrame(
                {
                    "City": text_series(df, city_col),
                    "Active Users": numeric_series(df, user_col),
                }
            )
            .groupby("City", as_index=False)["Active Users"]
            .sum()
            .sort_values("Active Users", ascending=False)
            .head(15)
        )

        if not g.empty:
            safe_chart(
                px.bar(
                    g,
                    x="Active Users",
                    y="City",
                    orientation="h",
                    text_auto=".2s",
                    title="Active Users by City",
                )
            )

    st.dataframe(df.head(500), use_container_width=True, hide_index=True)


# ============================================================
# FB YOUTH
# ============================================================
elif page == "FB Youth":
    section("FB Youth – Competitive Share")

    if fb.empty:
        st.warning("fb_youth_16082026.xlsx tidak tersedia atau kosong.")
        st.stop()

    df = fb.copy()

    territory_col = find_col(
        df,
        ["TERITORI", "TERRITORY", "REGION", "REGIONAL", "AREA"],
    )
    date_col, dates = parse_date_col(df, ["TANGGAL", "DATE", "PERIOD", "MONTH"])

    operator_candidates = []
    for c in ["TSEL", "TELKOMSEL", "IOH", "INDOSAT", "XL", "XLSMART", "SMARTFREN"]:
        col = find_col(df, [c])
        if col:
            operator_candidates.append(col)

    if territory_col:
        territories = sorted(
            [x for x in text_series(df, territory_col).unique() if x]
        )
        selected = st.selectbox("Territory", ["ALL"] + territories)
        if selected != "ALL":
            df = df[text_series(df, territory_col) == selected]

    if operator_candidates and territory_col:
        temp = pd.DataFrame()
        temp["Territory"] = text_series(df, territory_col)

        for col in operator_candidates:
            temp[str(col)] = numeric_series(df, col)

        g = temp.groupby("Territory", as_index=False).mean(numeric_only=True)

        melted = g.melt(
            id_vars=["Territory"],
            var_name="Operator",
            value_name="Share",
        )

        safe_chart(
            px.bar(
                melted,
                x="Share",
                y="Territory",
                color="Operator",
                orientation="h",
                barmode="group",
                title="Competitive Share by Territory",
            )
        )

    st.dataframe(df.head(500), use_container_width=True, hide_index=True)


# ============================================================
# PROGRAM SEMARAK
# ============================================================
elif page == "Program Semarak":
    section("Program Semarak – SO to FR")

    if semarak.empty:
        st.warning("Program Semarak.xlsx tidak tersedia atau kosong.")
        st.stop()

    df = semarak.copy()

    region_col = find_col(df, ["REGION", "REGIONAL", "AREA"])
    date_col, dates = parse_date_col(df, ["TANGGAL", "DATE", "PERIOD", "MONTH"])
    so_col = find_col(df, ["SO", "SALES ORDER", "SALES_ORDER", "ORDER"])
    fr_col = find_col(df, ["FR", "FULFILLMENT", "REDEEM", "ACTIVATION"])

    if region_col:
        regions = sorted([x for x in text_series(df, region_col).unique() if x])
        selected = st.selectbox("Region", ["ALL"] + regions)
        if selected != "ALL":
            df = df[text_series(df, region_col) == selected]

    so_total = numeric_series(df, so_col).sum() if so_col else 0
    fr_total = numeric_series(df, fr_col).sum() if fr_col else 0
    conversion = (fr_total / so_total * 100) if so_total else 0

    c1, c2, c3 = st.columns(3)
    c1.metric("SO", fmt_compact(so_total))
    c2.metric("FR", fmt_compact(fr_total))
    c3.metric("SO → FR", f"{conversion:.1f}%")

    if so_col or fr_col:
        chart_df = pd.DataFrame(
            {
                "Metric": ["SO", "FR"],
                "Value": [so_total, fr_total],
            }
        )
        safe_chart(
            px.bar(
                chart_df,
                x="Metric",
                y="Value",
                text_auto=".2s",
                title="Program Semarak – SO vs FR",
            )
        )

    st.dataframe(df.head(500), use_container_width=True, hide_index=True)


# ============================================================
# DATA EXPLORER
# ============================================================
elif page == "Data Explorer":
    section("Data Explorer")

    names = list(DATA.keys())
    selected_name = st.selectbox("Select dataset", names)

    df = DATA[selected_name]["df"].copy()

    st.write(
        f"**File:** `{DATA[selected_name]['file']}`  |  "
        f"**Rows:** {len(df):,}  |  **Columns:** {len(df.columns):,}"
    )

    # Safe column search
    search = st.text_input("Search columns", "")
    columns = list(df.columns)

    if search:
        columns = [
            c for c in columns
            if search.lower() in str(c).lower()
        ]

    st.write("Columns:")
    st.code(", ".join(map(str, columns)))

    # Preview
    st.dataframe(
        df[columns].head(1000) if columns else df.head(1000),
        use_container_width=True,
        hide_index=True,
    )

    st.download_button(
        "⬇️ Download current preview as CSV",
        data=df.to_csv(index=False).encode("utf-8-sig"),
        file_name=f"{selected_name}.csv",
        mime="text/csv",
    )
