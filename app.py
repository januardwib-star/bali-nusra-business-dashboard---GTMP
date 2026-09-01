
import os
import re
import pandas as pd
import numpy as np
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go

st.set_page_config(page_title="Bali Nusra | Business Command Center",page_icon="📊",layout="wide")

DATA=os.path.join(os.path.dirname(__file__),"data")
FILES={
 "Channel":"Channel Dashboard.xlsx",
 "Halo":"Halo Dashboard.xlsx",
 "Skul.id":"Skul.id.xlsx",
 "FB Youth":"fb_youth_16082026.xlsx",
 "Semarak":"Program Semarak.xlsx"
}

# ---------- Styling ----------
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
html,body,[class*="css"]{font-family:Inter,sans-serif}
.block-container{max-width:1500px;padding-top:1rem;padding-bottom:2rem}
.hero{background:linear-gradient(120deg,#111827,#1f2937 65%,#374151);color:white;border-radius:24px;padding:24px 28px;margin-bottom:18px}
.hero h1{margin:0;font-size:30px;letter-spacing:-.04em}.hero p{margin:6px 0 0;color:#D0D5DD;font-size:13px}
.kpi{background:#fff;border:1px solid #EAECF0;border-radius:18px;padding:15px 17px;min-height:112px;box-shadow:0 2px 12px rgba(16,24,40,.045)}
.kpi .l{font-size:10px;color:#667085;font-weight:800;letter-spacing:.10em;text-transform:uppercase}.kpi .v{font-size:26px;color:#101828;font-weight:800;margin-top:7px}.kpi .s{font-size:11px;color:#667085;margin-top:4px}
.section{font-size:20px;font-weight:800;color:#101828;margin:22px 0 10px}
.small{font-size:11px;color:#667085}.good{color:#067647;font-weight:800}.bad{color:#B42318;font-weight:800}
div[data-testid="stMetric"]{background:#fff;border:1px solid #EAECF0;padding:12px 14px;border-radius:16px}
</style>
""",unsafe_allow_html=True)

# ---------- Generic helpers ----------
def clean(x):
    if x is None or x.empty:return pd.DataFrame()
    x=x.copy()
    x.columns=[str(c).strip() for c in x.columns]
    return x.dropna(how="all")

def n(x):
    if pd.api.types.is_numeric_dtype(x): return pd.to_numeric(x,errors="coerce")
    return pd.to_numeric(x.astype(str).str.replace(",","",regex=False).str.replace("%","",regex=False),errors="coerce")

def money(v):
    if v is None or pd.isna(v):return "-"
    v=float(v)
    if abs(v)>=1e9:return f"Rp {v/1e9:.2f} B"
    if abs(v)>=1e6:return f"Rp {v/1e6:.2f} M"
    if abs(v)>=1e3:return f"Rp {v/1e3:.1f} K"
    return f"Rp {v:,.0f}"

def count(v):
    if v is None or pd.isna(v):return "-"
    v=float(v)
    if abs(v)>=1e6:return f"{v/1e6:.2f} M"
    if abs(v)>=1e3:return f"{v/1e3:.1f} K"
    return f"{v:,.0f}"

def pct(v):
    if v is None or pd.isna(v):return "-"
    return f"{float(v)*100:.1f}%"

def col(df, names):
    low={str(c).lower():c for c in df.columns}
    for x in names:
        if x.lower() in low:return low[x.lower()]
    for x in names:
        for c in df.columns:
            if x.lower() in str(c).lower():return c
    return None

def numeric(df):
    return [c for c in df.columns if n(df[c]).notna().sum()>=max(3,int(len(df)*.2))]

def promote_header(df):
    """Excel exports often have a title row followed by real headers."""
    x=clean(df)
    if x.empty:return x
    first=x.iloc[0].astype(str).tolist()
    nonempty=sum(v not in ("nan","None","") for v in first)
    unnamed=sum(str(c).lower().startswith("unnamed") for c in x.columns)
    if unnamed and nonempty>=max(3,int(len(x.columns)*.25)):
        headers=[]
        for i,v in enumerate(first):
            v=str(v).strip()
            headers.append(v if v and v.lower()!="nan" else f"col_{i}")
        y=x.iloc[1:].copy();y.columns=headers
        return y.reset_index(drop=True)
    return x

@st.cache_data(show_spinner=False)
def book(path):
    return {s:pd.read_excel(path,sheet_name=s) for s in pd.ExcelFile(path).sheet_names}

@st.cache_data(show_spinner=False)
def allbooks():
    return {k:book(os.path.join(DATA,fn)) for k,fn in FILES.items() if os.path.exists(os.path.join(DATA,fn))}

B=allbooks()

def raw(src,sheet):
    return promote_header(B[src][sheet])

def source_rows(src):
    return [(s,raw(src,s)) for s in B.get(src,{})]

# ---------- Semarak clean model ----------
@st.cache_data(show_spinner=False)
def semarak_model():
    if "Semarak" not in B:return pd.DataFrame()
    candidates=[]
    for s in B["Semarak"]:
        x=raw("Semarak",s)
        if {"cluster_name","branch","status_fr"}.issubset({c.lower() for c in x.columns}):
            candidates.append(x)
    if not candidates:return pd.DataFrame()
    x=candidates[0]
    # case-insensitive canonicalization
    mp={str(c).lower():c for c in x.columns}
    rename={}
    for target in ["cluster_name","branch","status_fr","channel","period","created_at","code_semarak","uniq_SO","uniq_FR","price"]:
        if target in mp:rename[mp[target]]=target
    x=x.rename(columns=rename)
    if "period" in x:x["period"]=pd.to_datetime(x["period"],errors="coerce")
    if "uniq_SO" in x:x["uniq_SO"]=n(x["uniq_SO"]).fillna(0)
    if "uniq_FR" in x:x["uniq_FR"]=n(x["uniq_FR"]).fillna(0)
    return x

# ---------- Sidebar ----------
with st.sidebar:
    st.markdown("## 📊 BALI NUSRA")
    st.caption("Business Command Center • V3")
    if st.button("🔄 Refresh data",use_container_width=True):
        st.cache_data.clear();st.rerun()
    st.divider()
    page=st.radio("MENU",[
        "Executive Overview","Channel Performance","Revenue & RGB",
        "Halo & Device","Skul.id","FB Youth","Program Semarak","Data Explorer"
    ])
    st.divider()
    st.caption("DATA UPDATE")
    st.caption("Replace Excel in `data/` → keep filename → Refresh.")
    st.caption("Last app refresh is generated dynamically when data is loaded.")

# ---------- Global filters ----------
def values(terms):
    vals=set()
    for src in B:
        for _,df0 in source_rows(src):
            df=df0
            c=col(df,terms)
      if c:
    data = df[c]
    if isinstance(data, pd.DataFrame):
        data = data.iloc[:, 0]
    vals.update(data.dropna().astype(str).str.strip())
    return sorted(v for v in vals if v and v.lower() not in {"nan","none"})

f1,f2,f3,f4=st.columns(4)
region=f1.selectbox("Region",["All"]+values(["regional","region"])[:300])
branch=f2.selectbox("Branch",["All"]+values(["branch","branch_lacci"])[:300])
cluster=f3.selectbox("Cluster",["All"]+values(["cluster","cluster_name","cluster_lacci","cluster_sales"])[:400])
city=f4.selectbox("City / Kabupaten",["All"]+values(["city","kota","kabupaten","city_lacci"])[:500])

def apply_filters(df):
    x=clean(df)
    for v,terms in [(region,["regional","region"]),(branch,["branch","branch_lacci"]),
                    (cluster,["cluster","cluster_name","cluster_lacci","cluster_sales"]),
                    (city,["city","kota","kabupaten","city_lacci"])]:
        if v!="All":
            c=col(x,terms)
            if c:x=x[x[c].astype(str).str.strip()==v]
    return x

st.markdown("""<div class="hero">
<h1>Bali Nusra Business Command Center</h1>
<p>Executive performance • channel productivity • revenue • digital adoption • competitive share</p>
</div>""",unsafe_allow_html=True)

def kpi(c,label,value,sub=""):
    c.markdown(f'<div class="kpi"><div class="l">{label}</div><div class="v">{value}</div><div class="s">{sub}</div></div>',unsafe_allow_html=True)

# ---------- Executive ----------
if page=="Executive Overview":
    st.markdown('<div class="section">Executive Scorecard</div>',unsafe_allow_html=True)
    cs=st.columns(6)

    # Revenue
    rev=0; rev_src=0
    if "Channel" in B:
        for s in B["Channel"]:
            x=apply_filters(raw("Channel",s)); c=col(x,["REV_ALL_BB","REV_ALL","Total Omzet","Total_Omzet"])
            if c:
                rev+=n(x[c]).sum();rev_src+=1
    kpi(cs[0],"Revenue",money(rev) if rev_src else "-","Revenue / broadband source")

    # RGB
    rgb=0
    if "Channel" in B and "RGB" in B["Channel"]:
        x=apply_filters(raw("Channel","RGB"));c=col(x,["rgb_all","rgb"])
        if c:rgb=n(x[c]).sum()
    kpi(cs[1],"RGB",count(rgb),"Registered / base metric")

    # Halo subs
    halo=0
    if "Halo" in B and "Rev Halo All" in B["Halo"]:
        x=apply_filters(raw("Halo","Rev Halo All"));c=col(x,["subs"])
        if c:halo=n(x[c]).sum()
    kpi(cs[2],"Halo Subs",count(halo),"Activation dataset")

    # Skul active users
    active=0; productive=0; schools=0
    if "Skul.id" in B:
        raw_sheets=[raw("Skul.id",s) for s in B["Skul.id"] if "raw data" in s.lower()]
        if raw_sheets:
            x=apply_filters(raw_sheets[-1])
            ca=col(x,["User Active"]);cp=col(x,["User Productive"]);ci=col(x,["ID Sekolah"])
            if ca:active=n(x[ca]).sum()
            if cp:productive=n(x[cp]).sum()
            if ci:schools=x[ci].nunique()
    kpi(cs[3],"Skul Active",count(active),"Latest raw period")
    kpi(cs[4],"Skul Productive",count(productive),"Latest raw period")
    kpi(cs[5],"Schools",count(schools),"Unique school IDs")

    st.markdown('<div class="section">Management View</div>',unsafe_allow_html=True)
    left,right=st.columns([1.35,1])
    with left:
        # Revenue by city
        if "Channel" in B and "RevAll_BB" in B["Channel"]:
            x=apply_filters(raw("Channel","RevAll_BB")); cc=col(x,["CITY"]); mc=col(x,["REV_ALL_BB","REV_ALL"])
            if cc and mc:
                g=x.groupby(cc,dropna=False)[mc].sum().reset_index().sort_values(mc,ascending=False).head(15)
                st.plotly_chart(px.bar(g,x=mc,y=cc,orientation="h",text=mc,template="plotly_white",title="Revenue by City — Top 15"),use_container_width=True)
    with right:
        if "FB Youth" in B and "FB YOUTH" in B["FB Youth"]:
            x=apply_filters(raw("FB Youth","FB YOUTH"))
            cterrit=col(x,["TERITORI"]);ctsel=col(x,["TSEL"]);cioh=col(x,["IOH"]);cxl=col(x,["XL"])
            if cterrit and ctsel:
                g=x.groupby(cterrit)[[ctsel]+([cioh] if cioh else [])+([cxl] if cxl else [])].mean().reset_index()
                melted=g.melt(id_vars=[cterrit],var_name="Operator",value_name="Share")
                st.plotly_chart(px.bar(melted,x="Share",y=cterrit,color="Operator",orientation="h",template="plotly_white",title="FB Youth — Avg Share by Territory"),use_container_width=True)

elif page=="Channel Performance":
    st.markdown('<div class="section">Channel Performance</div>',unsafe_allow_html=True)
    tabs=st.tabs(["Revenue","RGB","Outlet"])
    with tabs[0]:
        if "Channel" in B and "RevAll_BB" in B["Channel"]:
            x=apply_filters(raw("Channel","RevAll_BB")); m=col(x,["REV_ALL_BB"]); cityc=col(x,["CITY"]); datec=col(x,["TGL"])
            a,b=st.columns(2)
            if m:a.metric("Revenue BB",money(n(x[m]).sum()))
            if cityc and m:
                g=x.groupby(cityc)[m].sum().reset_index().sort_values(m,ascending=False).head(15)
                st.plotly_chart(px.bar(g,x=m,y=cityc,orientation="h",text=m,template="plotly_white"),use_container_width=True)
            if datec and m:
                x[datec]=pd.to_datetime(x[datec],errors="coerce");g=x.groupby(datec)[m].sum().reset_index()
                st.plotly_chart(px.line(g,x=datec,y=m,markers=True,template="plotly_white",title="Revenue Trend"),use_container_width=True)
    with tabs[1]:
        if "Channel" in B and "RGB" in B["Channel"]:
            x=apply_filters(raw("Channel","RGB"));m=col(x,["rgb_all","rgb"]);d=col(x,["period"]);c=col(x,["kabupaten"])
            if m:
                st.metric("RGB",count(n(x[m]).sum()))
                if c:
                    g=x.groupby(c)[m].sum().reset_index().sort_values(m,ascending=False).head(15)
                    st.plotly_chart(px.bar(g,x=m,y=c,orientation="h",text=m,template="plotly_white",title="RGB by Kabupaten"),use_container_width=True)
                if d:
                    x[d]=pd.to_datetime(x[d],errors="coerce");g=x.groupby(d)[m].sum().reset_index()
                    st.plotly_chart(px.line(g,x=d,y=m,markers=True,template="plotly_white",title="RGB Trend"),use_container_width=True)
    with tabs[2]:
        if "Channel" in B and "Omzet Outlet" in B["Channel"]:
            x=apply_filters(raw("Channel","Omzet Outlet"));m=col(x,["Total Omzet"]);g1=col(x,["cluster"]);g2=col(x,["fisik"])
            if m:
                st.metric("Total Outlet Omzet",money(n(x[m]).sum()))
                if g1:
                    g=x.groupby(g1)[m].sum().reset_index().sort_values(m,ascending=False)
                    st.plotly_chart(px.bar(g,x=g1,y=m,text=m,template="plotly_white",title="Outlet Revenue by Cluster"),use_container_width=True)
                st.dataframe(x.head(1000),use_container_width=True)

elif page=="Revenue & RGB":
    st.markdown('<div class="section">Revenue & RGB</div>',unsafe_allow_html=True)
    a,b=st.columns(2)
    if "Channel" in B and "RevAll_BB" in B["Channel"]:
        x=apply_filters(raw("Channel","RevAll_BB"));m=col(x,["REV_ALL_BB"]);d=col(x,["TGL"])
        with a:
            st.subheader("Revenue BB")
            if m:st.metric("Revenue",money(n(x[m]).sum()))
            if d and m:
                x[d]=pd.to_datetime(x[d],errors="coerce");g=x.groupby(d)[m].sum().reset_index()
                st.plotly_chart(px.line(g,x=d,y=m,markers=True,template="plotly_white"),use_container_width=True)
    if "Channel" in B and "RGB" in B["Channel"]:
        x=apply_filters(raw("Channel","RGB"));m=col(x,["rgb_all"]);c=col(x,["brand"])
        with b:
            st.subheader("RGB")
            if m:st.metric("RGB",count(n(x[m]).sum()))
            if c and m:
                g=x.groupby(c)[m].sum().reset_index().sort_values(m,ascending=False)
                st.plotly_chart(px.pie(g,names=c,values=m,hole=.55,template="plotly_white"),use_container_width=True)

elif page=="Halo & Device":
    st.markdown('<div class="section">Halo & Device</div>',unsafe_allow_html=True)
    if "Halo" in B:
        x=apply_filters(raw("Halo","Rev Halo All"))
        a,b,c,d=st.columns(4)
        subs=col(x,["subs"]);price=col(x,["price"]);ch=col(x,["channel","Channel (standar dashboard)"]);flag=col(x,["flag"])
        a.metric("Subscriptions",count(n(x[subs]).sum()) if subs else "-")
        b.metric("Revenue / Price",money(n(x[price]).sum()) if price else "-")
        if ch:
            g=x.groupby(ch).size().reset_index(name="Activations").sort_values("Activations",ascending=False).head(12)
            c.plotly_chart(px.bar(g,x="Activations",y=ch,orientation="h",template="plotly_white"),use_container_width=True)
        if flag:
            g=x.groupby(flag).size().reset_index(name="Count")
            d.plotly_chart(px.pie(g,names=flag,values="Count",hole=.55,template="plotly_white"),use_container_width=True)
        st.dataframe(x.head(1000),use_container_width=True)

elif page=="Skul.id":
    st.markdown('<div class="section">Skul.id — School & User Funnel</div>',unsafe_allow_html=True)
    raws=[(s,raw("Skul.id",s)) for s in B.get("Skul.id",{}) if "raw data" in s.lower()]
    if raws:
        labels=[s for s,_ in raws]; pick=st.selectbox("Period",labels,index=len(labels)-1)
        x=apply_filters(dict(raws)[pick])
        ci=col(x,["ID Sekolah"]);tu=col(x,["Total Users"]);ua=col(x,["User Active"]);up=col(x,["User Productive"])
        a,b,c,d=st.columns(4)
        a.metric("Schools",count(x[ci].nunique()) if ci else "-")
        b.metric("Total Users",count(n(x[tu]).sum()) if tu else "-")
        c.metric("Active Users",count(n(x[ua]).sum()) if ua else "-")
        d.metric("Productive Users",count(n(x[up]).sum()) if up else "-")
        if ua and up:
            x["_active"]=n(x[ua]);x["_productive"]=n(x[up])
            g=pd.DataFrame({"Stage":["Active","Productive"],"Users":[x["_active"].sum(),x["_productive"].sum()]})
            st.plotly_chart(px.funnel(g,x="Users",y="Stage",template="plotly_white",title="User Funnel"),use_container_width=True)
        cityc=col(x,["Kota","City"])
        if cityc and ua:
            g=x.groupby(cityc)[ua].sum().reset_index().sort_values(ua,ascending=False).head(15)
            st.plotly_chart(px.bar(g,x=ua,y=cityc,orientation="h",text=ua,template="plotly_white",title="Active Users by City"),use_container_width=True)
    else:
        st.warning("Raw Skul.id sheet tidak ditemukan.")

elif page=="FB Youth":
    st.markdown('<div class="section">FB Youth — Competitive Share</div>',unsafe_allow_html=True)
    if "FB Youth" in B and "FB YOUTH" in B["FB Youth"]:
        x=apply_filters(raw("FB Youth","FB YOUTH"))
        terr=col(x,["TERITORI"]);dt=col(x,["TANGGAL"]);ops=[z for z in [col(x,["TSEL"]),col(x,["IOH"]),col(x,["XL"])] if z]
        if dt:x[dt]=pd.to_datetime(x[dt],errors="coerce")
        if terr and ops:
            g=x.groupby(terr)[ops].mean().reset_index()
            melted=g.melt(id_vars=[terr],var_name="Operator",value_name="Share")
            st.plotly_chart(px.bar(melted,x="Share",y=terr,color="Operator",orientation="h",barmode="group",template="plotly_white",title="Average Share by Territory"),use_container_width=True)
        if dt and ops:
            g=x.groupby(dt)[ops].mean().reset_index()
            melted=g.melt(id_vars=[dt],var_name="Operator",value_name="Share")
            st.plotly_chart(px.line(melted,x=dt,y="Share",color="Operator",markers=True,template="plotly_white",title="Share Trend"),use_container_width=True)
        st.dataframe(x.head(1500),use_container_width=True)

elif page=="Program Semarak":
    st.markdown('<div class="section">Program Semarak — SO to FR</div>',unsafe_allow_html=True)
    x=semarak_model()
    if not x.empty:
        x=apply_filters(x)
        a,b,c,d=st.columns(4)
        so=n(x["uniq_SO"]).sum() if "uniq_SO" in x else 0
        fr=n(x["uniq_FR"]).sum() if "uniq_FR" in x else 0
        rate=fr/so if so else np.nan
        a.metric("SO Unique",count(so));b.metric("FR Sukses",count(fr));c.metric("FR Rate",pct(rate))
        if "status_fr" in x:d.metric("Rows",count(len(x)))
        if "cluster_name" in x:
            g=x.groupby("cluster_name").agg(SO=("uniq_SO","sum"),FR=("uniq_FR","sum")).reset_index()
            g["FR Rate"]=np.where(g.SO>0,g.FR/g.SO,0)
            st.plotly_chart(px.bar(g.sort_values("SO",ascending=False),x="cluster_name",y=["SO","FR"],barmode="group",template="plotly_white",title="SO vs FR by Cluster"),use_container_width=True)
            st.dataframe(g.sort_values("FR Rate",ascending=False),use_container_width=True)
        if "branch" in x and "uniq_FR" in x:
            g=x.groupby("branch")["uniq_FR"].sum().reset_index().sort_values("uniq_FR",ascending=False).head(15)
            st.plotly_chart(px.bar(g,x="uniq_FR",y="branch",orientation="h",text="uniq_FR",template="plotly_white",title="Top Branch — FR Sukses"),use_container_width=True)
    else: st.warning("Raw Semarak belum bisa dipetakan.")

else:
    st.markdown('<div class="section">Data Explorer</div>',unsafe_allow_html=True)
    src=st.selectbox("Source",list(B.keys()))
    sh=st.selectbox("Sheet",list(B[src].keys()))
    x=apply_filters(raw(src,sh))
    st.caption(f"{len(x):,} rows × {len(x.columns):,} columns")
    st.dataframe(x,use_container_width=True,height=680)

st.caption("Bali Nusra Business Command Center • Data source: Excel files in /data • Replace files and use Refresh data to update.")
