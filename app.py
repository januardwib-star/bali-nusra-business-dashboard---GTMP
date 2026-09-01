import os
from datetime import datetime
import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go

st.set_page_config(page_title="Bali Nusra | Business Command Center", page_icon="📊", layout="wide", initial_sidebar_state="expanded")
DATA_DIR = os.path.join(os.path.dirname(__file__), "data")
FILES = {"Channel":"Channel Dashboard.xlsx","Halo":"Halo Dashboard.xlsx","Skul.id":"Skul.id.xlsx","FB Youth":"fb_youth_16082026.xlsx","Semarak":"Program Semarak.xlsx"}

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
html,body,[class*="css"]{font-family:Inter,sans-serif}.block-container{max-width:1600px;padding-top:1rem}
[data-testid="stSidebar"]{border-right:1px solid #e5e7eb}
.hero{background:linear-gradient(135deg,#111827 0%,#172554 48%,#0f766e 100%);color:#fff;border-radius:24px;padding:28px 32px;margin-bottom:18px;box-shadow:0 14px 40px rgba(15,23,42,.14)}
.hero h1{font-size:32px;font-weight:800;margin:0;letter-spacing:-.04em}.hero p{margin:7px 0 0;color:#dbeafe;font-size:13px}
.badge{display:inline-block;margin-top:13px;background:rgba(255,255,255,.13);padding:6px 11px;border-radius:999px;font-size:11px;font-weight:700}
.kpi{background:#fff;border:1px solid #e5e7eb;border-radius:18px;padding:15px 17px;min-height:116px;box-shadow:0 6px 22px rgba(15,23,42,.05)}
.kpi .label{font-size:10px;font-weight:800;letter-spacing:.1em;color:#667085;text-transform:uppercase}.kpi .value{font-size:25px;font-weight:800;color:#111827;margin-top:7px}.kpi .note{font-size:11px;color:#667085;margin-top:4px}
.section{font-size:20px;font-weight:800;color:#101828;margin:22px 0 8px}.sub{font-size:12px;color:#667085;margin-bottom:12px}
.insight{background:#f8fafc;border:1px solid #e5e7eb;border-radius:15px;padding:13px 15px;margin:7px 0}.good{color:#067647;font-weight:800}.warn{color:#b54708;font-weight:800}.bad{color:#b42318;font-weight:800}
</style>
""", unsafe_allow_html=True)

def clean(df):
    if df is None or not isinstance(df,pd.DataFrame) or df.empty:return pd.DataFrame()
    x=df.copy();seen={};new=[]
    for i,c in enumerate(x.columns):
        base=str(c).strip()
        if not base or base.lower()=="nan":base=f"Column_{i+1}"
        n=seen.get(base,0)+1;seen[base]=n;new.append(base if n==1 else f"{base}__{n}")
    x.columns=new
    return x.dropna(how="all").reset_index(drop=True)

def safe_series(df,col):
    if df is None or col is None or col not in df.columns:return pd.Series(index=df.index if isinstance(df,pd.DataFrame) else [],dtype="object")
    s=df[col]
    return s.iloc[:,0] if isinstance(s,pd.DataFrame) else s

def num(s):
    if isinstance(s,pd.DataFrame):s=s.iloc[:,0]
    if s is None:return pd.Series(dtype=float)
    if pd.api.types.is_numeric_dtype(s):return pd.to_numeric(s,errors="coerce")
    return pd.to_numeric(s.astype(str).str.replace(",","",regex=False).str.replace("%","",regex=False).str.replace("Rp","",regex=False).str.replace(" ","",regex=False),errors="coerce")

def pct_raw(v):
    try:
        v=float(v);return v/100 if abs(v)>1.5 else v
    except:return np.nan

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

def pct(v):return "-" if v is None or pd.isna(v) else f"{float(v)*100:.1f}%"

def find_col(df,names):
    if df is None or df.empty:return None
    cols=list(df.columns);exact={str(c).strip().lower():c for c in cols}
    for name in names:
        if str(name).lower() in exact:return exact[str(name).lower()]
    for name in names:
        key=str(name).lower()
        for c in cols:
            if key in str(c).lower():return c
    return None

def unique_values(df,col):
    if df.empty or col is None or col not in df.columns:return []
    s=safe_series(df,col).dropna().astype(str).str.strip()
    return sorted([v for v in s.unique() if v and v.lower() not in {"nan","none","null"}])

@st.cache_data(show_spinner=False)
def read_book(path):
    out={}
    if not os.path.exists(path):return out
    try:
        xl=pd.ExcelFile(path)
        for sheet in xl.sheet_names:
            try:
                x=clean(pd.read_excel(path,sheet_name=sheet))
                if x.empty:continue
                unnamed=sum(str(c).lower().startswith("unnamed") for c in x.columns)
                if unnamed and len(x)>0:
                    first=x.iloc[0].astype(str).tolist()
                    if sum(v not in {"nan","none",""} for v in first)>=max(3,len(first)//4):
                        headers=[v.strip() if v.strip().lower() not in {"nan","none",""} else f"Column_{i+1}" for i,v in enumerate(first)]
                        y=x.iloc[1:].copy();y.columns=headers;x=clean(y)
                out[sheet]=x
            except Exception:continue
    except Exception:return {}
    return out

@st.cache_data(show_spinner=False)
def load_all():
    return {key:read_book(os.path.join(DATA_DIR,fn)) for key,fn in FILES.items()}

B=load_all()

def all_frames():
    for src,sheets in B.items():
        for sheet,df in sheets.items():
            if not df.empty:yield src,sheet,df

def all_values(terms):
    vals=set()
    for _,_,df in all_frames():
        c=find_col(df,terms)
        if c:vals.update(unique_values(df,c))
    return sorted(vals)

def apply_filters(df,region="All",branch="All",cluster="All",city="All"):
    x=clean(df)
    rules=[(region,["regional","region"]),(branch,["branch","branch_lacci"]),(cluster,["cluster","cluster_name","cluster_lacci","cluster_sales"]),(city,["city","kota","kabupaten","city_lacci"])]
    for value,terms in rules:
        if value!="All":
            c=find_col(x,terms)
            if c:x=x[safe_series(x,c).astype(str).str.strip().eq(value)]
    return x

def sheet_with(src,required):
    if src not in B:return None
    for s,df in B[src].items():
        if all(find_col(df,terms) for terms in required):return s
    return next(iter(B[src]),None) if B[src] else None

def chart(fig,height=380):
    fig.update_layout(height=height,margin=dict(l=10,r=10,t=50,b=10),paper_bgcolor="rgba(0,0,0,0)",plot_bgcolor="rgba(0,0,0,0)",font=dict(family="Inter",color="#344054"),legend=dict(orientation="h",yanchor="bottom",y=1.02,x=0),hovermode="x unified")
    fig.update_xaxes(showgrid=True,gridcolor="#eef2f6");fig.update_yaxes(showgrid=False)
    st.plotly_chart(fig,use_container_width=True,config={"displaylogo":False})

def kpi(c,label,value,note=""):
    c.markdown(f'<div class="kpi"><div class="label">{label}</div><div class="value">{value}</div><div class="note">{note}</div></div>',unsafe_allow_html=True)

def title(t,s=None):
    st.markdown(f'<div class="section">{t}</div>',unsafe_allow_html=True)
    if s:st.markdown(f'<div class="sub">{s}</div>',unsafe_allow_html=True)

def insight(t):st.markdown(f'<div class="insight">💡 {t}</div>',unsafe_allow_html=True)

def metric_from_source(src,terms,region,branch,cluster,city):
    for _,df in B.get(src,{}).items():
        c=find_col(df,terms)
        if c:
            x=apply_filters(df,region,branch,cluster,city);return num(safe_series(x,c)).sum()
    return 0

with st.sidebar:
    st.markdown("## 📊 BALI NUSRA");st.caption("Business Command Center • V5")
    if st.button("🔄 Refresh data",use_container_width=True):st.cache_data.clear();st.rerun()
    st.divider()
    page=st.radio("NAVIGATION",["🏠 Executive Overview","📈 Channel Performance","💰 Revenue & RGB","📱 Halo & Device","🎓 Skul.id","⚔️ FB Youth","🚀 Program Semarak","🗂️ Data Explorer"])
    st.divider();st.caption("DATA SOURCES")
    for key,fn in FILES.items():st.write(("🟢 " if os.path.exists(os.path.join(DATA_DIR,fn)) else "🔴 ")+key)

st.markdown("<div class='hero'><h1>Bali Nusra Business Command Center</h1><p>Executive performance • revenue • channel • digital adoption • competitive share • program effectiveness</p><div class='badge'>LIVE EXCEL • INTERACTIVE FILTERS • SMART INSIGHTS</div></div>",unsafe_allow_html=True)
f1,f2,f3,f4=st.columns(4)
regions=["All"]+all_values(["regional","region"]);branches=["All"]+all_values(["branch","branch_lacci"]);clusters=["All"]+all_values(["cluster","cluster_name","cluster_lacci","cluster_sales"]);cities=["All"]+all_values(["city","kota","kabupaten","city_lacci"])
region=f1.selectbox("Region",regions);branch=f2.selectbox("Branch",branches);cluster=f3.selectbox("Cluster",clusters);city=f4.selectbox("City / Kabupaten",cities)

if page=="🏠 Executive Overview":
    title("Executive Scorecard","Ringkasan performa lintas sumber data dengan filter yang sama.")
    revenue=metric_from_source("Channel",["REV_ALL_BB","REV_ALL","Total Omzet","Total_Omzet"],region,branch,cluster,city);rgb=metric_from_source("Channel",["rgb_all","rgb"],region,branch,cluster,city);halo=metric_from_source("Halo",["subs","subscription","subscriber"],region,branch,cluster,city);active=metric_from_source("Skul.id",["User Active","Active Users"],region,branch,cluster,city);productive=metric_from_source("Skul.id",["User Productive","Productive Users"],region,branch,cluster,city);so=metric_from_source("Semarak",["uniq_SO","unique SO"],region,branch,cluster,city);fr=metric_from_source("Semarak",["uniq_FR","unique FR"],region,branch,cluster,city)
    cs=st.columns(6);kpi(cs[0],"REVENUE",money(revenue),"Channel data");kpi(cs[1],"RGB",count(rgb),"Channel data");kpi(cs[2],"HALO SUBS",count(halo),"Halo data");kpi(cs[3],"SKUL ACTIVE",count(active),"Skul.id");kpi(cs[4],"SKUL PRODUCTIVE",count(productive),"Skul.id");kpi(cs[5],"FR RATE",pct(fr/so if so else np.nan),f"{count(fr)} FR / {count(so)} SO")
    title("Business Pulse","Visual prioritas untuk melihat konsentrasi, tren dan competitive position.")
    left,right=st.columns([1.35,1])
    with left:
        s=sheet_with("Channel",[["REV_ALL_BB","REV_ALL","Total Omzet","Total_Omzet"]])
        if s:
            x=apply_filters(B["Channel"][s],region,branch,cluster,city);m=find_col(x,["REV_ALL_BB","REV_ALL","Total Omzet","Total_Omzet"]);c=find_col(x,["CITY","City","Kota","Kabupaten"])
            if m and c:
                g=x.assign(_m=num(safe_series(x,m))).groupby(c,dropna=False)["_m"].sum().reset_index().sort_values("_m",ascending=False).head(15);chart(px.bar(g,x="_m",y=c,orientation="h",text="_m",title="Revenue by City • Top 15",template="plotly_white"))
    with right:
        s=sheet_with("FB Youth",[["TSEL"],["IOH"],["XL"]])
        if s:
            x=apply_filters(B["FB Youth"][s],region,branch,cluster,city);terr=find_col(x,["TERITORI","Territory"]);ops=[find_col(x,[z]) for z in ["TSEL","IOH","XL"]];ops=[z for z in ops if z]
            if terr and ops:
                g=x.groupby(terr)[ops].mean().reset_index()
                for c in ops:g[c]=g[c].apply(pct_raw)
                melt=g.melt(id_vars=terr,var_name="Operator",value_name="Share");chart(px.bar(melt,x="Share",y=terr,color="Operator",orientation="h",barmode="group",title="Competitive Share by Territory",template="plotly_white"))
    title("Management Signals")
    if revenue:insight(f"<b>Revenue</b> = {money(revenue)} pada filter aktif.")
    if active and productive:insight(f"<b>Skul productivity</b> = <b>{pct(productive/active)}</b> dari active users.")
    if so:insight(f"<b>Semarak FR rate</b> = <b>{pct(fr/so)}</b>. Gunakan filter wilayah untuk menemukan area lemah.")
    if not any([revenue,rgb,halo,active,productive,so]):insight("Belum ada KPI yang terdeteksi pada filter ini. Buka Data Explorer untuk melihat struktur kolom Excel.")

elif page=="📈 Channel Performance":
    title("Channel Performance","Revenue, RGB, channel mix dan tren.")
    s=sheet_with("Channel",[["REV_ALL_BB","REV_ALL","Total Omzet","Total_Omzet"]])
    if s:
        x=apply_filters(B["Channel"][s],region,branch,cluster,city);m=find_col(x,["REV_ALL_BB","REV_ALL","Total Omzet","Total_Omzet"]);c=find_col(x,["CITY","City","Kota","Kabupaten"]);d=find_col(x,["TGL","Tanggal","Date","Period"]);a,b,c3,d3=st.columns(4);total=num(safe_series(x,m)).sum() if m else 0;a.metric("Revenue",money(total));b.metric("Rows",f"{len(x):,}");c3.metric("Cities",f"{safe_series(x,c).nunique():,}" if c else "-");d3.metric("Columns",f"{len(x.columns):,}")
        l,r=st.columns(2)
        with l:
            if m and c:
                g=x.assign(_m=num(safe_series(x,m))).groupby(c)["_m"].sum().reset_index().sort_values("_m",ascending=False).head(15);chart(px.bar(g,x="_m",y=c,orientation="h",title="Top Cities by Revenue",template="plotly_white"))
        with r:
            if m and d:
                y=x.copy();y[d]=pd.to_datetime(safe_series(y,d),errors="coerce");y["_m"]=num(safe_series(y,m));g=y.dropna(subset=[d]).groupby(d)["_m"].sum().reset_index().sort_values(d);chart(px.area(g,x=d,y="_m",title="Revenue Trend",template="plotly_white"))
        with st.expander("🔎 Filtered data"):st.dataframe(x,use_container_width=True,height=450,hide_index=True)
        st.download_button("⬇️ Download CSV",x.to_csv(index=False).encode(),"channel_filtered.csv","text/csv")
    else:st.warning("Sheet revenue Channel belum terdeteksi.")

elif page=="💰 Revenue & RGB":
    title("Revenue & RGB","Konsentrasi revenue dan kontribusi RGB.")
    srev=sheet_with("Channel",[["REV_ALL_BB","REV_ALL","Total Omzet","Total_Omzet"]]);srgb=sheet_with("Channel",[["rgb_all","rgb"]])
    if srev:
        x=apply_filters(B["Channel"][srev],region,branch,cluster,city);m=find_col(x,["REV_ALL_BB","REV_ALL","Total Omzet","Total_Omzet"]);c=find_col(x,["CITY","City","Kota","Kabupaten"]);total=num(safe_series(x,m)).sum() if m else 0;l,r=st.columns(2);l.metric("Revenue",money(total))
        if c and m:
            g=x.assign(_m=num(safe_series(x,m))).groupby(c)["_m"].sum().reset_index().sort_values("_m",ascending=False);r.metric("Top City",str(g.iloc[0][c]) if len(g) else "-");chart(px.bar(g.head(15),x="_m",y=c,orientation="h",title="Revenue Concentration",template="plotly_white"))
    if srgb:
        x=apply_filters(B["Channel"][srgb],region,branch,cluster,city);m=find_col(x,["rgb_all","rgb"]);seg=find_col(x,["brand","channel","operator"])
        if m and seg:
            g=x.assign(_m=num(safe_series(x,m))).groupby(seg)["_m"].sum().reset_index().sort_values("_m",ascending=False);l,r=st.columns(2)
            with l:chart(px.pie(g,names=seg,values="_m",hole=.58,title="RGB Mix",template="plotly_white"))
            with r:chart(px.bar(g,x="_m",y=seg,orientation="h",title="RGB by Segment",template="plotly_white"))

elif page=="📱 Halo & Device":
    title("Halo & Device","Activation, revenue, channel dan status.")
    s=sheet_with("Halo",[["subs","subscription","subscriber"]])
    if s:
        x=apply_filters(B["Halo"][s],region,branch,cluster,city);subs=find_col(x,["subs","subscription","subscriber"]);price=find_col(x,["price","revenue","amount"]);ch=find_col(x,["channel","Channel (standar dashboard)"]);flag=find_col(x,["flag","status"]);dt=find_col(x,["date","tanggal","period"])
        a,b,c,d=st.columns(4);a.metric("Subscriptions",count(num(safe_series(x,subs)).sum()) if subs else f"{len(x):,}");b.metric("Revenue / Amount",money(num(safe_series(x,price)).sum()) if price else "-");c.metric("Channels",f"{safe_series(x,ch).nunique():,}" if ch else "-");d.metric("Rows",f"{len(x):,}")
        l,r=st.columns(2)
        with l:
            if ch:
                g=x.groupby(ch).size().reset_index(name="Activations").sort_values("Activations",ascending=False).head(15);chart(px.bar(g,x="Activations",y=ch,orientation="h",title="Activation by Channel",template="plotly_white"))
        with r:
            if flag:
                g=x.groupby(flag).size().reset_index(name="Count");chart(px.pie(g,names=flag,values="Count",hole=.58,title="Status Mix",template="plotly_white"))
        if dt and subs:
            y=x.copy();y[dt]=pd.to_datetime(safe_series(y,dt),errors="coerce");y["_subs"]=num(safe_series(y,subs));g=y.dropna(subset=[dt]).groupby(dt)["_subs"].sum().reset_index();chart(px.area(g,x=dt,y="_subs",title="Subscription Trend",template="plotly_white"))
        with st.expander("🔎 Detail data"):st.dataframe(x,use_container_width=True,height=500,hide_index=True)
    else:st.warning("Sheet Halo belum terdeteksi.")

elif page=="🎓 Skul.id":
    title("Skul.id — Digital Adoption Funnel","School → users → active → productive.")
    raws=[(s,df) for s,df in B.get("Skul.id",{}).items() if "raw" in s.lower()]
    if raws:
        labels=[z[0] for z in raws];pick=st.selectbox("Period / Raw Sheet",labels,index=len(labels)-1);x=apply_filters(dict(raws)[pick],region,branch,cluster,city);school=find_col(x,["ID Sekolah","School ID"]);total=find_col(x,["Total Users","Total User"]);active=find_col(x,["User Active","Active Users"]);prod=find_col(x,["User Productive","Productive Users"]);cc=find_col(x,["Kota","City","Kabupaten"]);schools=x[school].nunique() if school else 0;tu=num(safe_series(x,total)).sum() if total else 0;au=num(safe_series(x,active)).sum() if active else 0;pu=num(safe_series(x,prod)).sum() if prod else 0
        a,b,c,d=st.columns(4);a.metric("Schools",count(schools));b.metric("Total Users",count(tu));c.metric("Active",count(au));d.metric("Productive",count(pu))
        if tu:st.progress(min(au/tu,1),text=f"Active conversion: {pct(au/tu)}")
        if au:st.progress(min(pu/au,1),text=f"Productive conversion: {pct(pu/au)}")
        l,r=st.columns(2)
        with l:
            stages=pd.DataFrame({"Stage":["Total Users","Active","Productive"],"Users":[tu,au,pu]});chart(px.funnel(stages,x="Users",y="Stage",title="Adoption Funnel",template="plotly_white"))
        with r:
            if cc and active:
                g=x.assign(_a=num(safe_series(x,active))).groupby(cc)["_a"].sum().reset_index().sort_values("_a",ascending=False).head(15);chart(px.bar(g,x="_a",y=cc,orientation="h",title="Active Users by City",template="plotly_white"))
        if cc and active and prod:
            g=x.assign(_a=num(safe_series(x,active)),_p=num(safe_series(x,prod))).groupby(cc).agg(Active=("_a","sum"),Productive=("_p","sum")).reset_index();g["Productive Rate"]=np.where(g.Active>0,g.Productive/g.Active,0);st.dataframe(g.sort_values("Productive Rate",ascending=False),use_container_width=True,hide_index=True,column_config={"Productive Rate":st.column_config.ProgressColumn("Productive Rate",min_value=0,max_value=1,format="%.1%")})
    else:st.warning("Sheet Raw Skul.id belum ditemukan.")

elif page=="⚔️ FB Youth":
    title("FB Youth — Competitive Intelligence","TSEL vs IOH vs XL per territory.")
    s=sheet_with("FB Youth",[["TSEL"],["IOH"],["XL"]])
    if s:
        x=apply_filters(B["FB Youth"][s],region,branch,cluster,city);terr=find_col(x,["TERITORI","Territory"]);dt=find_col(x,["TANGGAL","Tanggal","Date"]);ops=[find_col(x,[z]) for z in ["TSEL","IOH","XL"]];ops=[z for z in ops if z]
        for c in ops:x[c]=safe_series(x,c).apply(pct_raw)
        if dt:x[dt]=pd.to_datetime(safe_series(x,dt),errors="coerce")
        l,r=st.columns(2)
        with l:
            if terr and ops:
                g=x.groupby(terr)[ops].mean().reset_index();melt=g.melt(id_vars=terr,var_name="Operator",value_name="Share");chart(px.bar(melt,x="Share",y=terr,color="Operator",orientation="h",barmode="group",title="Average Share by Territory",template="plotly_white"))
        with r:
            if dt and ops:
                g=x.dropna(subset=[dt]).groupby(dt)[ops].mean().reset_index();melt=g.melt(id_vars=dt,var_name="Operator",value_name="Share");fig=px.line(melt,x=dt,y="Share",color="Operator",markers=True,title="Competitive Share Trend",template="plotly_white");fig.update_yaxes(tickformat=".0%");chart(fig)
        if terr and ops:
            g=x.groupby(terr)[ops].mean().reset_index();g["Winner"]=g[ops].idxmax(axis=1);wins=g["Winner"].value_counts().reset_index();wins.columns=["Operator","Territories Won"];chart(px.bar(wins,x="Operator",y="Territories Won",text="Territories Won",title="Territory Win Count",template="plotly_white"))
        with st.expander("🔎 Detail competitive data"):st.dataframe(x,use_container_width=True,height=500,hide_index=True)
    else:st.warning("Sheet FB Youth belum terdeteksi.")

elif page=="🚀 Program Semarak":
    title("Program Semarak — SO → FR Conversion","Cari area dengan conversion terbaik dan terlemah.")
    s=sheet_with("Semarak",[["uniq_SO","unique SO"],["uniq_FR","unique FR"]])
    if s:
        x=apply_filters(B["Semarak"][s],region,branch,cluster,city);so_c=find_col(x,["uniq_SO","unique SO"]);fr_c=find_col(x,["uniq_FR","unique FR"]);cl=find_col(x,["cluster_name","cluster"]);br=find_col(x,["branch"]);status=find_col(x,["status_fr","status"]);so=num(safe_series(x,so_c)).sum();fr=num(safe_series(x,fr_c)).sum();a,b,c,d=st.columns(4);a.metric("Unique SO",count(so));b.metric("Successful FR",count(fr));c.metric("FR Rate",pct(fr/so if so else np.nan));d.metric("Rows",f"{len(x):,}")
        if cl:
            z=x.assign(_so=num(safe_series(x,so_c)),_fr=num(safe_series(x,fr_c))).groupby(cl).agg(SO=("_so","sum"),FR=("_fr","sum")).reset_index();z["FR Rate"]=np.where(z.SO>0,z.FR/z.SO,0);l,r=st.columns(2)
            with l:chart(px.bar(z.sort_values("SO",ascending=False),x=cl,y=["SO","FR"],barmode="group",title="SO vs FR by Cluster",template="plotly_white"))
            with r:chart(px.bar(z.sort_values("FR Rate",ascending=False),x="FR Rate",y=cl,orientation="h",title="FR Rate Ranking",template="plotly_white"))
            st.dataframe(z.sort_values("FR Rate",ascending=False),use_container_width=True,hide_index=True,column_config={"FR Rate":st.column_config.ProgressColumn("FR Rate",min_value=0,max_value=1,format="%.1%")})
        if br:
            z=x.assign(_fr=num(safe_series(x,fr_c))).groupby(br)["_fr"].sum().reset_index().sort_values("_fr",ascending=False).head(15);chart(px.bar(z,x="_fr",y=br,orientation="h",title="Top Branch by Successful FR",template="plotly_white"))
        if status:
            z=x.groupby(status).size().reset_index(name="Count");chart(px.pie(z,names=status,values="Count",hole=.58,title="FR Status Mix",template="plotly_white"))
    else:st.warning("Data Program Semarak belum terdeteksi.")

else:
    title("Data Explorer","Telusuri sheet Excel secara langsung. Tabel dapat search, sort dan download.")
    sources=[s for s,v in B.items() if v]
    if sources:
        src=st.selectbox("Source",sources);sheet=st.selectbox("Sheet",list(B[src].keys()));x=apply_filters(B[src][sheet],region,branch,cluster,city);a,b,c=st.columns(3);a.metric("Rows",f"{len(x):,}");b.metric("Columns",f"{len(x.columns):,}");c.metric("Missing Cells",f"{int(x.isna().sum().sum()):,}");q=st.text_input("🔎 Search across all columns")
        if q:
            mask=x.astype(str).apply(lambda row:row.str.contains(q,case=False,na=False).any(),axis=1);x=x[mask]
        st.dataframe(x,use_container_width=True,height=650,hide_index=True);st.download_button("⬇️ Download filtered CSV",x.to_csv(index=False).encode(),"bali_nusra_filtered.csv","text/csv")
    else:st.error("Tidak ada file Excel yang berhasil dibaca dari folder data/.")

st.markdown("---");st.caption(f"Bali Nusra Business Command Center V5 • Last refresh: {datetime.now().strftime('%d %b %Y %H:%M')} • Excel sources in /data")
