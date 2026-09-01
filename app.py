import os
import re
from datetime import datetime
import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go

st.set_page_config(page_title="Bali Nusra | Business Command Center", page_icon="📊", layout="wide", initial_sidebar_state="expanded")
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
:root{--ink:#101828;--muted:#667085;--line:#e4e7ec;--soft:#f7f9fc;--brand:#00a88f;--brand2:#0b5cab;--danger:#e5484d;}
html,body,[class*="css"]{font-family:Inter,sans-serif}
.block-container{max-width:1680px;padding:1.1rem 1.5rem 2rem}
[data-testid="stSidebar"]{background:#f8fafc;border-right:1px solid var(--line)}
[data-testid="stSidebar"] .block-container{padding:1.2rem .9rem}
.hero{background:linear-gradient(120deg,#081a2b 0%,#10355b 58%,#087d72 100%);border-radius:24px;padding:27px 32px;color:white;box-shadow:0 16px 42px rgba(16,24,40,.16);margin-bottom:18px}
.hero-top{display:flex;justify-content:space-between;gap:18px;align-items:flex-start}.brand{font-size:13px;font-weight:800;letter-spacing:.06em;text-transform:uppercase;opacity:.92}.hero h1{font-size:34px;line-height:1.05;margin:7px 0 6px;font-weight:800;letter-spacing:-.045em}.hero p{margin:0;color:#d9e7f5;font-size:13px}.live{display:inline-flex;align-items:center;gap:7px;background:rgba(255,255,255,.11);border:1px solid rgba(255,255,255,.14);padding:7px 11px;border-radius:999px;font-size:10px;font-weight:800;white-space:nowrap}.dot{width:7px;height:7px;border-radius:50%;background:#54e2c0;display:inline-block}
.filterbar{background:#fff;border:1px solid var(--line);border-radius:18px;padding:13px 14px;box-shadow:0 5px 20px rgba(16,24,40,.045);margin-bottom:18px}.filter-label{font-size:10px;font-weight:800;text-transform:uppercase;letter-spacing:.08em;color:var(--muted);margin-bottom:3px}
.kpi{background:#fff;border:1px solid var(--line);border-radius:17px;padding:16px 17px;min-height:128px;box-shadow:0 5px 18px rgba(16,24,40,.045);position:relative;overflow:hidden}.kpi:after{content:"";position:absolute;right:-22px;bottom:-28px;width:80px;height:80px;border-radius:50%;background:rgba(0,168,143,.055)}.kpi .top{display:flex;justify-content:space-between;align-items:center}.kpi .ico{width:30px;height:30px;border-radius:9px;background:#eef8f6;display:flex;align-items:center;justify-content:center;font-size:14px}.kpi .trend{font-size:10px;font-weight:800;color:#07916f}.kpi .label{font-size:10px;color:var(--muted);font-weight:800;letter-spacing:.08em;text-transform:uppercase;margin-top:13px}.kpi .value{font-size:25px;font-weight:800;color:var(--ink);margin-top:4px;letter-spacing:-.03em}.kpi .note{font-size:10px;color:var(--muted);margin-top:4px}
.panel{background:#fff;border:1px solid var(--line);border-radius:18px;padding:17px 18px;box-shadow:0 5px 18px rgba(16,24,40,.04);height:100%}.panel-title{font-size:16px;font-weight:800;color:var(--ink);margin-bottom:2px}.panel-sub{font-size:11px;color:var(--muted);margin-bottom:8px}.section{font-size:20px;font-weight:800;color:var(--ink);margin:22px 0 9px;letter-spacing:-.025em}.sub{font-size:11px;color:var(--muted);margin-bottom:10px}.insight{background:#f8fafc;border:1px solid var(--line);border-radius:13px;padding:12px 13px;font-size:12px;color:#344054;margin-top:8px}.status{font-size:10px;color:#475467;background:#f8fafc;border:1px solid #eaecf0;border-radius:10px;padding:7px 9px}.small{font-size:10px;color:var(--muted)}
div[data-testid="stMetric"]{background:#fff;border:1px solid var(--line);border-radius:14px;padding:10px}
button[kind="secondary"]{border-radius:10px}
</style>
""", unsafe_allow_html=True)

def norm(v): return re.sub(r"[^a-z0-9]+", "", str(v).strip().casefold())

def clean(df):
    if df is None or not isinstance(df,pd.DataFrame) or df.empty: return pd.DataFrame()
    x=df.copy(); seen={}; cols=[]
    for i,c in enumerate(x.columns):
        b=re.sub(r"\s+"," ",str(c).strip())
        if not b or norm(b) in {"nan","none","null"}: b=f"Column_{i+1}"
        n=seen.get(norm(b),0)+1; seen[norm(b)]=n; cols.append(b if n==1 else f"{b}__{n}")
    x.columns=cols
    return x.dropna(how="all").reset_index(drop=True)

def safe_series(df,col):
    if df is None or not isinstance(df,pd.DataFrame) or col is None or col not in df.columns: return pd.Series(index=df.index if isinstance(df,pd.DataFrame) else [],dtype=object)
    s=df[col]; return s.iloc[:,0] if isinstance(s,pd.DataFrame) else s

def num(s):
    if isinstance(s,pd.DataFrame): s=s.iloc[:,0]
    if s is None: return pd.Series(dtype=float)
    if pd.api.types.is_numeric_dtype(s): return pd.to_numeric(s,errors="coerce")
    z=s.astype(str).str.strip().str.replace(r"\((.*?)\)",r"-\1",regex=True)
    z=z.str.replace(",","",regex=False).str.replace("%","",regex=False).str.replace("Rp","",regex=False).str.replace("IDR","",regex=False).str.replace(" ","",regex=False)
    return pd.to_numeric(z.str.replace(r"[^0-9eE.\-+]","",regex=True),errors="coerce")

def money(v):
    if v is None or pd.isna(v): return "-"
    v=float(v); a=abs(v)
    if a>=1e9:return f"Rp {v/1e9:.2f} B"
    if a>=1e6:return f"Rp {v/1e6:.2f} M"
    if a>=1e3:return f"Rp {v/1e3:.1f} K"
    return f"Rp {v:,.0f}"

def count(v):
    if v is None or pd.isna(v): return "-"
    v=float(v); a=abs(v)
    if a>=1e6:return f"{v/1e6:.2f} M"
    if a>=1e3:return f"{v/1e3:.1f} K"
    return f"{v:,.0f}"

def pct(v):
    if v is None or pd.isna(v): return "-"
    return f"{float(v)*100:.1f}%"

def pct_value(v):
    if v is None or pd.isna(v): return np.nan
    v=float(v); return v/100 if abs(v)>1.5 else v

ALIASES={
 "region":["regional","region","region name","region_name","region lacci","region sales","regional name","area region","area"],
 "branch":["branch","branch name","branch_name","branch lacci","branch sales","branch code"],
 "cluster":["cluster","cluster name","cluster_name","cluster lacci","cluster sales","new cluster","cluster territory"],
 "city":["city","city name","city_name","city lacci","city_name_lacci","kota","kabupaten","kabupaten kota","sales territory value","territory","teritori","sales territory"],
 "date":["date","tanggal","tgl","period","periode","month","bulan","week","minggu","activation date","tgl trx","tgl fr","so date","regis date","used date","trx date","transaction date"],
 "revenue":["rev_all_bb","rev_all","revenue_all","revenue","total revenue","total omzet","omzet","sales revenue","total amount sellthru","amount revenue","revenue amount","rev"],
 "rgb":["rgb_all","rgb","total rgb","rgb users","paying user rgb all","rgb user"],
 "halo":["subs","subscription","subscriber","halo subs","halo subscriber","halo","jumlah subs","total subscriber"],
 "price":["price","harga paket","amount","nominal","value","total amount","selling price","harga"],
 "active":["user active","active users","active_user","users active","active user","active"],
 "productive":["user productive","productive users","productive_user","users productive","productive user","productive"],
 "users":["total users","total user","users","user count","jumlah user","registered users","register user"],
 "school":["school","schools","sekolah","school count","jumlah sekolah","active school","register school","school name"],
 "school_id":["id sekolah","school id","school_id","npsn","school code","kode sekolah"],
 "student":["student","students","siswa","student count","jumlah siswa"],
 "so":["uniq_so","unique so","total so","sales order","so unik","so (unik)","so"],
 "fr":["uniq_fr","unique fr","first recharge","fr sukses","fr success","fr"],
 "channel":["channel","channel standar dashboard","channel standard","channel name","sales channel"],
 "operator":["operator","ops","tsel","ioh","xl","xl smart","xlsmart"],
}

def find_col(df,key,exclude=None):
    if df is None or df.empty:return None
    aliases=ALIASES.get(key,[key]); cols=list(df.columns); ex={norm(x) for x in (exclude or [])}
    exact={norm(c):c for c in cols if norm(c) not in ex}
    for a in aliases:
        if norm(a) in exact:return exact[norm(a)]
    ranked=[]
    for c in cols:
        nc=norm(c)
        if nc in ex:continue
        for a in aliases:
            na=norm(a)
            if not na:continue
            score=0
            if na in nc: score=90+len(na)
            elif nc in na: score=70+len(nc)
            else:
                at=set(re.findall(r"[a-z0-9]+",str(a).casefold())); ct=set(re.findall(r"[a-z0-9]+",str(c).casefold()))
                if at and len(at & ct)>=max(1,len(at)//2): score=40+len(at&ct)*5
            if score:
                bad=(key in {"active","productive","users","school","so","fr","halo","price","revenue","rgb"} and any(b in nc for b in ["flag","growth","target","ach","weight","score","rate","pct"]))
                if not bad: ranked.append((score,c)); break
    return max(ranked,key=lambda z:z[0])[1] if ranked else None

def detect_dim(df,key): return find_col(df,key)
def detect_metric(df,key): return find_col(df,key)

def header_score(row):
    vals=[str(v).strip() for v in row.tolist() if pd.notna(v)]
    if not vals:return -999
    tokens=["region","branch","cluster","city","kota","kabupaten","territory","active","productive","users","revenue","rev","rgb","subs","price","so","fr","channel","tanggal","date","school","sekolah","siswa","operator","tsel","ioh","xl","trx","transaction"]
    hit=sum(any(t in norm(v) for t in tokens) for v in vals); text=sum(not bool(re.fullmatch(r"[-+]?\d+(\.\d+)?",v)) for v in vals)/max(len(vals),1); numeric=sum(bool(re.fullmatch(r"[-+]?\d+(\.\d+)?",v)) for v in vals)/max(len(vals),1)
    return hit*5+text*1.5-numeric*2

def merged_channel(path,sheet):
    raw=pd.read_excel(path,sheet_name=sheet,header=None)
    if len(raw)<4:return pd.DataFrame()
    scores=[header_score(raw.iloc[r]) for r in range(min(8,len(raw)))]
    h=int(np.argmax(scores)); h=max(h,2) if h<3 and len(raw)>3 else h
    x=raw.iloc[h+1:].copy(); x.columns=raw.iloc[h].tolist(); return clean(x)

@st.cache_data(show_spinner=False)
def read_book(path):
    out={}
    if not os.path.exists(path): return out
    try: xl=pd.ExcelFile(path)
    except Exception:return out
    for sheet in xl.sheet_names:
        try:
            if os.path.basename(path)=="Channel Dashboard.xlsx" and str(sheet).strip() in {"202607","202608"}: x=merged_channel(path,sheet)
            else:
                raw=pd.read_excel(path,sheet_name=sheet,header=None,nrows=15)
                if raw.empty:continue
                scores=[header_score(raw.iloc[r]) for r in range(min(10,len(raw)))]
                h=int(np.argmax(scores)); x=pd.read_excel(path,sheet_name=sheet,header=h); x=clean(x)
            if not x.empty: out[str(sheet).strip()]=x
        except Exception: continue
    return out

@st.cache_data(show_spinner=False)
def load_all(): return {k:read_book(os.path.join(DATA_DIR,f)) for k,f in FILES.items()}
B=load_all()

def all_frames(source=None):
    sources=[source] if source else list(B)
    for src in sources:
        for sheet,df in B.get(src,{}).items():
            if isinstance(df,pd.DataFrame) and not df.empty: yield src,sheet,df

def unique_values(df,col):
    if df is None or df.empty or col is None:return []
    s=safe_series(df,col).dropna().astype(str).str.strip(); bad={"","nan","none","null","-","all","unknown"}
    return sorted([v for v in s.unique() if v.casefold() not in bad],key=str.casefold)

def apply_filters(df,filters):
    x=clean(df)
    for dim,val in (filters or {}).items():
        if val in (None,"All") or x.empty:continue
        c=detect_dim(x,dim)
        if c:
            s=safe_series(x,c).astype(str).str.strip().str.casefold(); x=x[s==str(val).strip().casefold()]
    return x

def dimension_values(dim,filters=None):
    vals=set(); filters=filters or {}
    for src,_,df in all_frames():
        x=apply_filters(df,{k:v for k,v in filters.items() if k!=dim}); c=detect_dim(x,dim)
        if c: vals.update(unique_values(x,c))
    return sorted(vals,key=str.casefold)

def best_sheet(source,metrics=None,dims=None,keywords=None):
    metrics=metrics or []; dims=dims or []; keys=[norm(k) for k in (keywords or [])]; best=None; bestscore=-1e9
    for sheet,df in B.get(source,{}).items():
        score=0; sn=norm(sheet)
        for m in metrics:
            if detect_metric(df,m):score+=12
        for d in dims:
            if detect_dim(df,d):score+=4
        score+=sum(6 for k in keys if k and k in sn)
        if source=="Skul.id":
            if "rawdata" in sn or "raw" in sn:score+=22
            if "summary" in sn:score-=5
        if source=="Semarak":
            if "rawsemarak" in sn or "rawallnew" in sn or "raw" in sn:score+=22
            if "summary" in sn or "dashboard" in sn:score-=5
        if source=="Channel":
            if "revall" in sn and "revenue" in metrics:score+=25
            if sn=="rgb" and "rgb" in metrics:score+=25
        if source=="Halo" and "halo" in metrics:
            if "revhaloall" in sn:score+=25
            if "device" in sn:score-=8
        score+=min(len(df),100000)/100000+min(len(df.columns),100)*.02
        if score>bestscore:bestscore,best=score,sheet
    return best

def source_data(source,metrics=None,dims=None,keywords=None,filters=None):
    sh=best_sheet(source,metrics,dims,keywords)
    return (sh,apply_filters(B[source][sh],filters or {})) if sh else (None,pd.DataFrame())

def sum_metric(source,metric,filters):
    sh=best_sheet(source,[metric],["region","branch","cluster","city"])
    if not sh:return np.nan
    x=apply_filters(B[source][sh],filters); c=detect_metric(x,metric)
    if not c:return np.nan
    v=num(safe_series(x,c)).sum(min_count=1); return float(v) if pd.notna(v) else np.nan

def make_chart(fig,height=340):
    fig.update_layout(height=height,margin=dict(l=8,r=8,t=46,b=8),paper_bgcolor="rgba(0,0,0,0)",plot_bgcolor="rgba(0,0,0,0)",font=dict(family="Inter",color="#344054",size=11),legend=dict(orientation="h",yanchor="bottom",y=1.01,x=0),hovermode="x unified")
    fig.update_xaxes(showgrid=True,gridcolor="#eef2f6",zeroline=False); fig.update_yaxes(showgrid=False,zeroline=False)
    return fig

def chart(fig,height=340): st.plotly_chart(make_chart(fig,height),use_container_width=True,config={"displaylogo":False,"responsive":True})

def kpi(c,label,value,note="",icon="◉",trend=""):
    c.markdown(f'<div class="kpi"><div class="top"><div class="ico">{icon}</div><div class="trend">{trend}</div></div><div class="label">{label}</div><div class="value">{value}</div><div class="note">{note}</div></div>',unsafe_allow_html=True)

def title(t,s=None):
    st.markdown(f'<div class="section">{t}</div>',unsafe_allow_html=True)
    if s: st.markdown(f'<div class="sub">{s}</div>',unsafe_allow_html=True)

def detected(items):
    st.dataframe(pd.DataFrame({"Parameter":[a for a,_ in items],"Detected Column":[b if b else "— not found —" for _,b in items]}),use_container_width=True,hide_index=True)

def trend_text(df,metric,date_col):
    if not metric or not date_col:return ""
    z=pd.DataFrame({"d":pd.to_datetime(safe_series(df,date_col),errors="coerce"),"v":num(safe_series(df,metric))}).dropna()
    if z.empty:return ""
    z["p"]=z.d.dt.to_period("M"); g=z.groupby("p").v.sum()
    if len(g)<2 or g.iloc[-2]==0:return ""
    ch=(g.iloc[-1]/g.iloc[-2]-1)*100
    return f"↗ {ch:.1f}% MoM" if ch>=0 else f"↘ {abs(ch):.1f}% MoM"

with st.sidebar:
    st.markdown("## 📊 BALI NUSRA")
    st.caption("Business Command Center • V8")
    if st.button("🔄 Refresh data",use_container_width=True): st.cache_data.clear(); st.rerun()
    st.divider()
    page=st.radio("NAVIGATION",["🏠 Executive Overview","📈 Channel Performance","💰 Revenue & RGB","📱 Halo & Device","🎓 Skul.id","⚔️ FB Youth","🚀 Program Semarak","🗂️ Data Explorer"])
    st.divider(); st.caption("DATA SOURCES")
    for key,fn in FILES.items():
        sheets=len(B.get(key,{})); st.write(f"{'🟢' if sheets else '🔴'} {key} · {sheets} sheet")

st.markdown("<div class='hero'><div class='hero-top'><div><div class='brand'>Telkomsel • Bali Nusra</div><h1>Business Command Center</h1><p>Executive performance • revenue • channel • digital adoption • competitive share • program effectiveness</p></div><div class='live'><span class='dot'></span> LIVE EXCEL • SMART DETECTION</div></div></div>",unsafe_allow_html=True)

st.markdown("<div class='filterbar'>",unsafe_allow_html=True)
filters={}; f1,f2,f3,f4=st.columns(4)
with f1:
    st.markdown("<div class='filter-label'>Region</div>",unsafe_allow_html=True); filters["region"]=st.selectbox("",["All"]+dimension_values("region"),key="f_region",label_visibility="collapsed")
with f2:
    st.markdown("<div class='filter-label'>Branch</div>",unsafe_allow_html=True); filters["branch"]=st.selectbox("",["All"]+dimension_values("branch",{"region":filters["region"]}),key="f_branch",label_visibility="collapsed")
with f3:
    st.markdown("<div class='filter-label'>Cluster</div>",unsafe_allow_html=True); filters["cluster"]=st.selectbox("",["All"]+dimension_values("cluster",{"region":filters["region"],"branch":filters["branch"]}),key="f_cluster",label_visibility="collapsed")
with f4:
    st.markdown("<div class='filter-label'>City / Kabupaten</div>",unsafe_allow_html=True); filters["city"]=st.selectbox("",["All"]+dimension_values("city",{"region":filters["region"],"branch":filters["branch"],"cluster":filters["cluster"]}),key="f_city",label_visibility="collapsed")
st.markdown("</div>",unsafe_allow_html=True)

if page=="🏠 Executive Overview":
    title("Executive Dashboard","Satu layar untuk membaca performance, growth, concentration dan competitive position.")
    rev=sum_metric("Channel","revenue",filters); rgb=sum_metric("Channel","rgb",filters); halo=sum_metric("Halo","halo",filters); active=sum_metric("Skul.id","active",filters); productive=sum_metric("Skul.id","productive",filters); so=sum_metric("Semarak","so",filters); fr=sum_metric("Semarak","fr",filters)
    shc,xc=source_data("Channel",["revenue"],["date","city"],filters=filters); shs,xs=source_data("Skul.id",["active","productive"],["date","city"],filters=filters)
    cards=st.columns(6)
    vals=[(money(rev),"Revenue","Channel","Rp","↗"),(count(rgb),"RGB","Channel","👥","↗"),(count(halo),"Halo Subscribers","Halo","📱","↗"),(count(active),"Skul Active","Skul.id","🎓","↗"),(count(productive),"Skul Productive","Skul.id","⚡","↗"),(pct(fr/so) if pd.notna(fr) and pd.notna(so) and so else "-","FR Conversion","Semarak","🚀","↗")]
    for c,(v,l,n,i,t) in zip(cards,vals): kpi(c,l,v,n,i,t)

    title("Performance Overview","Trend otomatis mengikuti periode yang tersedia di Excel.")
    left,right=st.columns([1.65,1])
    with left:
        st.markdown('<div class="panel"><div class="panel-title">Revenue & RGB Trend</div><div class="panel-sub">Monthly movement after applying the global filters</div>',unsafe_allow_html=True)
        if shc:
            dc=detect_dim(xc,"date"); rc=detect_metric(xc,"revenue"); gc=detect_metric(xc,"rgb")
            if dc and (rc or gc):
                z=pd.DataFrame({"Date":pd.to_datetime(safe_series(xc,dc),errors="coerce")})
                if rc:z["Revenue"]=num(safe_series(xc,rc))
                if gc:z["RGB"]=num(safe_series(xc,gc))
                z=z.dropna(subset=["Date"]); z["Period"]=z.Date.dt.to_period("M").astype(str); g=z.groupby("Period")[[c for c in ["Revenue","RGB"] if c in z]].sum().reset_index()
                if not g.empty:
                    fig=go.Figure()
                    if "Revenue" in g: fig.add_trace(go.Scatter(x=g.Period,y=g.Revenue,name="Revenue",mode="lines+markers",fill="tozeroy"))
                    if "RGB" in g: fig.add_trace(go.Scatter(x=g.Period,y=g.RGB,name="RGB",mode="lines+markers"))
                    chart(fig,350)
            else: st.info("Date parameter belum ditemukan untuk membuat trend.")
        else: st.warning("Channel data belum terbaca.")
        st.markdown('</div>',unsafe_allow_html=True)
    with right:
        st.markdown('<div class="panel"><div class="panel-title">Business Mix</div><div class="panel-sub">Revenue contribution by city</div>',unsafe_allow_html=True)
        if shc:
            city=detect_dim(xc,"city"); rc=detect_metric(xc,"revenue")
            if city and rc:
                g=xc.assign(_v=num(safe_series(xc,rc))).groupby(city)._v.sum().reset_index().sort_values("_v",ascending=False).head(8); chart(px.pie(g,names=city,values="_v",hole=.62,title="",template="plotly_white"),320)
            else: st.info("City/Revenue belum terdeteksi.")
        st.markdown('</div>',unsafe_allow_html=True)

    title("Management View","Area dengan contribution terbesar, adoption funnel dan competitive share.")
    a,b,c=st.columns(3)
    with a:
        st.markdown('<div class="panel"><div class="panel-title">Top Cities by Revenue</div><div class="panel-sub">Top 10 filtered cities</div>',unsafe_allow_html=True)
        if shc:
            city=detect_dim(xc,"city"); rc=detect_metric(xc,"revenue")
            if city and rc:
                g=xc.assign(_v=num(safe_series(xc,rc))).groupby(city)._v.sum().reset_index().sort_values("_v",ascending=False).head(10); chart(px.bar(g,x="_v",y=city,orientation="h",text_auto=True,template="plotly_white",title=""),330)
        st.markdown('</div>',unsafe_allow_html=True)
    with b:
        st.markdown('<div class="panel"><div class="panel-title">Digital Adoption Funnel</div><div class="panel-sub">School → users → active → productive</div>',unsafe_allow_html=True)
        if shs:
            sid=detect_metric(xs,"school_id"); users=detect_metric(xs,"users"); act=detect_metric(xs,"active"); prod=detect_metric(xs,"productive"); school=xs[sid].nunique() if sid else 0; uv=num(safe_series(xs,users)).sum() if users else np.nan; av=num(safe_series(xs,act)).sum() if act else np.nan; pv=num(safe_series(xs,prod)).sum() if prod else np.nan
            fg=pd.DataFrame({"Stage":["Schools","Users","Active","Productive"],"Value":[school,uv,av,pv]}); chart(px.funnel(fg,y="Stage",x="Value",template="plotly_white",title=""),330)
        else: st.warning("Skul.id sheet belum terdeteksi.")
        st.markdown('</div>',unsafe_allow_html=True)
    with c:
        st.markdown('<div class="panel"><div class="panel-title">Competitive Share</div><div class="panel-sub">TSEL vs IOH vs XL Smart</div>',unsafe_allow_html=True)
        sh,x=source_data("FB Youth",dims=["region","branch","cluster","city"],keywords=["fb","youth"],filters=filters)
        if sh:
            terr=find_col(x,"city") or find_col(x,"region"); opcols=[]
            for key in ["tsel","ioh","xl"]:
                col=None
                for cc in x.columns:
                    nn=norm(cc)
                    if key in nn and not any(bb in nn for bb in ["flag","growth","target","ach","score"]): col=cc; break
                if col:opcols.append(col)
            if terr and opcols:
                g=x.groupby(terr)[opcols].mean().reset_index()
                for oc in opcols:g[oc]=g[oc].apply(pct_value)
                m=g.melt(id_vars=[terr],var_name="Operator",value_name="Share"); chart(px.bar(m,x="Share",y=terr,color="Operator",orientation="h",barmode="group",template="plotly_white",title=""),330)
            else: st.info("Operator/territory belum terdeteksi.")
        else: st.warning("FB Youth belum terbaca.")
        st.markdown('</div>',unsafe_allow_html=True)

    title("Executive Signals")
    s1,s2,s3=st.columns(3)
    with s1: st.markdown(f'<div class="insight">💰 <b>Revenue</b><br>{money(rev)} setelah filter wilayah.</div>',unsafe_allow_html=True)
    with s2: st.markdown(f'<div class="insight">🎓 <b>Productivity</b><br>{pct(productive/active) if pd.notna(active) and active and pd.notna(productive) else "-"} active users menjadi productive.</div>',unsafe_allow_html=True)
    with s3: st.markdown(f'<div class="insight">🚀 <b>Semarak</b><br>{pct(fr/so) if pd.notna(so) and so and pd.notna(fr) else "-"} FR conversion dari SO.</div>',unsafe_allow_html=True)

elif page=="📈 Channel Performance":
    title("Channel Performance","Revenue, RGB, channel mix, city contribution dan trend.")
    sh,x=source_data("Channel",["revenue","rgb"],["region","branch","cluster","city","date"],filters=filters)
    if sh:
        rev,rgb,city,date=[detect_metric(x,"revenue"),detect_metric(x,"rgb"),detect_dim(x,"city"),detect_dim(x,"date")]
        c=st.columns(4); kpi(c[0],"REVENUE",money(num(safe_series(x,rev)).sum()) if rev else "-",sh,"💰",trend_text(x,rev,date)); kpi(c[1],"RGB",count(num(safe_series(x,rgb)).sum()) if rgb else "-",sh,"👥",trend_text(x,rgb,date)); kpi(c[2],"CITIES",f"{x[city].nunique():,}" if city else "-","unique","📍",""); kpi(c[3],"ROWS",f"{len(x):,}","filtered","▤","")
        l,r=st.columns(2)
        with l:
            st.markdown('<div class="panel"><div class="panel-title">Revenue by City</div><div class="panel-sub">Top 20 cities</div>',unsafe_allow_html=True)
            if city and rev:
                g=x.assign(_v=num(safe_series(x,rev))).groupby(city)._v.sum().reset_index().sort_values("_v",ascending=False).head(20); chart(px.bar(g,x="_v",y=city,orientation="h",template="plotly_white",title=""),400)
            st.markdown('</div>',unsafe_allow_html=True)
        with r:
            st.markdown('<div class="panel"><div class="panel-title">Revenue Trend</div><div class="panel-sub">Monthly trend</div>',unsafe_allow_html=True)
            if date and rev:
                z=pd.DataFrame({"Date":pd.to_datetime(safe_series(x,date),errors="coerce"),"Revenue":num(safe_series(x,rev))}).dropna(subset=["Date"]); z["Period"]=z.Date.dt.to_period("M").astype(str); g=z.groupby("Period").Revenue.sum().reset_index(); chart(px.area(g,x="Period",y="Revenue",markers=True,template="plotly_white",title=""),400)
            st.markdown('</div>',unsafe_allow_html=True)
        st.divider(); detected([("Revenue",rev),("RGB",rgb),("Region",detect_dim(x,"region")),("Branch",detect_dim(x,"branch")),("Cluster",detect_dim(x,"cluster")),("City",city),("Date",date)])
    else: st.warning("Channel Dashboard.xlsx belum terbaca.")

elif page=="💰 Revenue & RGB":
    title("Revenue & RGB","Contribution, concentration dan productivity per area.")
    sh,x=source_data("Channel",["revenue","rgb"],["region","branch","cluster","city","date"],filters=filters)
    if sh:
        rev,rgb,city,date=[detect_metric(x,"revenue"),detect_metric(x,"rgb"),detect_dim(x,"city"),detect_dim(x,"date")]
        c=st.columns(4); kpi(c[0],"REVENUE",money(num(safe_series(x,rev)).sum()) if rev else "-",sh,"💰",trend_text(x,rev,date)); kpi(c[1],"RGB",count(num(safe_series(x,rgb)).sum()) if rgb else "-",sh,"👥",trend_text(x,rgb,date)); kpi(c[2],"AVG RGB / CITY",count(num(safe_series(x,rgb)).sum()/max(x[city].nunique(),1)) if rgb and city else "-","simple average","📊",""); kpi(c[3],"TOP CITY",str(x.assign(_v=num(safe_series(x,rev))).groupby(city)._v.sum().idxmax()) if rev and city and not x.empty else "-","revenue leader","🏆","")
        l,r=st.columns(2)
        with l:
            st.markdown('<div class="panel"><div class="panel-title">Revenue Concentration</div><div class="panel-sub">Where the business comes from</div>',unsafe_allow_html=True)
            if city and rev:
                g=x.assign(_v=num(safe_series(x,rev))).groupby(city)._v.sum().reset_index().sort_values("_v",ascending=False).head(15); chart(px.bar(g,x="_v",y=city,orientation="h",template="plotly_white",title=""),400)
            st.markdown('</div>',unsafe_allow_html=True)
        with r:
            st.markdown('<div class="panel"><div class="panel-title">RGB by City</div><div class="panel-sub">Top contribution</div>',unsafe_allow_html=True)
            if city and rgb:
                g=x.assign(_v=num(safe_series(x,rgb))).groupby(city)._v.sum().reset_index().sort_values("_v",ascending=False).head(15); chart(px.bar(g,x="_v",y=city,orientation="h",template="plotly_white",title=""),400)
            st.markdown('</div>',unsafe_allow_html=True)
        st.divider(); detected([("Revenue",rev),("RGB",rgb),("Region",detect_dim(x,"region")),("Branch",detect_dim(x,"branch")),("Cluster",detect_dim(x,"cluster")),("City",city),("Date",date)])

elif page=="📱 Halo & Device":
    title("Halo & Device","Subscriber, value, channel mix dan geographic contribution.")
    sh,x=source_data("Halo",["halo","price"],["region","branch","cluster","city","date"],keywords=["rev halo"],filters=filters)
    if sh:
        subs,price,city,date=[detect_metric(x,"halo"),detect_metric(x,"price"),detect_dim(x,"city"),detect_dim(x,"date")]; channel=detect_metric(x,"channel"); flag=find_col(x,"flag") or find_col(x,"status")
        c=st.columns(4); kpi(c[0],"SUBSCRIPTIONS",count(num(safe_series(x,subs)).sum()) if subs else count(len(x)),sh,"📱",trend_text(x,subs,date)); kpi(c[1],"VALUE",money(num(safe_series(x,price)).sum()) if price else "-","detected amount","💰",""); kpi(c[2],"CHANNELS",f"{x[channel].nunique():,}" if channel else "-","unique","🛒",""); kpi(c[3],"ROWS",f"{len(x):,}","filtered","▤","")
        l,r=st.columns(2)
        with l:
            st.markdown('<div class="panel"><div class="panel-title">Halo Subscribers by City</div><div class="panel-sub">Top 20</div>',unsafe_allow_html=True)
            if city and subs:
                g=x.assign(_v=num(safe_series(x,subs))).groupby(city)._v.sum().reset_index().sort_values("_v",ascending=False).head(20); chart(px.bar(g,x="_v",y=city,orientation="h",template="plotly_white",title=""),400)
            st.markdown('</div>',unsafe_allow_html=True)
        with r:
            st.markdown('<div class="panel"><div class="panel-title">Activation Channel Mix</div><div class="panel-sub">Share of records by channel</div>',unsafe_allow_html=True)
            if channel:
                g=x[channel].astype(str).replace("nan","Unknown").value_counts().reset_index(); g.columns=["Channel","Count"]; chart(px.pie(g,names="Channel",values="Count",hole=.6,template="plotly_white",title=""),400)
            else: st.info("Channel belum terdeteksi.")
            st.markdown('</div>',unsafe_allow_html=True)
        st.divider(); detected([("Subs",subs),("Price",price),("Channel",channel),("Flag/Status",flag),("Region",detect_dim(x,"region")),("Branch",detect_dim(x,"branch")),("Cluster",detect_dim(x,"cluster")),("City",city),("Date",date)])
    else: st.warning("Halo Dashboard.xlsx belum terbaca.")

elif page=="🎓 Skul.id":
    title("Skul.id — Digital Adoption","School coverage, users, active, productive dan conversion funnel.")
    candidates=[]
    for s,df in B.get("Skul.id",{}).items():
        score=sum(bool(detect_metric(df,m)) for m in ["school_id","users","active","productive","student"])
        if "raw" in norm(s):score+=20
        candidates.append((s,score,len(df)))
    candidates.sort(key=lambda z:(z[1],z[2]),reverse=True)
    if candidates:
        selected=st.selectbox("Data sheet",[z[0] for z in candidates],key="skul_sheet"); x=apply_filters(B["Skul.id"][selected],filters)
        sid,student,users,active,productive=[detect_metric(x,m) for m in ["school_id","student","users","active","productive"]]; city,date=[detect_dim(x,"city"),detect_dim(x,"date")]
        school_col=detect_metric(x,"school"); schools=x[sid].nunique() if sid else (x[school_col].nunique() if school_col else 0); sv=[num(safe_series(x,c)).sum() if c else np.nan for c in [student,users,active,productive]]
        c=st.columns(5)
        for col,label,val,ic in zip(c,["SCHOOLS","STUDENTS","TOTAL USERS","ACTIVE USERS","PRODUCTIVE"],[schools,*sv],["🏫","👨‍🎓","👥","🟢","⚡"]): kpi(col,label,count(val),selected,ic,"")
        l,r=st.columns(2)
        with l:
            st.markdown('<div class="panel"><div class="panel-title">Adoption Funnel</div><div class="panel-sub">School → users → active → productive</div>',unsafe_allow_html=True)
            fg=pd.DataFrame({"Stage":["Schools","Users","Active","Productive"],"Value":[schools,sv[1],sv[2],sv[3]]}); chart(px.funnel(fg,y="Stage",x="Value",template="plotly_white",title=""),390); st.markdown('</div>',unsafe_allow_html=True)
        with r:
            st.markdown('<div class="panel"><div class="panel-title">Active Users by City</div><div class="panel-sub">Top 20 cities</div>',unsafe_allow_html=True)
            if city and active:
                g=x.assign(_v=num(safe_series(x,active))).groupby(city)._v.sum().reset_index().sort_values("_v",ascending=False).head(20); chart(px.bar(g,x="_v",y=city,orientation="h",template="plotly_white",title=""),390)
            st.markdown('</div>',unsafe_allow_html=True)
        st.divider(); detected([("School ID",sid),("School",school_col),("Student",student),("Total Users",users),("Active",active),("Productive",productive),("Region",detect_dim(x,"region")),("Branch",detect_dim(x,"branch")),("Cluster",detect_dim(x,"cluster")),("City",city),("Date",date)])
        if pd.notna(sv[1]) and sv[1]: st.progress(min(max((sv[2] if pd.notna(sv[2]) else 0)/sv[1],0),1),text=f"Active conversion: {pct((sv[2] if pd.notna(sv[2]) else 0)/sv[1])}")
        if pd.notna(sv[2]) and sv[2]: st.progress(min(max((sv[3] if pd.notna(sv[3]) else 0)/sv[2],0),1),text=f"Productive conversion: {pct((sv[3] if pd.notna(sv[3]) else 0)/sv[2])}")
    else: st.error("Tidak ada sheet Skul.id yang berhasil dibaca.")

elif page=="⚔️ FB Youth":
    title("FB Youth — Competitive Intelligence","TSEL vs IOH vs XL Smart by territory and period.")
    sh,x=source_data("FB Youth",dims=["region","branch","cluster","city"],keywords=["fb","youth"],filters=filters)
    if sh:
        terr=find_col(x,"city") or find_col(x,"region"); opcols=[]
        for key in ["tsel","ioh","xl"]:
            col=None
            for cc in x.columns:
                nn=norm(cc)
                if key in nn and not any(bb in nn for bb in ["flag","growth","target","ach","score"]): col=cc; break
            if col:opcols.append(col)
        if terr and opcols:
            g=x.groupby(terr)[opcols].mean().reset_index()
            for oc in opcols:g[oc]=g[oc].apply(pct_value)
            m=g.melt(id_vars=[terr],var_name="Operator",value_name="Share"); chart(px.bar(m,x="Share",y=terr,color="Operator",orientation="h",barmode="group",template="plotly_white",title=""),500); st.dataframe(g,use_container_width=True,hide_index=True)
        detected([("Territory",terr),("TSEL",opcols[0] if len(opcols)>0 else None),("IOH",opcols[1] if len(opcols)>1 else None),("XL Smart",opcols[2] if len(opcols)>2 else None),("Date",detect_dim(x,"date"))])
    else: st.warning("FB Youth belum terbaca.")

elif page=="🚀 Program Semarak":
    title("Program Semarak — SO to FR","Funnel conversion, city contribution dan monthly movement.")
    sh,x=source_data("Semarak",["so","fr"],["region","branch","cluster","city","date"],keywords=["raw","semarak"],filters=filters)
    if sh:
        soc,frc,city,date=[detect_metric(x,"so"),detect_metric(x,"fr"),detect_dim(x,"city"),detect_dim(x,"date")]; so_v=num(safe_series(x,soc)).sum() if soc else np.nan; fr_v=num(safe_series(x,frc)).sum() if frc else np.nan
        c=st.columns(4); kpi(c[0],"UNIQUE SO",count(so_v),sh,"🧾",trend_text(x,soc,date)); kpi(c[1],"UNIQUE FR",count(fr_v),"detected","⚡",trend_text(x,frc,date)); kpi(c[2],"FR RATE",pct(fr_v/so_v) if pd.notna(fr_v) and pd.notna(so_v) and so_v else "-","FR / SO","🎯",""); kpi(c[3],"ROWS",f"{len(x):,}","filtered","▤","")
        l,r=st.columns(2)
        with l:
            st.markdown('<div class="panel"><div class="panel-title">SO by City</div><div class="panel-sub">Top 20</div>',unsafe_allow_html=True)
            if city and soc:
                g=x.assign(_v=num(safe_series(x,soc))).groupby(city)._v.sum().reset_index().sort_values("_v",ascending=False).head(20); chart(px.bar(g,x="_v",y=city,orientation="h",template="plotly_white",title=""),410)
            st.markdown('</div>',unsafe_allow_html=True)
        with r:
            st.markdown('<div class="panel"><div class="panel-title">SO vs FR Trend</div><div class="panel-sub">Monthly</div>',unsafe_allow_html=True)
            if date and (soc or frc):
                z=pd.DataFrame({"Date":pd.to_datetime(safe_series(x,date),errors="coerce")}).dropna()
                if soc:z["SO"]=num(safe_series(x.loc[z.index],soc))
                if frc:z["FR"]=num(safe_series(x.loc[z.index],frc))
                z["Period"]=z.Date.dt.to_period("M").astype(str); g=z.groupby("Period")[[c for c in ["SO","FR"] if c in z]].sum().reset_index(); chart(px.line(g,x="Period",y=[c for c in ["SO","FR"] if c in g],markers=True,template="plotly_white",title=""),410)
            st.markdown('</div>',unsafe_allow_html=True)
        st.divider(); detected([("SO",soc),("FR",frc),("Region",detect_dim(x,"region")),("Branch",detect_dim(x,"branch")),("Cluster",detect_dim(x,"cluster")),("City",city),("Date",date)])
    else: st.warning("Program Semarak.xlsx belum terbaca.")

else:
    title("Data Explorer","Audit seluruh workbook, sheet, parameter dan sample data.")
    rows=[]
    for src,sh,df in all_frames():
        row={"Source":src,"Sheet":sh,"Rows":len(df),"Columns":len(df.columns)}
        for d in ["region","branch","cluster","city","date"]: row[d.title()]=detect_dim(df,d)
        for m in ["revenue","rgb","halo","price","users","active","productive","school_id","student","so","fr"]: row[m.title()]=detect_metric(df,m)
        rows.append(row)
    if rows:
        diag=pd.DataFrame(rows); st.dataframe(diag,use_container_width=True,hide_index=True); st.download_button("⬇️ Download detection report",diag.to_csv(index=False).encode(),"data_detection_report.csv","text/csv")
        st.divider(); opts=[f"{r['Source']} / {r['Sheet']}" for r in rows]; pick=st.selectbox("Inspect sheet",opts); src,sh=pick.split(" / ",1); df=B[src][sh]; st.markdown(f'<div class="status">{src} • {sh} • {len(df):,} rows × {len(df.columns):,} columns</div>',unsafe_allow_html=True); st.write("**Detected columns**"); detected([(c,c) for c in df.columns]); st.dataframe(df.head(300),use_container_width=True,hide_index=True); st.download_button("⬇️ Download sample",df.head(5000).to_csv(index=False).encode(),"sample.csv","text/csv")
    else: st.error("Tidak ada file Excel yang berhasil dibaca dari data/.")

st.divider(); st.caption(f"Bali Nusra Business Command Center V8 • Last refresh: {datetime.now().strftime('%d %b %Y %H:%M')} • Excel sources: data/")
