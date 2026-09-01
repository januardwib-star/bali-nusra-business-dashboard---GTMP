import os
import re
from datetime import datetime
import numpy as np
import pandas as pd
import streamlit as st
import plotly.express as px

st.set_page_config(page_title='Bali Nusra | Business Command Center', page_icon='📊', layout='wide', initial_sidebar_state='expanded')
DATA_DIR = os.path.join(os.path.dirname(__file__), 'data')
FILES = {
    'Channel': 'Channel Dashboard.xlsx',
    'Halo': 'Halo Dashboard.xlsx',
    'Skul.id': 'Skul.id.xlsx',
    'FB Youth': 'fb_youth_16082026.xlsx',
    'Semarak': 'Program Semarak.xlsx',
}

st.markdown('''<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
html,body,[class*="css"]{font-family:Inter,sans-serif}.block-container{max-width:1600px;padding-top:1rem}
[data-testid="stSidebar"]{border-right:1px solid #e5e7eb}.hero{background:linear-gradient(135deg,#0f172a,#172554 50%,#0f766e);color:#fff;border-radius:24px;padding:28px 32px;margin-bottom:18px;box-shadow:0 14px 40px rgba(15,23,42,.14)}
.hero h1{font-size:32px;font-weight:800;margin:0;letter-spacing:-.04em}.hero p{margin:7px 0 0;color:#dbeafe;font-size:13px}.badge{display:inline-block;margin-top:13px;background:rgba(255,255,255,.13);padding:6px 11px;border-radius:999px;font-size:11px;font-weight:700}
.kpi{background:#fff;border:1px solid #e5e7eb;border-radius:18px;padding:15px 17px;min-height:108px;box-shadow:0 6px 22px rgba(15,23,42,.05)}.kpi .label{font-size:10px;font-weight:800;letter-spacing:.1em;color:#667085;text-transform:uppercase}.kpi .value{font-size:25px;font-weight:800;color:#111827;margin-top:7px}.kpi .note{font-size:11px;color:#667085;margin-top:4px}.section{font-size:20px;font-weight:800;color:#101828;margin:22px 0 8px}.sub{font-size:12px;color:#667085;margin-bottom:12px}.insight{background:#f8fafc;border:1px solid #e5e7eb;border-radius:15px;padding:13px 15px;margin:7px 0}
.detected{font-size:11px;color:#475467;background:#f8fafc;border:1px solid #eaecf0;border-radius:10px;padding:8px 10px;margin-top:8px}
</style>''', unsafe_allow_html=True)

def clean(df):
    if df is None or not isinstance(df, pd.DataFrame) or df.empty: return pd.DataFrame()
    x=df.copy(); seen={}; names=[]
    for i,c in enumerate(x.columns):
        base=str(c).strip()
        if not base or base.casefold() in {'nan','none'}: base=f'Column_{i+1}'
        base=re.sub(r'\s+',' ',base)
        n=seen.get(base.casefold(),0)+1; seen[base.casefold()]=n
        names.append(base if n==1 else f'{base}__{n}')
    x.columns=names
    return x.dropna(how='all').reset_index(drop=True)

def norm(v): return re.sub(r'[^a-z0-9]+','',str(v).strip().casefold())

def safe_series(df,col):
    if df is None or not isinstance(df,pd.DataFrame) or col is None or col not in df.columns: return pd.Series(index=df.index if isinstance(df,pd.DataFrame) else [],dtype='object')
    s=df[col]; return s.iloc[:,0] if isinstance(s,pd.DataFrame) else s

def num(s):
    if isinstance(s,pd.DataFrame): s=s.iloc[:,0]
    if s is None: return pd.Series(dtype=float)
    if pd.api.types.is_numeric_dtype(s): return pd.to_numeric(s,errors='coerce')
    z=s.astype(str).str.strip().str.replace(r'\((.*?)\)',r'-\1',regex=True).str.replace(',','',regex=False).str.replace('%','',regex=False).str.replace('Rp','',regex=False).str.replace(' ','',regex=False)
    return pd.to_numeric(z.str.replace(r'[^0-9eE.\-+]','',regex=True),errors='coerce')

def money(v):
    if v is None or pd.isna(v): return '-'
    v=float(v)
    if abs(v)>=1e9:return f'Rp {v/1e9:.2f} B'
    if abs(v)>=1e6:return f'Rp {v/1e6:.2f} M'
    if abs(v)>=1e3:return f'Rp {v/1e3:.1f} K'
    return f'Rp {v:,.0f}'

def count(v):
    if v is None or pd.isna(v): return '-'
    v=float(v)
    if abs(v)>=1e6:return f'{v/1e6:.2f} M'
    if abs(v)>=1e3:return f'{v/1e3:.1f} K'
    return f'{v:,.0f}'

def pct(v):
    if v is None or pd.isna(v): return '-'
    return f'{float(v)*100:.1f}%'

def pct_value(v):
    if v is None or pd.isna(v): return np.nan
    v=float(v); return v/100 if abs(v)>1.5 else v

DIM={
 'region':['regional','region','region_name','region_lacci','region_sales','regional_name','area_region'],
 'branch':['branch','branch_name','branch_lacci','branch_sales'],
 'cluster':['cluster','cluster_name','cluster_lacci','cluster_sales','new_cluster'],
 'city':['city','city_name','city_lacci','city_name_lacci','kota','kabupaten','sales_territory_value'],
 'date':['date','tanggal','tgl','period','periode','month','bulan','week','minggu','activation_date','tgl_trx','tgl_fr','so_date','regis_date','used_date'],
}
METRICS={
 'revenue':['rev_all_bb','rev_all','revenue_all','revenue','total revenue','total_omzet','total omzet','omzet','sales revenue','total_amount_sellthru'],
 'rgb':['rgb_all','rgb','total rgb','rgb users','paying user rgb all'],
 'halo':['subs','subscription','subscriber','halo subs','halo subscriber','halo'],
 'price':['price','harga paket','amount','nominal','value','total amount'],
 'active':['user active','active users','active_user','users active','active'],
 'productive':['user productive','productive users','productive_user','productive'],
 'users':['total users','total user','users','user count'],
 'school':['school','schools','sekolah','school count','jumlah sekolah','active school','register school'],
 'school_id':['id sekolah','school id','school_id','npsn'],
 'student':['student','students','siswa','student count','total std tch'],
 'so':['uniq_so','unique so','total so','sales order','so (unik)','so'],
 'fr':['uniq_fr','unique fr','first recharge','fr sukses','fr'],
}

def find_col(df,aliases,exclude=None):
    if df is None or df.empty:return None
    cols=list(df.columns); exclude={norm(x) for x in (exclude or [])}
    exact={norm(c):c for c in cols if norm(c) not in exclude}
    for a in aliases:
        k=norm(a)
        if k in exact:return exact[k]
    for a in sorted(aliases,key=lambda z:len(norm(z)),reverse=True):
        k=norm(a)
        if not k:continue
        for c in cols:
            nc=norm(c)
            if nc in exclude:continue
            if k in nc or nc in k:
                if k in {'active','productive','users','school','so','fr','halo','price','revenue','rgb'} and any(bad in nc for bad in ['flag','growth','ach','target','weight','score']): continue
                return c
    return None

def detect_dim(df,d): return find_col(df,DIM.get(d,[d]))
def detect_metric(df,m): return find_col(df,METRICS.get(m,[m]))

def header_score(row):
    vals=[str(v).strip() for v in row.tolist() if pd.notna(v)]
    if not vals:return -999
    tokens=['region','regional','branch','cluster','city','kota','kabupaten','territory','user active','user productive','total users','revenue','rev','rgb','subs','price','uniq so','uniq fr','status fr','channel','tanggal','date','school','sekolah','siswa','operator','tsel','ioh','xl','trx','transaction']
    hit=sum(any(t in norm(v) for t in tokens) for v in vals)
    text=sum(not bool(re.fullmatch(r'[-+]?\d+(\.\d+)?',v)) for v in vals)/max(len(vals),1)
    numeric=sum(bool(re.fullmatch(r'[-+]?\d+(\.\d+)?',v)) for v in vals)/max(len(vals),1)
    return hit*5+text*1.5-numeric*2-len([v for v in vals if v.lower().startswith('unnamed')])*.5

def merged_channel(path,sheet):
    raw=pd.read_excel(path,sheet_name=sheet,header=None)
    if len(raw)<3:return pd.DataFrame()
    r1=raw.iloc[1]; r2=raw.iloc[2]; names=[]
    for i in range(raw.shape[1]):
        a='' if pd.isna(r1.iloc[i]) else str(r1.iloc[i]).strip(); b='' if pd.isna(r2.iloc[i]) else str(r2.iloc[i]).strip()
        name=a or b
        if a and b and i>=4 and a.casefold() not in {'kpi params calculation','omzet outlet','final score'}: name=b
        names.append(name or f'Column_{i+1}')
    x=raw.iloc[3:].copy(); x.columns=names; return clean(x)

@st.cache_data(show_spinner=False)
def read_book(path):
    out={}
    if not os.path.exists(path):return out
    try: xl=pd.ExcelFile(path)
    except Exception:return out
    for sheet in xl.sheet_names:
        try:
            if os.path.basename(path)=='Channel Dashboard.xlsx' and str(sheet) in {'202607','202608'}: x=merged_channel(path,sheet)
            else:
                raw=pd.read_excel(path,sheet_name=sheet,header=None,nrows=12)
                if raw.empty:continue
                scores=[header_score(raw.iloc[r]) for r in range(min(8,len(raw)))]
                h=int(np.argmax(scores)); x=pd.read_excel(path,sheet_name=sheet,header=h); x=clean(x)
            if not x.empty:out[str(sheet).strip()]=x
        except Exception: continue
    return out

@st.cache_data(show_spinner=False)
def load_all(): return {k:read_book(os.path.join(DATA_DIR,f)) for k,f in FILES.items()}
B=load_all()

def all_frames(source=None):
    for src in ([source] if source else list(B)):
        for sheet,df in B.get(src,{}).items():
            if isinstance(df,pd.DataFrame) and not df.empty:yield src,sheet,df

def unique_values(df,col):
    if df is None or df.empty or col is None:return []
    s=safe_series(df,col).dropna().astype(str).str.strip(); bad={'','nan','none','null','-','all','unknown'}
    return sorted([v for v in s.unique() if v.casefold() not in bad],key=str.casefold)

def apply_filters(df,filters):
    x=clean(df)
    for dim,val in (filters or {}).items():
        if val in (None,'All') or x.empty:continue
        c=detect_dim(x,dim)
        if c:x=x[safe_series(x,c).astype(str).str.strip().str.casefold()==str(val).strip().casefold()]
    return x

def dimension_values(dim,filters=None):
    vals=set(); filters=filters or {}
    for src in ['Skul.id','Channel','Halo','Semarak','FB Youth']:
        for _,_,df in all_frames(src):
            x=apply_filters(df,{k:v for k,v in filters.items() if k!=dim}); c=detect_dim(x,dim)
            if c:vals.update(unique_values(x,c))
    return sorted(vals,key=str.casefold)

def best_sheet(source,metrics=None,dims=None,keywords=None):
    metrics=metrics or []; dims=dims or []; keys=[norm(k) for k in (keywords or [])]; best=None; bestscore=-1e9
    for sheet,df in B.get(source,{}).items():
        score=0; sn=norm(sheet)
        for m in metrics:
            if detect_metric(df,m):score+=10
        for d in dims:
            if detect_dim(df,d):score+=3
        score+=sum(5 for k in keys if k and k in sn)
        if source=='Skul.id':
            if 'rawdata' in sn:score+=20
            if 'summary' in sn:score-=3
        if source=='Semarak':
            if 'rawsemarak' in sn or 'rawallnew' in sn:score+=20
            if 'dashboard' in sn or 'summary' in sn:score-=5
        if source=='Channel':
            if 'revall' in sn and 'revenue' in metrics:score+=25
            if sn=='rgb' and 'rgb' in metrics:score+=25
        if source=='Halo' and 'halo' in metrics:
            if 'revhaloall' in sn:score+=25
            if 'device' in sn:score-=5
        score+=min(len(df),100000)/100000+min(len(df.columns),80)*.02
        if score>bestscore:bestscore,best=score,sheet
    return best

def source_data(source,metrics=None,dims=None,keywords=None,filters=None):
    sh=best_sheet(source,metrics,dims,keywords); return (sh,apply_filters(B[source][sh],filters or {})) if sh else (None,None)

def sum_metric(source,metric,filters):
    sh=best_sheet(source,[metric],['region','branch','cluster','city'])
    if not sh:return np.nan
    x=apply_filters(B[source][sh],filters); c=detect_metric(x,metric)
    if not c:return np.nan
    v=num(safe_series(x,c)).sum(min_count=1); return float(v) if pd.notna(v) else np.nan

def show_chart(fig,height=380):
    fig.update_layout(height=height,margin=dict(l=10,r=10,t=55,b=10),paper_bgcolor='rgba(0,0,0,0)',plot_bgcolor='rgba(0,0,0,0)',font=dict(family='Inter',color='#344054'),legend=dict(orientation='h',yanchor='bottom',y=1.02,x=0),hovermode='x unified')
    fig.update_xaxes(showgrid=True,gridcolor='#eef2f6'); fig.update_yaxes(showgrid=False); st.plotly_chart(fig,use_container_width=True,config={'displaylogo':False})

def kpi(c,label,value,note=''): c.markdown(f'<div class="kpi"><div class="label">{label}</div><div class="value">{value}</div><div class="note">{note}</div></div>',unsafe_allow_html=True)
def title(t,s=None):
    st.markdown(f'<div class="section">{t}</div>',unsafe_allow_html=True)
    if s:st.markdown(f'<div class="sub">{s}</div>',unsafe_allow_html=True)
def detected_table(df,items): st.dataframe(pd.DataFrame({'Parameter':[x[0] for x in items],'Detected Column':[x[1] for x in items]}),use_container_width=True,hide_index=True)

with st.sidebar:
    st.markdown('## 📊 BALI NUSRA'); st.caption('Business Command Center • V7')
    if st.button('🔄 Refresh data',use_container_width=True):st.cache_data.clear();st.rerun()
    st.divider(); page=st.radio('NAVIGATION',['🏠 Executive Overview','📈 Channel Performance','💰 Revenue & RGB','📱 Halo & Device','🎓 Skul.id','⚔️ FB Youth','🚀 Program Semarak','🗂️ Data Explorer'])
    st.divider(); st.caption('DATA SOURCES')
    for key,fn in FILES.items():
        sh=B.get(key,{})
        st.write(f"{'🟢' if sh else '🔴'} {key} · {len(sh)} sheet" if sh else f'🔴 {key} · file tidak terbaca')

st.markdown("<div class='hero'><h1>Bali Nusra Business Command Center</h1><p>Executive performance • revenue • channel • digital adoption • competitive share • program effectiveness</p><div class='badge'>LIVE EXCEL • AUTO SHEET/COLUMN DETECTION • CASCADING FILTERS</div></div>",unsafe_allow_html=True)
filters={}; f1,f2,f3,f4=st.columns(4)
with f1:filters['region']=st.selectbox('Region',['All']+dimension_values('region'),key='filter_region')
with f2:filters['branch']=st.selectbox('Branch',['All']+dimension_values('branch',{'region':filters['region']}),key='filter_branch')
with f3:filters['cluster']=st.selectbox('Cluster',['All']+dimension_values('cluster',{'region':filters['region'],'branch':filters['branch']}),key='filter_cluster')
with f4:filters['city']=st.selectbox('City / Kabupaten',['All']+dimension_values('city',{'region':filters['region'],'branch':filters['branch'],'cluster':filters['cluster']}),key='filter_city')

if page=='🏠 Executive Overview':
    title('Executive Scorecard','KPI menggunakan sumber Excel yang paling relevan, bukan sekadar sheet pertama.')
    revenue=sum_metric('Channel','revenue',filters); rgb=sum_metric('Channel','rgb',filters); halo=sum_metric('Halo','halo',filters); active=sum_metric('Skul.id','active',filters); productive=sum_metric('Skul.id','productive',filters); so=sum_metric('Semarak','so',filters); fr=sum_metric('Semarak','fr',filters)
    cards=st.columns(7)
    for c,l,v,n in zip(cards,['REVENUE','RGB','HALO SUBS','SKUL ACTIVE','SKUL PRODUCTIVE','FR','FR RATE'],[money(revenue),count(rgb),count(halo),count(active),count(productive),count(fr),pct(fr/so) if pd.notna(fr) and pd.notna(so) and so else '-'],['Channel','Channel','Halo','Skul.id','Skul.id','Semarak','FR / SO']):kpi(c,l,v,n)
    title('Business Pulse','Filter wilayah di atas berlaku ke seluruh modul bila kolom wilayah tersedia.')
    l,r=st.columns([1.35,1])
    with l:
        sh,x=source_data('Channel',['revenue'],['city','date'],filters=filters)
        if sh:
            rev,city,date=detect_metric(x,'revenue'),detect_dim(x,'city'),detect_dim(x,'date')
            if rev and city:g=x.assign(_v=num(safe_series(x,rev))).groupby(city)['_v'].sum().reset_index().sort_values('_v',ascending=False).head(15);show_chart(px.bar(g,x='_v',y=city,orientation='h',text='_v',title=f'Revenue by City • {sh}',template='plotly_white'))
    with r:
        sh,x=source_data('FB Youth',dims=['region','branch','cluster','city'],keywords=['fb','youth'],filters=filters)
        if sh:
            terr=find_col(x,['TERITORI','territory','city','kota','kabupaten']); ops=[c for c in [find_col(x,['TSEL']),find_col(x,['IOH']),find_col(x,['XL','XL+','XLSMART'])] if c]
            if terr and ops:
                g=x.groupby(terr)[ops].mean().reset_index(); [g.__setitem__(c,g[c].apply(pct_value)) for c in ops]; m=g.melt(id_vars=[terr],var_name='Operator',value_name='Share'); show_chart(px.bar(m,x='Share',y=terr,color='Operator',orientation='h',barmode='group',title='Competitive Share',template='plotly_white'))
    title('Management Signals')
    if pd.notna(active) and pd.notna(productive) and active:st.info(f'🎓 Skul productivity rate: **{pct(productive/active)}**')
    if pd.notna(so) and so:st.info(f'🚀 Semarak FR conversion: **{pct(fr/so)}**')

elif page=='📈 Channel Performance':
    title('Channel Performance','Revenue, RGB dan outlet/channel mix.'); t1,t2,t3,t4=st.tabs(['Revenue','RGB','Channel Mix','Raw Data'])
    with t1:
        sh,x=source_data('Channel',['revenue'],['region','branch','cluster','city','date'],filters=filters)
        if sh:
            rev,city,date=detect_metric(x,'revenue'),detect_dim(x,'city'),detect_dim(x,'date');a,b,c=st.columns(3);kpi(a,'REVENUE',money(num(safe_series(x,rev)).sum()) if rev else '-',sh);kpi(b,'ROWS',f'{len(x):,}','filtered');kpi(c,'CITY',f'{x[city].nunique():,}' if city else '-','unique')
            if rev and city:g=x.assign(_v=num(safe_series(x,rev))).groupby(city)['_v'].sum().reset_index().sort_values('_v',ascending=False).head(20);show_chart(px.bar(g,x='_v',y=city,orientation='h',title='Revenue by City',template='plotly_white'))
            if rev and date:
                z=x.copy();z['_date']=pd.to_datetime(safe_series(z,date),errors='coerce');z=z.dropna(subset=['_date']);z['_v']=num(safe_series(z,rev))
                if not z.empty:g=z.groupby(z['_date'].dt.to_period('M').astype(str))['_v'].sum().reset_index(name='Revenue');show_chart(px.line(g,x=g.columns[0],y='Revenue',markers=True,title='Revenue Trend',template='plotly_white'))
            detected_table(x,[('Revenue',rev),('City',city),('Date',date)])
    with t2:
        sh,x=source_data('Channel',['rgb'],['region','branch','cluster','city','date'],filters=filters)
        if sh:
            rgb,city,date=detect_metric(x,'rgb'),detect_dim(x,'city'),detect_dim(x,'date');kpi(st.columns(3)[0],'RGB',count(num(safe_series(x,rgb)).sum()) if rgb else '-',sh)
            if rgb and city:g=x.assign(_v=num(safe_series(x,rgb))).groupby(city)['_v'].sum().reset_index().sort_values('_v',ascending=False).head(20);show_chart(px.bar(g,x='_v',y=city,orientation='h',title='RGB by City',template='plotly_white'))
            detected_table(x,[('RGB',rgb),('City',city),('Date',date)])
    with t3:
        sh,x=source_data('Halo',['halo'],['region','branch','cluster','city'],keywords=['rev halo'],filters=filters)
        if sh:
            ch=find_col(x,['channel','Channel (standar dashboard)','channel standard'])
            if ch:g=x[ch].astype(str).value_counts().reset_index();g.columns=['Channel','Count'];show_chart(px.pie(g,names='Channel',values='Count',hole=.58,title='Activation Mix',template='plotly_white'))
            detected_table(x,[('Channel',ch),('Halo Subs',detect_metric(x,'halo')),('Price',detect_metric(x,'price'))])
    with t4:
        sh,x=source_data('Channel',['revenue'],['city'],filters=filters)
        if sh:st.caption(f'Sheet: {sh}');st.dataframe(x.head(1000),use_container_width=True,hide_index=True);st.download_button('⬇️ Download CSV',x.to_csv(index=False).encode(),'channel_filtered.csv','text/csv')

elif page=='💰 Revenue & RGB':
    title('Revenue & RGB','Contribution per city and trend.'); sh,x=source_data('Channel',['revenue','rgb'],['region','branch','cluster','city','date'],filters=filters)
    if sh:
        rev,rgb,city,date=detect_metric(x,'revenue'),detect_metric(x,'rgb'),detect_dim(x,'city'),detect_dim(x,'date');a,b,c,d=st.columns(4);kpi(a,'REVENUE',money(num(safe_series(x,rev)).sum()) if rev else '-',sh);kpi(b,'RGB',count(num(safe_series(x,rgb)).sum()) if rgb else '-','Channel');kpi(c,'TOP CITY',str(x.assign(_v=num(safe_series(x,rev))).groupby(city)['_v'].sum().idxmax()) if city and rev and not x.empty else '-','revenue');kpi(d,'ROWS',f'{len(x):,}','filtered')
        l,r=st.columns(2)
        if city and rev:
            with l:g=x.assign(_v=num(safe_series(x,rev))).groupby(city)['_v'].sum().reset_index().sort_values('_v',ascending=False).head(15);show_chart(px.bar(g,x='_v',y=city,orientation='h',title='Revenue Concentration',template='plotly_white'))
        if city and rgb:
            with r:g=x.assign(_v=num(safe_series(x,rgb))).groupby(city)['_v'].sum().reset_index().sort_values('_v',ascending=False).head(15);show_chart(px.bar(g,x='_v',y=city,orientation='h',title='RGB by City',template='plotly_white'))
        detected_table(x,[('Revenue',rev),('RGB',rgb),('City',city),('Date',date)])

elif page=='📱 Halo & Device':
    title('Halo & Device','Activation, subscription, price, channel dan status.'); sh,x=source_data('Halo',['halo'],['region','branch','cluster','city','date'],keywords=['rev halo'],filters=filters)
    if sh:
        subs,price,city,date=detect_metric(x,'halo'),detect_metric(x,'price'),detect_dim(x,'city'),detect_dim(x,'date');channel=find_col(x,['channel','Channel (standar dashboard)','channel standard']);flag=find_col(x,['flag','status']);a,b,c,d=st.columns(4);kpi(a,'SUBSCRIPTIONS',count(num(safe_series(x,subs)).sum()) if subs else count(len(x)),sh);kpi(b,'VALUE / PRICE',money(num(safe_series(x,price)).sum()) if price else '-','auto');kpi(c,'CHANNELS',f'{x[channel].nunique():,}' if channel else '-','unique');kpi(d,'ROWS',f'{len(x):,}','filtered')
        if city and subs:g=x.assign(_v=num(safe_series(x,subs))).groupby(city)['_v'].sum().reset_index().sort_values('_v',ascending=False).head(20);show_chart(px.bar(g,x='_v',y=city,orientation='h',title='Halo Base by City',template='plotly_white'))
        if channel:g=x[channel].astype(str).value_counts().reset_index();g.columns=['Channel','Count'];show_chart(px.pie(g,names='Channel',values='Count',hole=.58,title='Activation by Channel',template='plotly_white'))
        detected_table(x,[('Subs',subs),('Price',price),('City',city),('Date',date),('Channel',channel),('Flag',flag),('Regional',detect_dim(x,'region')),('Branch',detect_dim(x,'branch')),('Cluster',detect_dim(x,'cluster'))])

elif page=='🎓 Skul.id':
    title('Skul.id — Digital Adoption Funnel','Menggunakan Raw Data terbaru dan mendeteksi parameter sekolah, user, active, productive serta wilayah.')
    candidates=[]
    for s,df in B.get('Skul.id',{}).items():
        score=sum(bool(detect_metric(df,m)) for m in ['school_id','users','active','productive','student'])
        if 'rawdata' in norm(s):score+=20
        if score>=3:candidates.append((s,score,len(df)))
    candidates.sort(key=lambda z:(z[1],z[2]),reverse=True)
    if candidates:
        selected=st.selectbox('Periode / Sheet Raw Skul.id',[x[0] for x in candidates],key='skul_sheet');x=apply_filters(B['Skul.id'][selected],filters)
        school_id,student,users,active,productive=[detect_metric(x,m) for m in ['school_id','student','users','active','productive']];city,date=detect_dim(x,'city'),detect_dim(x,'date');region,branch,cluster=detect_dim(x,'region'),detect_dim(x,'branch'),detect_dim(x,'cluster')
        schools=x[school_id].nunique() if school_id else 0;vals=[num(safe_series(x,c)).sum() if c else np.nan for c in [student,users,active,productive]];cards=st.columns(5)
        for c,l,v,n in zip(cards,['SCHOOLS','STUDENTS','TOTAL USERS','ACTIVE USERS','PRODUCTIVE'],[schools,*vals],['unique','detected','detected','detected','detected']):kpi(c,l,count(v),n)
        if pd.notna(vals[1]) and vals[1]:st.progress(min(max(vals[2]/vals[1],0),1),text=f'Active conversion: {pct(vals[2]/vals[1])}')
        if pd.notna(vals[2]) and vals[2]:st.progress(min(max(vals[3]/vals[2],0),1),text=f'Productive conversion: {pct(vals[3]/vals[2])}')
        l,r=st.columns(2)
        if city and active:
            with l:g=x.assign(_v=num(safe_series(x,active))).groupby(city)['_v'].sum().reset_index().sort_values('_v',ascending=False).head(20);show_chart(px.bar(g,x='_v',y=city,orientation='h',title='Active Users by City',template='plotly_white'))
        if city and active and productive:
            with r:g=x.assign(Active=num(safe_series(x,active)),Productive=num(safe_series(x,productive))).groupby(city)[['Active','Productive']].sum().reset_index();g['Productive Rate']=np.where(g.Active>0,g.Productive/g.Active,0);show_chart(px.bar(g.sort_values('Productive Rate',ascending=False).head(20),x='Productive Rate',y=city,orientation='h',title='Productive Rate by City',template='plotly_white'))
        detected_table(x,[('School ID',school_id),('Student',student),('Total Users',users),('Active',active),('Productive',productive),('Region',region),('Branch',branch),('Cluster',cluster),('City',city),('Date',date)])
        st.success(f'Sheet terdeteksi: {selected}')
    else:st.error('Raw Data Skul.id belum dapat dipetakan. Buka Data Explorer untuk melihat semua sheet.')

elif page=='⚔️ FB Youth':
    title('FB Youth — Competitive Intelligence','TSEL vs IOH vs XL per territory dan periode.'); sh,x=source_data('FB Youth',dims=['region','branch','cluster','city'],keywords=['fb','youth'],filters=filters)
    if sh:
        terr=find_col(x,['TERITORI','territory']);date=find_col(x,['TANGGAL','date']);ops=[c for c in [find_col(x,['TSEL']),find_col(x,['IOH']),find_col(x,['XL','XL+','XLSMART'])] if c]
        if terr and ops:
            g=x.groupby(terr)[ops].mean().reset_index();[g.__setitem__(c,g[c].apply(pct_value)) for c in ops];m=g.melt(id_vars=[terr],var_name='Operator',value_name='Share');show_chart(px.bar(m,x='Share',y=terr,color='Operator',orientation='h',barmode='group',title='Competitive Share by Territory',template='plotly_white'));st.dataframe(g,use_container_width=True,hide_index=True)
        detected_table(x,[('Territory',terr),('Date',date),('TSEL',ops[0] if len(ops)>0 else None),('IOH',ops[1] if len(ops)>1 else None),('XL',ops[2] if len(ops)>2 else None)])

elif page=='🚀 Program Semarak':
    title('Program Semarak — SO to FR','Funnel SO → FR, conversion rate, city dan trend.'); sh,x=source_data('Semarak',['so','fr'],['region','branch','cluster','city','date'],keywords=['raw','semarak'],filters=filters)
    if sh:
        soc,frc,city,date=detect_metric(x,'so'),detect_metric(x,'fr'),detect_dim(x,'city'),detect_dim(x,'date');so_v=num(safe_series(x,soc)).sum() if soc else np.nan;fr_v=num(safe_series(x,frc)).sum() if frc else np.nan;k=st.columns(4);kpi(k[0],'UNIQUE SO',count(so_v),sh);kpi(k[1],'UNIQUE FR',count(fr_v),'detected');kpi(k[2],'FR RATE',pct(fr_v/so_v) if pd.notna(so_v) and so_v else '-','FR / SO');kpi(k[3],'ROWS',f'{len(x):,}','filtered')
        if city and soc:g=x.assign(_v=num(safe_series(x,soc))).groupby(city)['_v'].sum().reset_index().sort_values('_v',ascending=False).head(20);show_chart(px.bar(g,x='_v',y=city,orientation='h',title='SO by City',template='plotly_white'))
        if date and (soc or frc):
            z=x.copy();z['_date']=pd.to_datetime(safe_series(z,date),errors='coerce');z=z.dropna(subset=['_date'])
            if not z.empty:
                if soc:z['_so']=num(safe_series(z,soc))
                if frc:z['_fr']=num(safe_series(z,frc))
                cols=[c for c in ['_so','_fr'] if c in z];g=z.groupby(z['_date'].dt.to_period('M').astype(str))[cols].sum().reset_index();g.columns=['Period']+[c.replace('_','').upper() for c in cols];show_chart(px.line(g,x='Period',y=[c for c in g.columns if c!='Period'],markers=True,title='SO vs FR Trend',template='plotly_white'))
        detected_table(x,[('SO',soc),('FR',frc),('City',city),('Date',date),('Region',detect_dim(x,'region')),('Branch',detect_dim(x,'branch')),('Cluster',detect_dim(x,'cluster'))])

else:
    title('Data Explorer','Semua sheet dan parameter yang benar-benar berhasil terbaca.')
    rows=[]
    for src,sh,df in all_frames():
        row={'Source':src,'Sheet':sh,'Rows':len(df),'Columns':len(df.columns)}
        for d in DIM:row[d.title()]=detect_dim(df,d)
        for m in ['revenue','rgb','halo','price','users','active','productive','school_id','student','so','fr']:row[m.title()]=detect_metric(df,m)
        rows.append(row)
    if rows:
        diag=pd.DataFrame(rows);st.dataframe(diag,use_container_width=True,hide_index=True);st.download_button('⬇️ Download detection report',diag.to_csv(index=False).encode(),'data_detection_report.csv','text/csv');st.divider();opts=[f"{r['Source']} / {r['Sheet']}" for r in rows];pick=st.selectbox('Inspect sheet',opts);src,sh=pick.split(' / ',1);df=B[src][sh];st.write(f'**{len(df):,} rows × {len(df.columns):,} columns**');st.code('\n'.join(map(str,df.columns.tolist())));st.dataframe(df.head(200),use_container_width=True,hide_index=True)
    else:st.error('Tidak ada file Excel yang berhasil dibaca dari data/.')

st.divider();st.caption(f"Bali Nusra Business Command Center V7 • Last refresh: {datetime.now().strftime('%d %b %Y %H:%M')} • Excel sources: data/")
