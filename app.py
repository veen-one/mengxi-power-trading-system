import hashlib
import re
from datetime import date, datetime
from io import BytesIO

import numpy as np
import pandas as pd
import streamlit as st
from openpyxl import load_workbook

from supabase_backend import (
    all_daily, create_station, daily_rows, delete_daily, get_supabase, healthcheck,
    hourly_rows, month_daily, save_daily, stations, update_station,
)

APP_TITLE = "蒙西新能源多场站日清分与交易决策系统"
st.set_page_config(page_title=APP_TITLE, page_icon="⚡", layout="wide")
RULE_VERSION = "蒙西2026（自动结算V2.5）"
BASELINE_PRICE=282.9; MONTH_LOWER=.90; MONTH_UPPER=1.10; LOWER_FACTOR=1.30; UPPER_FACTOR=1.10
RISK_COMP_BASE=.50; RISK_REC_BASE=1.45; CURVE_LINK=.50

def fnum(v):
    try: return float(str(v).replace(",","")) if v not in (None,"") else 0.0
    except: return 0.0

def wavg(v,w):
    v,w=np.asarray(v,float),np.asarray(w,float); ok=np.isfinite(v)&np.isfinite(w)
    if not ok.any(): return 0.0
    sw=np.sum(w[ok]); return float(np.sum(v[ok]*w[ok])/sw) if sw else float(np.mean(v[ok]))

def detect_date(name,ws):
    texts=[name]
    for row in ws.iter_rows(min_row=1,max_row=min(ws.max_row,20),values_only=True):
        for v in row:
            if isinstance(v,datetime): return v.date()
            if isinstance(v,date): return v
            if isinstance(v,str): texts.append(v)
    for text in texts:
        for p in [r"(20\d{2})[-_/年](\d{1,2})[-_/月](\d{1,2})",r"(20\d{2})(\d{2})(\d{2})"]:
            m=re.search(p,text)
            if m:
                try: return date(*map(int,m.groups()))
                except: pass
    return None

def detect_station(filename):
    df=stations(False); hits=[]; fn=filename.lower()
    for _,r in df.iterrows():
        for text in (r.get("name"),r.get("short_name")):
            if isinstance(text,str) and text.strip() and text.strip().lower() in fn: hits.append((len(text.strip()),int(r.id),r["name"]))
    return (max(hits)[1],max(hits)[2]) if hits else (None,None)

def parse_excel(data,filename):
    wb=load_workbook(BytesIO(data),data_only=True,read_only=True); ws=wb[wb.sheetnames[0]]; trade_date=detect_date(filename,ws); best=None
    for start in range(1,max(2,min(ws.max_row,500)-95)):
        score=sum(sum(isinstance(ws.cell(r,c).value,(int,float)) for c in (3,4,10,11,12)) for r in range(start,start+96))
        if score>=288 and (best is None or score>best[0]): best=(score,start)
    if not best: raise ValueError("未识别到96点数据")
    start=best[1]
    raw=pd.DataFrame([{"lt_power":fnum(ws.cell(r,3).value),"lt_price":fnum(ws.cell(r,4).value),"spot_price":fnum(ws.cell(r,10).value),"actual_power":fnum(ws.cell(r,11).value),"unified_price":fnum(ws.cell(r,12).value)} for r in range(start,start+96)])
    rows=[]
    for h in range(24):
        g=raw.iloc[h*4:(h+1)*4]; le=float(g.lt_power.mean()); lp=float(g.lt_price.mean()); ae=float(g.actual_power.mean()); sp=wavg(g.spot_price,g.actual_power); up=float(g.unified_price.mean()); sf=ae*sp; lf=le*(lp-up)
        rows.append([h+1,le,lp,ae,sp,up,sf,lf,sf+lf])
    hourly=pd.DataFrame(rows,columns=["时点","中长期电量","中长期价","上网电量","实时价","统一结算点价","现货电费","中长期差价电费","电能量合计"])
    summary={"trade_date":trade_date,"lt_energy":hourly["中长期电量"].sum(),"lt_price":wavg(hourly["中长期价"],hourly["中长期电量"]),"actual_energy":hourly["上网电量"].sum(),"spot_price":wavg(hourly["实时价"],hourly["上网电量"]),"unified_price":wavg(hourly["统一结算点价"],hourly["中长期电量"]),"spot_fee":hourly["现货电费"].sum(),"lt_diff_fee":hourly["中长期差价电费"].sum(),"energy_total":hourly["电能量合计"].sum()}
    return summary,hourly,start

def month_summary(df):
    if df.empty:return {}
    x=df.copy()
    for c in ["lt_energy","lt_price","actual_energy","spot_price","unified_price","energy_total"]: x[c]=pd.to_numeric(x.get(c,0),errors="coerce").fillna(0)
    lt=float(x.lt_energy.sum()); act=float(x.actual_energy.sum())
    return {"lt":lt,"actual":act,"coverage":lt/act if act else 0,"lt_price":wavg(x.lt_price,x.lt_energy),"spot_price":wavg(x.spot_price,x.actual_energy),"unified_price":wavg(x.unified_price,x.lt_energy),"energy":float(x.energy_total.sum())}

def selector(key,active=True):
    df=stations(active)
    if df.empty: st.warning("请先创建场站"); return None
    n=st.selectbox("场站",df.name.tolist(),key=key); return df[df.name==n].iloc[0]

def load_input(sid,month):
    d=get_supabase().table("monthly_settlement").select("*").eq("station_id",sid).eq("trade_month",month).limit(1).execute().data or []; return d[0] if d else {}

def save_input(sid,month,payload): get_supabase().table("monthly_settlement").upsert({"station_id":sid,"trade_month":month,"rule_version":RULE_VERSION,**payload},on_conflict="station_id,trade_month").execute()

def auto_settlement(sm,inp):
    lt,act,p_lt,p_spot,energy=sm["lt"],sm["actual"],sm["lt_price"],sm["spot_price"],sm["energy"]
    me=max(0,float(inp.get("mechanism_energy") or 0)); assess=max(0,act-me); cov=lt/assess if assess else 0; lo=assess*MONTH_LOWER; hi=assess*MONTH_UPPER
    upper=max(0,(lt-hi)*(p_lt*UPPER_FACTOR-p_spot)) if assess and cov>MONTH_UPPER and p_lt>p_spot else 0
    lower=max(0,(lo-lt)*(p_spot*LOWER_FACTOR-p_lt)) if assess and cov<MONTH_LOWER and p_spot>p_lt else 0; assessment=max(upper,lower)
    market=float(inp.get("market_bilateral_listing_avg") or 0); curve=float(inp.get("curve_reasonability") or 0); ready=market>0 and curve>0
    congestion=float(inp.get("congestion") or 0); other=float(inp.get("pre_risk_other_fee") or 0); pre=energy+congestion+other-assessment; risk=0
    if ready and market<=BASELINE_PRICE and act>0 and p_lt>0:
        floor=act*p_lt*max(0,RISK_COMP_BASE-(1-curve)*CURVE_LINK); cap=act*p_lt*(RISK_REC_BASE+(1-curve)*CURVE_LINK)
        if pre<floor:risk=floor-pre
        elif pre>cap:risk=cap-pre
    # Q绿电合约直接等于本站中长期合约电量QLT，不再人工录入
    ge=max(0,lt); green_available=max(0,act-me); green=min(ge,green_available); gp=float(inp.get("green_environment_price") or 0); green_fee=green*gp
    mp=float(inp.get("mechanism_price") or BASELINE_PRICE); ms=float(inp.get("mechanism_spot_price") or 0); mechanism=me*(mp-ms) if me>0 and ms!=0 else 0
    regular=float(inp.get("regular_fee") or 0); unit=float(inp.get("unit_fee") or 0); manual=float(inp.get("manual_adjustment") or 0); final=energy+congestion-assessment+risk+green_fee+mechanism+regular-unit+manual
    return {"assessment_actual":assess,"assessment_coverage":cov,"assessment_lower_energy":lo,"assessment_upper_energy":hi,"p_lt":p_lt,"p_unified":p_spot,"p_node":p_spot,"p_regional":p_lt,"upper_assessment":upper,"lower_assessment":lower,"assessment":assessment,"risk_ready":ready,"risk_prevention":risk,"pre_risk_revenue":pre,"green_contract_energy":ge,"green_available_energy":green_available,"green_energy":green,"green_environment_price":gp,"green_fee":green_fee,"mechanism_fee":mechanism,"final_revenue":final,"final_price":final/act if act else 0}

try: healthcheck(); DB_OK=True; DB_ERROR=""
except Exception as e: DB_OK=False; DB_ERROR=str(e)
st.sidebar.title("⚡ 蒙西交易系统"); st.sidebar.caption("Supabase 云数据库版 · 自动结算V2.5"); st.sidebar.success("数据库已连接") if DB_OK else st.sidebar.error("数据库未连接")
page=st.sidebar.radio("导航",["总览","场站管理","日清分上传","月度累计","自动结算","交易决策","数据管理","系统状态"])
if not DB_OK and page!="系统状态": st.error(DB_ERROR); st.stop()

if page=="总览":
    st.title(APP_TITLE); d=all_daily(); st.metric("累计日清分",len(d)); st.dataframe(d.head(30),use_container_width=True,hide_index=True) if not d.empty else None
elif page=="场站管理":
    st.title("场站管理"); df=stations(False)
    with st.form("new"):
        a,b,c=st.columns(3); name=a.text_input("场站名称"); kw=b.text_input("识别关键词"); typ=c.selectbox("类型",["风电","光伏","其他"])
        if st.form_submit_button("创建") and name: create_station({"name":name,"short_name":kw or name,"station_type":typ,"capacity_mw":0,"risk_upper":1.1,"target_lower":.8,"trade_limit_mwh":100,"min_spread":10,"active":True}); st.rerun()
    st.dataframe(df,use_container_width=True,hide_index=True)
elif page=="日清分上传":
    st.title("日清分批量上传"); sdf=stations(); names=sdf.name.tolist(); ids=dict(zip(sdf.name,sdf.id)); files=st.file_uploader("选择多个Excel",type=["xlsx","xlsm"],accept_multiple_files=True); overwrite=st.checkbox("覆盖重复日期")
    parsed=[]; preview=[]
    for f in files or []:
        try:
            data=f.getvalue(); s,h,_=parse_excel(data,f.name); _,n=detect_station(f.name); parsed.append((f,data,s,h,n)); preview.append([f.name,n or "未识别",s.get("trade_date"),s["lt_energy"],s["actual_energy"]])
        except Exception as e: preview.append([f.name,"解析失败",str(e),0,0])
    if preview: st.dataframe(pd.DataFrame(preview,columns=["文件名","场站","日期","中长期电量","上网电量"]),use_container_width=True,hide_index=True)
    if parsed and st.button("🚀 一键全部入库",type="primary"):
        ok=0
        for f,data,s,h,n in parsed:
            if n and s.get("trade_date"):
                try: save_daily(int(ids[n]),f.name,hashlib.md5(data).hexdigest(),s,h,overwrite); ok+=1
                except: pass
        st.success(f"成功入库 {ok}/{len(parsed)}")
elif page=="月度累计":
    st.title("月度累计"); r=selector("m"); month=st.text_input("月份 YYYY-MM",date.today().strftime("%Y-%m"));
    if r is not None:
        df=month_daily(int(r.id),month); sm=month_summary(df)
        if sm: st.dataframe(pd.DataFrame([["上网电量",sm['actual']],["中长期电量",sm['lt']],["中长期均价",sm['lt_price']],["现货均价",sm['spot_price']],["电能量收入",sm['energy']]],columns=["项目","结果"]),use_container_width=True,hide_index=True); st.dataframe(df,use_container_width=True,hide_index=True)
elif page=="自动结算":
    st.title("自动结算｜月度结算表"); r=selector("s"); month=st.text_input("月份 YYYY-MM",date.today().strftime("%Y-%m"),key="sm")
    if r is not None:
        sm=month_summary(month_daily(int(r.id),month))
        if not sm: st.info("该月暂无数据"); st.stop()
        old=load_input(int(r.id),month)
        with st.expander("月度公共参数"):
            with st.form("params"):
                a,b=st.columns(2); market=a.number_input("区内协商/挂牌成交加权均价",value=float(old.get("market_bilateral_listing_avg") or 0)); curve=b.number_input("正式曲线合理度 %",0.,100.,value=float(old.get("curve_reasonability") or 0)*100)/100
                a,b,c=st.columns(3); congestion=a.number_input("阻塞盈余",value=float(old.get("congestion") or 0)); other=b.number_input("风险防范前其他净额",value=float(old.get("pre_risk_other_fee") or 0)); regular=c.number_input("风险防范后其他净额",value=float(old.get("regular_fee") or 0))
                gp=st.number_input("绿色权益价格 元/MWh",value=float(old.get("green_environment_price") or 0),help="绿电合约电量自动取中长期合约电量QLT")
                a,b,c=st.columns(3); me=a.number_input("机制电量 MWh",value=float(old.get("mechanism_energy") or 0)); mp=b.number_input("机制电价",value=float(old.get("mechanism_price") or BASELINE_PRICE)); ms=c.number_input("机制电量对应现货均价",value=float(old.get("mechanism_spot_price") or 0))
                a,b=st.columns(2); unit=a.number_input("机组/两个细则扣费",value=float(old.get("unit_fee") or 0)); manual=b.number_input("人工调整",value=float(old.get("manual_adjustment") or 0))
                if st.form_submit_button("保存并重算",type="primary"): save_input(int(r.id),month,{"market_bilateral_listing_avg":market,"curve_reasonability":curve,"congestion":congestion,"pre_risk_other_fee":other,"regular_fee":regular,"green_environment_price":gp,"mechanism_energy":me,"mechanism_price":mp,"mechanism_spot_price":ms,"unit_fee":unit,"manual_adjustment":manual}); st.rerun()
        calc=auto_settlement(sm,load_input(int(r.id),month)); me=max(0,float(load_input(int(r.id),month).get("mechanism_energy") or 0))
        rows=[
        ["基础","Qactual","日清分月累计",sm['actual'],"MWh"],["基础","Q机制","月度参数",me,"MWh"],["考核","考核基准","Qactual-Q机制",calc['assessment_actual'],"MWh"],["考核","90%下限","考核基准×90%",calc['assessment_lower_energy'],"MWh"],["考核","110%上限","考核基准×110%",calc['assessment_upper_energy'],"MWh"],["考核","QLT","中长期月累计",sm['lt'],"MWh"],["考核","签约率","QLT/(Qactual-Q机制)",calc['assessment_coverage']*100,"%"],["价格","P_LT","本站中长期均价",calc['p_lt'],"元/MWh"],["价格","P统一/P节点","现货均价",calc['p_unified'],"元/MWh"],["考核","上限考核","[QLT-基准×110%]×(1.1P_LT-P统一)",calc['upper_assessment'],"元"],["考核","下限考核","[基准×90%-QLT]×(1.3P节点-P_LT)",calc['lower_assessment'],"元"],["风险","风险防范", "自动计算",calc['risk_prevention'],"元"],["绿电","Q绿电合约","=中长期合约电量QLT",calc['green_contract_energy'],"MWh"],["绿电","可结算电量","Qactual-Q机制",calc['green_available_energy'],"MWh"],["绿电","Q绿电","min(QLT,Qactual-Q机制)",calc['green_energy'],"MWh"],["绿电","绿色权益价格","月度参数",calc['green_environment_price'],"元/MWh"],["绿电","绿电费用","Q绿电×绿色权益价格",calc['green_fee'],"元"],["机制","机制费用","Q机制×(P机制-P机制现货)",calc['mechanism_fee'],"元"],["结果","预计最终收益","自动汇总",calc['final_revenue'],"元"],["结果","最终结算均价","最终收益/Qactual",calc['final_price'],"元/MWh"]]
        out=pd.DataFrame(rows,columns=["类别","结算项目","计算口径/公式","结果","单位"]); out["结果"]=pd.to_numeric(out["结果"],errors="coerce").round(4); st.dataframe(out,use_container_width=True,hide_index=True,height=720)
        st.caption("绿电口径：Q绿电合约=QLT；Q绿电=min(QLT,Qactual-Q机制)；F绿电=Q绿电×绿色权益价格。")
elif page=="交易决策":
    st.title("月内交易决策"); r=selector("d"); st.info("决策模块保留，后续按结算规则继续完善。")
elif page=="数据管理":
    st.title("数据管理"); r=selector("dm",False)
    if r is not None: st.dataframe(daily_rows(int(r.id),500),use_container_width=True,hide_index=True)
elif page=="系统状态":
    st.title("系统状态"); st.success("Supabase 数据库已连接") if DB_OK else st.error(DB_ERROR); st.write("规则版本：",RULE_VERSION); st.write("绿电：Q绿电合约=中长期合约电量QLT；Q绿电=min(QLT,Qactual-Q机制)")