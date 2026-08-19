import hashlib
import re
import sqlite3
from datetime import date, datetime
from io import BytesIO
from pathlib import Path

import numpy as np
import pandas as pd
import streamlit as st
from openpyxl import load_workbook

APP_TITLE = "蒙西新能源多场站日清分与交易决策系统"
DB_PATH = Path(__file__).with_name("mengxi_trading.db")
st.set_page_config(page_title=APP_TITLE, page_icon="⚡", layout="wide")


def conn():
    c = sqlite3.connect(DB_PATH, check_same_thread=False)
    c.row_factory = sqlite3.Row
    return c


def init_db():
    with conn() as c:
        c.executescript("""
        CREATE TABLE IF NOT EXISTS stations(
          id INTEGER PRIMARY KEY AUTOINCREMENT,name TEXT UNIQUE NOT NULL,short_name TEXT,
          station_type TEXT DEFAULT '光伏',capacity_mw REAL DEFAULT 0,active INTEGER DEFAULT 1,
          risk_upper REAL DEFAULT 1.10,target_lower REAL DEFAULT .80,trade_limit_mwh REAL DEFAULT 100,
          min_spread REAL DEFAULT 10,notes TEXT,created_at TEXT DEFAULT CURRENT_TIMESTAMP);
        CREATE TABLE IF NOT EXISTS daily_summary(
          id INTEGER PRIMARY KEY AUTOINCREMENT,station_id INTEGER NOT NULL,trade_date TEXT NOT NULL,
          source_file TEXT,file_hash TEXT,lt_energy REAL DEFAULT 0,lt_price REAL DEFAULT 0,
          actual_energy REAL DEFAULT 0,spot_price REAL DEFAULT 0,unified_price REAL DEFAULT 0,
          spot_fee REAL DEFAULT 0,lt_diff_fee REAL DEFAULT 0,energy_total REAL DEFAULT 0,
          congestion REAL DEFAULT 0,assessment REAL DEFAULT 0,green_fee REAL DEFAULT 0,
          regular_fee REAL DEFAULT 0,mechanism_fee REAL DEFAULT 0,risk_prevention REAL DEFAULT 0,
          unit_fee REAL DEFAULT 0,final_revenue REAL DEFAULT 0,final_price REAL DEFAULT 0,
          uploaded_at TEXT DEFAULT CURRENT_TIMESTAMP,UNIQUE(station_id,trade_date));
        CREATE TABLE IF NOT EXISTS hourly_detail(
          id INTEGER PRIMARY KEY AUTOINCREMENT,station_id INTEGER NOT NULL,trade_date TEXT NOT NULL,
          hour_no INTEGER NOT NULL,lt_energy REAL,lt_price REAL,actual_energy REAL,spot_price REAL,
          unified_price REAL,spot_fee REAL,lt_diff_fee REAL,energy_total REAL,
          UNIQUE(station_id,trade_date,hour_no));
        """)


def qdf(sql, params=()):
    with conn() as c:
        return pd.read_sql_query(sql, c, params=params)


def execute(sql, params=()):
    with conn() as c:
        c.execute(sql, params)
        c.commit()


def stations(active=True):
    return qdf("SELECT * FROM stations" + (" WHERE active=1" if active else "") + " ORDER BY name")


def fnum(v):
    try:
        return float(str(v).replace(",", "")) if v not in (None, "") else 0.0
    except Exception:
        return 0.0


def wavg(v, w):
    v, w = np.asarray(v, float), np.asarray(w, float)
    ok = np.isfinite(v) & np.isfinite(w)
    if not ok.any(): return 0.0
    return float(np.sum(v[ok] * w[ok]) / np.sum(w[ok])) if np.sum(w[ok]) else float(np.mean(v[ok]))


def detect_date(name, ws):
    pats = [r"(20\d{2})[-_/年](\d{1,2})[-_/月](\d{1,2})", r"(20\d{2})(\d{2})(\d{2})"]
    texts = [name]
    for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row, 20), values_only=True):
        for v in row:
            if isinstance(v, (datetime, date)): return v.date() if isinstance(v, datetime) else v
            if isinstance(v, str): texts.append(v)
    for text in texts:
        for p in pats:
            m = re.search(p, text)
            if m:
                try: return date(*map(int, m.groups()))
                except ValueError: pass
    return None


def detect_station(filename):
    df = stations(False)
    hits = []
    for _, r in df.iterrows():
        for text in (r["name"], r["short_name"]):
            if isinstance(text, str) and text and text.lower() in filename.lower():
                hits.append((len(text), int(r["id"]), r["name"]))
    return max(hits, default=(0, None, None))[1:]


def parse_excel(data, filename):
    wb = load_workbook(BytesIO(data), data_only=True, read_only=True)
    ws = wb[wb.sheetnames[0]]
    trade_date = detect_date(filename, ws)
    best = None
    for start in range(1, max(2, min(ws.max_row, 500) - 95)):
        score = 0
        for r in range(start, start + 96):
            score += sum(isinstance(ws.cell(r, c).value, (int, float)) for c in (3, 4, 10, 11, 12))
        if score >= 288 and (best is None or score > best[0]): best = (score, start)
    if not best:
        raise ValueError("未识别到96点数据；当前模板要求 C/D/J/K/L 为中长期电力/中长期价/实时价/实际计量/统一结算点价。")
    start = best[1]
    raw = pd.DataFrame([{
        "lt_power": fnum(ws.cell(r,3).value), "lt_price": fnum(ws.cell(r,4).value),
        "spot_price": fnum(ws.cell(r,10).value), "actual_power": fnum(ws.cell(r,11).value),
        "unified_price": fnum(ws.cell(r,12).value)} for r in range(start, start+96)])
    rows=[]
    for h in range(24):
        g=raw.iloc[h*4:(h+1)*4]
        le=float(g.lt_power.mean()); lp=float(g.lt_price.mean()); ae=float(g.actual_power.mean())
        sp=wavg(g.spot_price,g.actual_power); up=float(g.unified_price.mean())
        sf=ae*sp; lf=le*(lp-up)
        rows.append([h+1,le,lp,ae,sp,up,sf,lf,sf+lf])
    hourly=pd.DataFrame(rows,columns=["时点","中长期电量","中长期价","上网电量","实时价","统一结算点价","现货电费","中长期差价电费","电能量合计"])
    s={"trade_date":trade_date,"lt_energy":hourly["中长期电量"].sum(),"lt_price":wavg(hourly["中长期价"],hourly["中长期电量"]),
       "actual_energy":hourly["上网电量"].sum(),"spot_price":wavg(hourly["实时价"],hourly["上网电量"]),
       "unified_price":wavg(hourly["统一结算点价"],hourly["中长期电量"]),"spot_fee":hourly["现货电费"].sum(),
       "lt_diff_fee":hourly["中长期差价电费"].sum(),"energy_total":hourly["电能量合计"].sum()}
    return s,hourly,start


def save_daily(sid, filename, data, s, hourly, overwrite):
    ds=s["trade_date"].isoformat(); h=hashlib.md5(data).hexdigest()
    with conn() as c:
        old=c.execute("SELECT id FROM daily_summary WHERE station_id=? AND trade_date=?",(sid,ds)).fetchone()
        if old and not overwrite: raise ValueError("该场站该日期已存在；如需重算请勾选覆盖。")
        if old:
            c.execute("DELETE FROM hourly_detail WHERE station_id=? AND trade_date=?",(sid,ds)); c.execute("DELETE FROM daily_summary WHERE station_id=? AND trade_date=?",(sid,ds))
        final=s["energy_total"]; fp=final/s["actual_energy"] if s["actual_energy"] else 0
        c.execute("""INSERT INTO daily_summary(station_id,trade_date,source_file,file_hash,lt_energy,lt_price,actual_energy,spot_price,unified_price,spot_fee,lt_diff_fee,energy_total,final_revenue,final_price) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                  (sid,ds,filename,h,s["lt_energy"],s["lt_price"],s["actual_energy"],s["spot_price"],s["unified_price"],s["spot_fee"],s["lt_diff_fee"],s["energy_total"],final,fp))
        for _,r in hourly.iterrows():
            c.execute("INSERT INTO hourly_detail(station_id,trade_date,hour_no,lt_energy,lt_price,actual_energy,spot_price,unified_price,spot_fee,lt_diff_fee,energy_total) VALUES(?,?,?,?,?,?,?,?,?,?,?)",
                      (sid,ds,int(r["时点"]),r["中长期电量"],r["中长期价"],r["上网电量"],r["实时价"],r["统一结算点价"],r["现货电费"],r["中长期差价电费"],r["电能量合计"]))
        c.commit()


def month_data(sid, month):
    df=qdf("SELECT * FROM daily_summary WHERE station_id=? AND trade_date LIKE ? ORDER BY trade_date",(sid,month+"%"))
    if df.empty: return df,{}
    lt=df.lt_energy.sum(); act=df.actual_energy.sum()
    sm={"lt":lt,"actual":act,"coverage":lt/act if act else 0,"lt_price":wavg(df.lt_price,df.lt_energy),"spot_price":wavg(df.spot_price,df.actual_energy),
        "energy":df.energy_total.sum(),"final":df.final_revenue.sum(),"final_price":df.final_revenue.sum()/act if act else 0}
    return df,sm


init_db()
st.sidebar.title("⚡ 蒙西交易系统")
page=st.sidebar.radio("导航",["总览","场站管理","日清分上传","月度累计","结算费用","交易决策","数据管理"])

if page=="总览":
    st.title(APP_TITLE); sdf=stations(False); daily=qdf("SELECT * FROM daily_summary")
    a,b,c=st.columns(3); a.metric("场站数",len(sdf)); b.metric("累计日清分",len(daily)); c.metric("累计上网电量",f"{daily.actual_energy.sum() if not daily.empty else 0:,.2f} MWh")
    st.info("工作流：创建场站 → 批量上传日清分 → 自动识别/计算 → 月度累计 → 结算费用 → 交易决策。")

elif page=="场站管理":
    st.title("场站管理")
    with st.form("new"):
        c1,c2,c3=st.columns(3); name=c1.text_input("场站名称 *"); short=c2.text_input("简称/文件名关键词"); typ=c3.selectbox("类型",["光伏","风电","其他"])
        c1,c2,c3,c4=st.columns(4); cap=c1.number_input("装机 MW",0.,10000.,50.); upper=c2.number_input("仓位风险上限 %",0.,300.,110.)/100; lower=c3.number_input("目标仓位下限 %",0.,300.,80.)/100; limit=c4.number_input("单次交易上限 MWh",0.,100000.,100.)
        spread=st.number_input("最小价差阈值 元/MWh",0.,1000.,10.)
        if st.form_submit_button("创建场站",type="primary"):
            try: execute("INSERT INTO stations(name,short_name,station_type,capacity_mw,risk_upper,target_lower,trade_limit_mwh,min_spread) VALUES(?,?,?,?,?,?,?,?)",(name.strip(),short.strip(),typ,cap,upper,lower,limit,spread)); st.success("创建成功"); st.rerun()
            except Exception as e: st.error(str(e))
    st.dataframe(stations(False),use_container_width=True,hide_index=True)

elif page=="日清分上传":
    st.title("日清分上传与自动识别"); sdf=stations()
    if sdf.empty: st.warning("请先创建场站"); st.stop()
    names=sdf.name.tolist(); ids=dict(zip(sdf.name,sdf.id)); fs=st.file_uploader("可一次上传多个 Excel",type=["xlsx","xlsm"],accept_multiple_files=True); overwrite=st.checkbox("覆盖同场站同日期数据")
    for i,f in enumerate(fs or []):
        st.divider(); st.subheader(f.name)
        try:
            data=f.getvalue(); s,h,start=parse_excel(data,f.name); _,auto=detect_station(f.name); c1,c2=st.columns(2)
            sn=c1.selectbox("归属场站",names,index=names.index(auto) if auto in names else 0,key=f"s{i}"); s["trade_date"]=c2.date_input("交易日期",s["trade_date"] or date.today(),key=f"d{i}")
            st.caption(f"96点起始行：{start}｜自动识别场站：{auto or '未识别'}")
            cs=st.columns(5); cs[0].metric("中长期电量",f"{s['lt_energy']:.2f}"); cs[1].metric("中长期均价",f"{s['lt_price']:.2f}"); cs[2].metric("上网电量",f"{s['actual_energy']:.2f}"); cs[3].metric("现货均价",f"{s['spot_price']:.2f}"); cs[4].metric("电能量合计",f"{s['energy_total']:,.2f}")
            with st.expander("24时点"): st.dataframe(h,use_container_width=True,hide_index=True)
            if st.button("确认入库",key=f"save{i}",type="primary"):
                save_daily(int(ids[sn]),f.name,data,s,h,overwrite); st.success("已入库")
        except Exception as e: st.error(f"解析失败：{e}")

elif page=="月度累计":
    st.title("月度累计"); sdf=stations()
    if sdf.empty: st.stop()
    c1,c2=st.columns(2); sn=c1.selectbox("场站",sdf.name.tolist()); month=c2.text_input("月份 YYYY-MM",date.today().strftime("%Y-%m")); sid=int(sdf[sdf.name==sn].iloc[0].id); df,sm=month_data(sid,month)
    if df.empty: st.info("暂无数据")
    else:
        cs=st.columns(6); cs[0].metric("累计上网",f"{sm['actual']:,.2f}"); cs[1].metric("累计中长期",f"{sm['lt']:,.2f}"); cs[2].metric("仓位",f"{sm['coverage']:.2%}"); cs[3].metric("中长期均价",f"{sm['lt_price']:.2f}"); cs[4].metric("现货均价",f"{sm['spot_price']:.2f}"); cs[5].metric("最终收益",f"{sm['final']:,.2f}")
        show=df[["trade_date","lt_energy","lt_price","actual_energy","spot_price","energy_total","assessment","risk_prevention","final_revenue","final_price"]]; st.dataframe(show,use_container_width=True,hide_index=True)
        st.download_button("导出CSV",show.to_csv(index=False).encode("utf-8-sig"),f"{sn}_{month}.csv","text/csv")

elif page=="结算费用":
    st.title("结算费用"); sdf=stations()
    if sdf.empty: st.stop()
    sn=st.selectbox("场站",sdf.name.tolist()); sid=int(sdf[sdf.name==sn].iloc[0].id); df=qdf("SELECT * FROM daily_summary WHERE station_id=? ORDER BY trade_date DESC",(sid,))
    if df.empty: st.info("暂无数据"); st.stop()
    ds=st.selectbox("日期",df.trade_date.tolist()); r=df[df.trade_date==ds].iloc[0]
    with st.form("fee"):
        vals=[st.number_input("阻塞盈余",value=float(r.congestion)),st.number_input("考核费用（正数扣除）",value=float(r.assessment)),st.number_input("绿电费用",value=float(r.green_fee)),st.number_input("常规费用",value=float(r.regular_fee)),st.number_input("机制费用",value=float(r.mechanism_fee)),st.number_input("风险防范（补偿正/回收负）",value=float(r.risk_prevention)),st.number_input("机组费用（正数扣除）",value=float(r.unit_fee))]
        if st.form_submit_button("保存并重算",type="primary"):
            final=float(r.energy_total)+vals[0]-vals[1]+vals[2]+vals[3]+vals[4]+vals[5]-vals[6]; fp=final/float(r.actual_energy) if r.actual_energy else 0
            execute("UPDATE daily_summary SET congestion=?,assessment=?,green_fee=?,regular_fee=?,mechanism_fee=?,risk_prevention=?,unit_fee=?,final_revenue=?,final_price=? WHERE station_id=? AND trade_date=?",(*vals,final,fp,sid,ds)); st.success("已保存"); st.rerun()

elif page=="交易决策":
    st.title("交易决策"); sdf=stations()
    if sdf.empty: st.stop()
    c1,c2=st.columns(2); sn=c1.selectbox("场站",sdf.name.tolist()); month=c2.text_input("月份",date.today().strftime("%Y-%m")); sr=sdf[sdf.name==sn].iloc[0]; _,sm=month_data(int(sr.id),month)
    if not sm: sm={"actual":0,"lt":0}
    c1,c2,c3=st.columns(3); p10=c1.number_input("剩余P10 MWh",0.); p50=c2.number_input("剩余P50 MWh",0.); p90=c3.number_input("剩余P90 MWh",0.)
    c1,c2,c3=st.columns(3); remain=c1.number_input("剩余已签中长期 MWh",0.); spot=c2.number_input("预计剩余现货价",0.); trade=c3.number_input("当前拟交易价",0.)
    total_lt=sm["lt"]+remain; spread=trade-spot; rows=[]
    for label,rem in [("P10低发",p10),("P50基准",p50),("P90高发",p90)]:
        total=sm["actual"]+rem; cov=total_lt/total if total else 0; space=max(0,total*float(sr.risk_upper)-total_lt); rows.append([label,total,total_lt,cov,space])
    out=pd.DataFrame(rows,columns=["情景","预计月末电量","预计总合同","仓位","风险上限内可售空间"]); st.dataframe(out.style.format({"仓位":"{:.2%}"}),use_container_width=True)
    p10_space=float(out.iloc[0]["风险上限内可售空间"]); p10_cov=float(out.iloc[0]["仓位"]); buy=max(0,total_lt-(sm["actual"]+p10)*float(sr.risk_upper)); sell=min(p10_space,float(sr.trade_limit_mwh)) if spread>=float(sr.min_spread) else 0
    if buy>0: st.error(f"P10低发情景仓位 {p10_cov:.2%} 超限，优先买回约 {min(buy,float(sr.trade_limit_mwh)):.2f} MWh。")
    elif sell>0: st.success(f"价差 {spread:.2f} 元/MWh 达标且P10仓位安全，可考虑卖出约 {sell:.2f} MWh，边际毛收益约 {sell*spread:,.2f} 元。")
    else: st.info(f"当前价差 {spread:.2f} 元/MWh，建议观望/保留现货。")
    st.warning("考核、曲线合理度、风险防范最终规则需继续按当期蒙西正式规则固化；当前模型用于交易风控辅助。")

else:
    st.title("数据管理"); sdf=stations(False)
    if sdf.empty: st.stop()
    sn=st.selectbox("场站",sdf.name.tolist()); sid=int(sdf[sdf.name==sn].iloc[0].id); df=qdf("SELECT trade_date,source_file,lt_energy,actual_energy,energy_total,final_revenue,uploaded_at FROM daily_summary WHERE station_id=? ORDER BY trade_date DESC",(sid,)); st.dataframe(df,use_container_width=True,hide_index=True)
    if not df.empty:
        ds=st.selectbox("删除日期",df.trade_date.tolist()); ok=st.checkbox("确认删除该日汇总和24时点")
        if st.button("删除",disabled=not ok):
            execute("DELETE FROM hourly_detail WHERE station_id=? AND trade_date=?",(sid,ds)); execute("DELETE FROM daily_summary WHERE station_id=? AND trade_date=?",(sid,ds)); st.success("已删除"); st.rerun()
