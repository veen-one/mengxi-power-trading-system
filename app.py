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

RULE_VERSION = "蒙西2026（自动结算V2）"
BASELINE_PRICE = 282.9
MONTH_LOWER = 0.90
MONTH_UPPER = 1.10
LOWER_FACTOR = 1.30
UPPER_FACTOR = 1.10
RISK_COMP_BASE = 0.50
RISK_REC_BASE = 1.45
CURVE_LINK = 0.50


def fnum(v):
    try:
        return float(str(v).replace(",", "")) if v not in (None, "") else 0.0
    except Exception:
        return 0.0


def wavg(v, w):
    v, w = np.asarray(v, float), np.asarray(w, float)
    ok = np.isfinite(v) & np.isfinite(w)
    if not ok.any():
        return 0.0
    sw = np.sum(w[ok])
    return float(np.sum(v[ok] * w[ok]) / sw) if sw else float(np.mean(v[ok]))


def money(v):
    return f"{float(v or 0):,.2f}"


def detect_date(name, ws):
    pats = [r"(20\d{2})[-_/年](\d{1,2})[-_/月](\d{1,2})", r"(20\d{2})(\d{2})(\d{2})"]
    texts = [name]
    for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row, 20), values_only=True):
        for v in row:
            if isinstance(v, datetime):
                return v.date()
            if isinstance(v, date):
                return v
            if isinstance(v, str):
                texts.append(v)
    for text in texts:
        for p in pats:
            m = re.search(p, text)
            if m:
                try:
                    return date(*map(int, m.groups()))
                except ValueError:
                    pass
    return None


def detect_station(filename):
    df = stations(False)
    hits = []
    if df.empty:
        return None, None
    fn = filename.lower()
    for _, r in df.iterrows():
        for text in (r.get("name"), r.get("short_name")):
            if isinstance(text, str) and text.strip() and text.lower() in fn:
                hits.append((len(text.strip()), int(r["id"]), r["name"]))
    if not hits:
        return None, None
    _, sid, name = max(hits)
    return sid, name


def parse_excel(data, filename):
    wb = load_workbook(BytesIO(data), data_only=True, read_only=True)
    ws = wb[wb.sheetnames[0]]
    trade_date = detect_date(filename, ws)
    best = None
    max_start = max(2, min(ws.max_row, 500) - 95)
    for start in range(1, max_start):
        score = 0
        for r in range(start, start + 96):
            score += sum(isinstance(ws.cell(r, c).value, (int, float)) for c in (3, 4, 10, 11, 12))
        if score >= 288 and (best is None or score > best[0]):
            best = (score, start)
    if not best:
        raise ValueError("未识别到96点数据；当前模板要求 C/D/J/K/L 为中长期电力/中长期价/实时价/实际计量/统一结算点价。")
    start = best[1]
    raw = pd.DataFrame([
        {
            "point_no": i + 1,
            "lt_power": fnum(ws.cell(r, 3).value),
            "lt_price": fnum(ws.cell(r, 4).value),
            "spot_price": fnum(ws.cell(r, 10).value),
            "actual_power": fnum(ws.cell(r, 11).value),
            "unified_price": fnum(ws.cell(r, 12).value),
        }
        for i, r in enumerate(range(start, start + 96))
    ])
    rows = []
    for h in range(24):
        g = raw.iloc[h * 4:(h + 1) * 4]
        le = float(g.lt_power.mean())
        lp = float(g.lt_price.mean())
        ae = float(g.actual_power.mean())
        sp = wavg(g.spot_price, g.actual_power)
        up = float(g.unified_price.mean())
        sf = ae * sp
        lf = le * (lp - up)
        rows.append([h + 1, le, lp, ae, sp, up, sf, lf, sf + lf])
    hourly = pd.DataFrame(rows, columns=[
        "时点", "中长期电量", "中长期价", "上网电量", "实时价", "统一结算点价",
        "现货电费", "中长期差价电费", "电能量合计"
    ])
    summary = {
        "trade_date": trade_date,
        "lt_energy": hourly["中长期电量"].sum(),
        "lt_price": wavg(hourly["中长期价"], hourly["中长期电量"]),
        "actual_energy": hourly["上网电量"].sum(),
        "spot_price": wavg(hourly["实时价"], hourly["上网电量"]),
        "unified_price": wavg(hourly["统一结算点价"], hourly["中长期电量"]),
        "spot_fee": hourly["现货电费"].sum(),
        "lt_diff_fee": hourly["中长期差价电费"].sum(),
        "energy_total": hourly["电能量合计"].sum(),
    }
    return summary, hourly, raw, start


def month_summary(df):
    if df.empty:
        return {}
    x = df.copy()
    for c in ["lt_energy", "lt_price", "actual_energy", "spot_price", "unified_price", "energy_total", "final_revenue"]:
        if c not in x.columns:
            x[c] = 0.0
        x[c] = pd.to_numeric(x[c], errors="coerce").fillna(0.0)
    lt = float(x.lt_energy.sum())
    act = float(x.actual_energy.sum())
    energy = float(x.energy_total.sum())
    return {
        "lt": lt, "actual": act, "coverage": lt / act if act else 0.0,
        "lt_price": wavg(x.lt_price, x.lt_energy),
        "spot_price": wavg(x.spot_price, x.actual_energy),
        "unified_price": wavg(x.unified_price, x.lt_energy),
        "energy": energy,
    }


def station_selector(key, active=True):
    sdf = stations(active)
    if sdf.empty:
        st.warning("请先在“场站管理”创建场站。")
        return None, sdf
    name = st.selectbox("场站", sdf["name"].tolist(), key=key)
    return sdf[sdf.name == name].iloc[0], sdf


def load_month_input(station_id, month):
    data = get_supabase().table("monthly_settlement").select("*").eq("station_id", station_id).eq("trade_month", month).limit(1).execute().data or []
    return data[0] if data else {}


def save_month_input(station_id, month, payload):
    row = {"station_id": station_id, "trade_month": month, "rule_version": RULE_VERSION, **payload}
    get_supabase().table("monthly_settlement").upsert(row, on_conflict="station_id,trade_month").execute()


def auto_settlement(sm, inp):
    if not sm:
        return {}
    lt, act, lp, sp, up, energy = sm["lt"], sm["actual"], sm["lt_price"], sm["spot_price"], sm["unified_price"], sm["energy"]
    coverage = sm["coverage"]
    upper_assess = 0.0
    if act > 0 and coverage > MONTH_UPPER and lp > up:
        upper_assess = max(0.0, (lt - act * MONTH_UPPER) * (lp * UPPER_FACTOR - up))
    regional_lt = float(inp.get("regional_same_type_lt_price") or 0.0)
    lower_ready = regional_lt > 0
    lower_assess = 0.0
    if lower_ready and act > 0 and coverage < MONTH_LOWER and sp > regional_lt:
        lower_assess = max(0.0, (act * MONTH_LOWER - lt) * (sp * LOWER_FACTOR - regional_lt))
    assessment_month = max(upper_assess, lower_assess if lower_ready else 0.0)
    market_avg = float(inp.get("market_bilateral_listing_avg") or 0.0)
    curve = float(inp.get("curve_reasonability") or 0.0)
    risk_ready = market_avg > 0 and curve > 0
    comp_ratio = max(0.0, RISK_COMP_BASE - (1 - curve) * CURVE_LINK) if risk_ready else 0.0
    rec_ratio = RISK_REC_BASE + (1 - curve) * CURVE_LINK if risk_ready else 0.0
    congestion = float(inp.get("congestion") or 0.0)
    pre_risk_other = float(inp.get("pre_risk_other_fee") or 0.0)
    pre_risk = energy + congestion + pre_risk_other - assessment_month
    risk = 0.0
    if risk_ready and market_avg <= BASELINE_PRICE and act > 0 and lp > 0:
        floor = act * lp * comp_ratio
        cap = act * lp * rec_ratio
        if pre_risk < floor:
            risk = floor - pre_risk
        elif pre_risk > cap:
            risk = cap - pre_risk
    ge = float(inp.get("green_contract_energy") or 0.0)
    gu = float(inp.get("green_user_actual_energy") or 0.0)
    gp = float(inp.get("green_environment_price") or 0.0)
    green_energy = min(ge, act, gu) if ge > 0 and gu > 0 else 0.0
    green_fee = green_energy * gp
    me = float(inp.get("mechanism_energy") or 0.0)
    mp = float(inp.get("mechanism_price") or BASELINE_PRICE)
    ms = float(inp.get("mechanism_spot_price") or 0.0)
    mechanism_fee = me * (mp - ms) if me > 0 and ms != 0 else 0.0
    regular = float(inp.get("regular_fee") or 0.0)
    unit_fee = float(inp.get("unit_fee") or 0.0)
    manual_adjustment = float(inp.get("manual_adjustment") or 0.0)
    final = energy + congestion - assessment_month + risk + green_fee + mechanism_fee + regular - unit_fee + manual_adjustment
    return {
        "upper_assessment": upper_assess, "lower_assessment": lower_assess, "lower_ready": lower_ready,
        "assessment": assessment_month, "risk_ready": risk_ready, "risk_prevention": risk,
        "risk_comp_ratio": comp_ratio, "risk_rec_ratio": rec_ratio, "pre_risk_revenue": pre_risk,
        "green_energy": green_energy, "green_fee": green_fee, "mechanism_fee": mechanism_fee,
        "congestion": congestion, "regular_fee": regular, "unit_fee": unit_fee,
        "manual_adjustment": manual_adjustment, "final_revenue": final,
        "final_price": final / act if act else 0.0,
    }


try:
    healthcheck()
    DB_OK, DB_ERROR = True, ""
except Exception as exc:
    DB_OK, DB_ERROR = False, str(exc)

st.sidebar.title("⚡ 蒙西交易系统")
st.sidebar.caption("Supabase 云数据库版 · 自动结算V2")
st.sidebar.success("数据库已连接") if DB_OK else st.sidebar.error("数据库未连接")
page = st.sidebar.radio("导航", ["总览", "场站管理", "日清分上传", "月度累计", "自动结算", "交易决策", "数据管理", "系统状态"])
if not DB_OK and page != "系统状态":
    st.error("Supabase 未连接，请进入系统状态查看错误。")
    st.stop()

if page == "总览":
    st.title(APP_TITLE)
    sdf, daily = stations(False), all_daily()
    c = st.columns(4)
    c[0].metric("场站数", len(sdf))
    c[1].metric("累计日清分", len(daily))
    c[2].metric("累计上网电量", f"{pd.to_numeric(daily.get('actual_energy', pd.Series(dtype=float)), errors='coerce').fillna(0).sum():,.2f} MWh")
    c[3].metric("结算引擎", RULE_VERSION)
    st.info("创建场站 → 上传日清分 → 自动识别/计算 → 月度累计 → 自动考核/风险防范/机制/绿电 → 交易决策。")
    if not daily.empty:
        st.dataframe(daily.head(30), use_container_width=True, hide_index=True)

elif page == "场站管理":
    st.title("场站管理")
    with st.form("new_station"):
        a,b,c = st.columns(3)
        name = a.text_input("场站名称 *")
        short = b.text_input("识别关键词", help="建议填最稳定的核心词，例如：战壕梁")
        typ = c.selectbox("类型", ["光伏", "风电", "其他"])
        a,b,c,d = st.columns(4)
        cap = a.number_input("装机 MW", 0.0, 10000.0, 50.0)
        upper = b.number_input("仓位风险上限 %", 0.0, 300.0, 110.0) / 100
        lower = c.number_input("目标仓位下限 %", 0.0, 300.0, 80.0) / 100
        limit = d.number_input("单次交易上限 MWh", 0.0, 100000.0, 100.0)
        spread = st.number_input("最小价差阈值 元/MWh", 0.0, 1000.0, 10.0)
        if st.form_submit_button("创建场站", type="primary"):
            if name.strip():
                create_station({"name":name.strip(),"short_name":short.strip() or name.strip(),"station_type":typ,"capacity_mw":cap,"risk_upper":upper,"target_lower":lower,"trade_limit_mwh":limit,"min_spread":spread,"active":True})
                st.success("已创建"); st.rerun()
            else:
                st.error("场站名称不能为空")
    sdf = stations(False)
    if not sdf.empty:
        st.dataframe(sdf, use_container_width=True, hide_index=True)
        with st.expander("编辑场站"):
            en = st.selectbox("场站", sdf.name.tolist())
            r = sdf[sdf.name==en].iloc[0]
            kw = st.text_input("识别关键词", str(r.get("short_name") or r["name"]))
            active = st.checkbox("启用", bool(r.get("active",True)))
            if st.button("保存"):
                update_station(int(r.id), {"short_name":kw.strip(),"active":active})
                st.success("已保存"); st.rerun()

elif page == "日清分上传":
    st.title("日清分上传与自动识别")
    sdf = stations()
    if sdf.empty:
        st.warning("请先创建场站"); st.stop()
    names = sdf.name.tolist(); ids = dict(zip(sdf.name, sdf.id))
    files = st.file_uploader("可一次上传多个 Excel", type=["xlsx","xlsm"], accept_multiple_files=True)
    overwrite = st.checkbox("覆盖同场站同日期数据")
    for i,f in enumerate(files or []):
        st.divider(); st.subheader(f.name)
        try:
            data=f.getvalue(); summary,hourly,raw,start=parse_excel(data,f.name)
            _,auto_name=detect_station(f.name)
            a,b=st.columns(2)
            idx=names.index(auto_name) if auto_name in names else 0
            station_name=a.selectbox("归属场站", names, index=idx, key=f"st_{i}")
            summary["trade_date"]=b.date_input("交易日期", summary["trade_date"] or date.today(), key=f"dt_{i}")
            if auto_name: st.success(f"已自动识别场站：{auto_name}")
            else: st.warning("未自动识别场站，请手动选择；建议在场站管理设置核心关键词。")
            st.caption(f"96点起始行：{start}")
            k=st.columns(5)
            k[0].metric("中长期电量",f"{summary['lt_energy']:.2f}"); k[1].metric("中长期均价",f"{summary['lt_price']:.2f}")
            k[2].metric("上网电量",f"{summary['actual_energy']:.2f}"); k[3].metric("现货均价",f"{summary['spot_price']:.2f}"); k[4].metric("电能量合计",money(summary['energy_total']))
            with st.expander("24时点"): st.dataframe(hourly,use_container_width=True,hide_index=True)
            if st.button("确认入库", key=f"save_{i}", type="primary"):
                save_daily(int(ids[station_name]), f.name, hashlib.md5(data).hexdigest(), summary, hourly, overwrite)
                st.success("已写入 Supabase，月累计和自动结算会同步刷新。")
        except Exception as exc: st.error(f"解析失败：{exc}")

elif page == "月度累计":
    st.title("月度累计")
    row,_=station_selector("month_station")
    if row is None: st.stop()
    month=st.text_input("月份 YYYY-MM", date.today().strftime("%Y-%m"), key="month")
    df=month_daily(int(row.id),month); sm=month_summary(df)
    if not sm: st.info("暂无数据"); st.stop()
    c=st.columns(7)
    c[0].metric("累计上网",f"{sm['actual']:,.2f}"); c[1].metric("累计中长期",f"{sm['lt']:,.2f}"); c[2].metric("仓位",f"{sm['coverage']:.2%}")
    c[3].metric("中长期均价",f"{sm['lt_price']:.2f}"); c[4].metric("现货均价",f"{sm['spot_price']:.2f}"); c[5].metric("统一结算点均价",f"{sm['unified_price']:.2f}"); c[6].metric("电能量收入",money(sm['energy']))
    st.dataframe(df,use_container_width=True,hide_index=True)

elif page == "自动结算":
    st.title("自动结算｜考核、风险防范、绿电、机制")
    st.caption("日常不需要逐日填费用。系统按月累计自动计算；只有全市场公共参数无法从场站日清分获取时，每月录入一次。")
    row,_=station_selector("settle_station")
    if row is None: st.stop()
    month=st.text_input("月份 YYYY-MM", date.today().strftime("%Y-%m"), key="settle_month")
    df=month_daily(int(row.id),month); sm=month_summary(df)
    if not sm: st.info("该月暂无日清分"); st.stop()
    old=load_month_input(int(row.id),month)
    st.subheader("月累计自动识别")
    c=st.columns(6)
    c[0].metric("上网电量",f"{sm['actual']:,.2f}"); c[1].metric("中长期电量",f"{sm['lt']:,.2f}"); c[2].metric("签约率",f"{sm['coverage']:.2%}")
    c[3].metric("中长期均价",f"{sm['lt_price']:.2f}"); c[4].metric("节点现货均价",f"{sm['spot_price']:.2f}"); c[5].metric("统一结算点均价",f"{sm['unified_price']:.2f}")
    with st.expander("月度市场公共参数（通常每月填一次，系统会记住）", expanded=not bool(old)):
        with st.form("market_inputs"):
            a,b,c=st.columns(3)
            regional=a.number_input("所在区域同类型月度中长期均价", value=float(old.get("regional_same_type_lt_price") or 0.0), help="用于签约比例下限考核；这是全市场公共量，不在单场站日清分中。")
            marketavg=b.number_input("区内协商/挂牌成交加权均价", value=float(old.get("market_bilateral_listing_avg") or 0.0), help="风险防范启动条件与282.9比较。")
            curve=c.number_input("正式曲线合理度 %",0.0,100.0,value=float(old.get("curve_reasonability") or 0.0)*100)/100
            a,b,c=st.columns(3)
            congestion=a.number_input("阻塞盈余返还/分摊（元）", value=float(old.get("congestion") or 0.0), help="需要全市场阻塞盈余池和分配系数，无法只靠本站日清分反推。")
            pre_other=b.number_input("风险防范前其他费用净额（元）", value=float(old.get("pre_risk_other_fee") or 0.0))
            regular=c.number_input("风险防范后其他常规净额（元）", value=float(old.get("regular_fee") or 0.0))
            st.markdown("**绿电**")
            a,b,c=st.columns(3)
            ge=a.number_input("绿电合约电量 MWh", value=float(old.get("green_contract_energy") or 0.0)); gu=b.number_input("对应用户实际用电量 MWh", value=float(old.get("green_user_actual_energy") or 0.0)); gp=c.number_input("绿色权益价格 元/MWh", value=float(old.get("green_environment_price") or 0.0))
            st.markdown("**机制电量**")
            a,b,c=st.columns(3)
            me=a.number_input("机制电量 MWh", value=float(old.get("mechanism_energy") or 0.0)); mp=b.number_input("机制电价 元/MWh", value=float(old.get("mechanism_price") or BASELINE_PRICE)); ms=c.number_input("机制电量对应现货均价 元/MWh", value=float(old.get("mechanism_spot_price") or 0.0))
            a,b=st.columns(2)
            unit=a.number_input("机组/两个细则等扣费（元）", value=float(old.get("unit_fee") or 0.0)); manual=b.number_input("人工最终调整（元）", value=float(old.get("manual_adjustment") or 0.0))
            if st.form_submit_button("保存月度参数并重算", type="primary"):
                save_month_input(int(row.id),month,{"regional_same_type_lt_price":regional,"market_bilateral_listing_avg":marketavg,"curve_reasonability":curve,"congestion":congestion,"pre_risk_other_fee":pre_other,"regular_fee":regular,"green_contract_energy":ge,"green_user_actual_energy":gu,"green_environment_price":gp,"mechanism_energy":me,"mechanism_price":mp,"mechanism_spot_price":ms,"unit_fee":unit,"manual_adjustment":manual})
                st.success("已保存，自动重算"); st.rerun()
    inp=load_month_input(int(row.id),month); calc=auto_settlement(sm,inp)
    st.subheader("自动计算结果")
    c=st.columns(5)
    c[0].metric("全月上限考核",money(calc.get('upper_assessment')))
    c[1].metric("全月下限考核",money(calc.get('lower_assessment')) if calc.get('lower_ready') else "待公共参数")
    c[2].metric("自动考核",money(calc.get('assessment')))
    c[3].metric("风险防范",money(calc.get('risk_prevention')) if calc.get('risk_ready') else "待公共参数")
    c[4].metric("预计最终收益",money(calc.get('final_revenue')))
    c=st.columns(4)
    c[0].metric("绿电费用",money(calc.get('green_fee'))); c[1].metric("机制费用",money(calc.get('mechanism_fee'))); c[2].metric("最终结算均价",f"{calc.get('final_price',0):.2f}"); c[3].metric("风险防范前收入",money(calc.get('pre_risk_revenue')))
    st.info(f"全月考核自动按 90%/110% 与 1.3/1.1 系数计算；风险防范基准价 {BASELINE_PRICE} 元/MWh，基础补偿/回收系数 {RISK_COMP_BASE:.0%}/{RISK_REC_BASE:.0%}，并按曲线合理度下降值的 {CURVE_LINK:.0%} 联动。")
    st.warning("分时签约比例考核、阻塞盈余最终返还、两个细则/辅助服务费用依赖交易中心全市场或15分钟级正式结算数据；当前页面不会伪造这些公共量。拿到相应清单后录入一次即可并入最终收益。")
    with st.expander("查看公式明细"):
        st.write({k:v for k,v in calc.items() if isinstance(v,(int,float,bool))})

elif page == "交易决策":
    st.title("月内交易决策")
    row,_=station_selector("decision_station")
    if row is None: st.stop()
    month=st.text_input("月份 YYYY-MM",date.today().strftime("%Y-%m"),key="decision_month")
    sm=month_summary(month_daily(int(row.id),month)) or {"actual":0,"lt":0,"coverage":0,"lt_price":0,"spot_price":0,"energy":0,"unified_price":0}
    st.caption(f"截至已上传日清分：已发 {sm['actual']:.2f} MWh｜已签 {sm['lt']:.2f} MWh｜当前覆盖率 {sm['coverage']:.2%}")
    a,b,c=st.columns(3)
    p10=a.number_input("剩余P10低发预测 MWh",0.0); p50=b.number_input("剩余P50预测 MWh",0.0); p90=c.number_input("剩余P90高发预测 MWh",0.0)
    a,b,c,d=st.columns(4)
    remain_lt=a.number_input("剩余已签中长期 MWh",0.0); expected_spot=b.number_input("预计剩余现货均价",value=float(sm['spot_price'])); trade_price=c.number_input("当前拟交易价格",value=float(sm['lt_price'])); proposed=d.number_input("试算交易量（卖出为正/买回为负）",value=0.0)
    risk_upper=float(row.get("risk_upper") or 1.10); total_lt=sm['lt']+remain_lt+proposed
    rows=[]
    for name,rem in [("P10",p10),("P50",p50),("P90",p90)]:
        total=sm['actual']+rem; cov=total_lt/total if total else 0; space=total*risk_upper-total_lt
        rows.append([name,total,total_lt,cov,space])
    sdf=pd.DataFrame(rows,columns=["情景","预计月底总电量","交易后总合同","仓位","距风险上限空间"])
    st.dataframe(sdf,use_container_width=True,hide_index=True)
    spread=trade_price-expected_spot; p10space=float(sdf.iloc[0]["距风险上限空间"]); limit=float(row.get("trade_limit_mwh") or 100); threshold=float(row.get("min_spread") or 10)
    if p10space<0: suggestion=f"优先买回/降仓，至少 {-p10space:.2f} MWh"
    elif spread>=threshold: suggestion=f"可考虑卖出，单次建议不超过 {min(p10space,limit):.2f} MWh"
    else: suggestion="观望/保留现货"
    st.metric("中长期-预计现货价差",f"{spread:.2f} 元/MWh"); st.success(suggestion)

elif page == "数据管理":
    st.title("数据管理")
    row,_=station_selector("data_station",False)
    if row is None: st.stop()
    df=daily_rows(int(row.id),500)
    st.dataframe(df,use_container_width=True,hide_index=True)
    if not df.empty:
        day=st.selectbox("查看/删除日期",df.trade_date.astype(str).tolist())
        rec=df[df.trade_date.astype(str)==day].iloc[0]
        h=hourly_rows(int(row.id),day)
        with st.expander("24时点明细"): st.dataframe(h,use_container_width=True,hide_index=True)
        if st.button("删除该日数据",type="secondary"):
            delete_daily(int(rec.id),int(row.id),day); st.success("已删除"); st.rerun()

elif page == "系统状态":
    st.title("系统状态")
    if DB_OK:
        st.success("Supabase 数据库已连接")
        st.write("规则版本：",RULE_VERSION)
        st.write("核心表：stations / daily_summary / hourly_detail / monthly_settlement")
    else:
        st.error(DB_ERROR)
