import hashlib
import re
from datetime import date, datetime
from io import BytesIO

import numpy as np
import pandas as pd
import streamlit as st
from openpyxl import load_workbook

from supabase_backend import (
    all_daily,
    create_station,
    daily_rows,
    delete_daily,
    healthcheck,
    hourly_rows,
    month_daily,
    save_daily,
    stations,
    update_settlement,
    update_station,
)

APP_TITLE = "蒙西新能源多场站日清分与交易决策系统"
st.set_page_config(page_title=APP_TITLE, page_icon="⚡", layout="wide")


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
    for _, r in df.iterrows():
        for text in (r.get("name"), r.get("short_name")):
            if isinstance(text, str) and text and text.lower() in filename.lower():
                hits.append((len(text), int(r["id"]), r["name"]))
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
        raise ValueError("未识别到96点数据；当前模板要求 C/D/J/K/L 分别为中长期电力/中长期价/实时价/实际计量/统一结算点价。")

    start = best[1]
    raw = pd.DataFrame([
        {
            "lt_power": fnum(ws.cell(r, 3).value),
            "lt_price": fnum(ws.cell(r, 4).value),
            "spot_price": fnum(ws.cell(r, 10).value),
            "actual_power": fnum(ws.cell(r, 11).value),
            "unified_price": fnum(ws.cell(r, 12).value),
        }
        for r in range(start, start + 96)
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

    hourly = pd.DataFrame(
        rows,
        columns=["时点", "中长期电量", "中长期价", "上网电量", "实时价", "统一结算点价", "现货电费", "中长期差价电费", "电能量合计"],
    )
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
    return summary, hourly, start


def month_summary(df):
    if df.empty:
        return {}
    for c in [
        "lt_energy", "lt_price", "actual_energy", "spot_price", "energy_total", "final_revenue",
        "assessment", "risk_prevention", "congestion", "green_fee", "regular_fee", "mechanism_fee", "unit_fee",
    ]:
        if c not in df.columns:
            df[c] = 0.0
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)
    lt = float(df.lt_energy.sum())
    act = float(df.actual_energy.sum())
    final = float(df.final_revenue.sum())
    return {
        "lt": lt,
        "actual": act,
        "coverage": lt / act if act else 0.0,
        "lt_price": wavg(df.lt_price, df.lt_energy),
        "spot_price": wavg(df.spot_price, df.actual_energy),
        "energy": float(df.energy_total.sum()),
        "assessment": float(df.assessment.sum()),
        "risk": float(df.risk_prevention.sum()),
        "final": final,
        "final_price": final / act if act else 0.0,
    }


def station_selector(key, active=True):
    sdf = stations(active)
    if sdf.empty:
        st.warning("请先在“场站管理”创建场站。")
        return None, sdf
    name = st.selectbox("场站", sdf["name"].tolist(), key=key)
    row = sdf[sdf.name == name].iloc[0]
    return row, sdf


def money(v):
    return f"{float(v):,.2f}"


try:
    healthcheck()
    DB_OK = True
    DB_ERROR = ""
except Exception as exc:
    DB_OK = False
    DB_ERROR = str(exc)

st.sidebar.title("⚡ 蒙西交易系统")
st.sidebar.caption("Supabase 云数据库版")
if DB_OK:
    st.sidebar.success("数据库已连接")
else:
    st.sidebar.error("数据库未连接")

page = st.sidebar.radio(
    "导航",
    ["总览", "场站管理", "日清分上传", "月度累计", "结算费用", "交易决策", "数据管理", "系统状态"],
)

if not DB_OK and page != "系统状态":
    st.error("Supabase 尚未完成连接。请先进入“系统状态”查看配置提示。")
    st.stop()

if page == "总览":
    st.title(APP_TITLE)
    sdf = stations(False)
    daily = all_daily()
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("场站数", len(sdf))
    c2.metric("累计日清分", len(daily))
    c3.metric("累计上网电量", f"{pd.to_numeric(daily.get('actual_energy', pd.Series(dtype=float)), errors='coerce').fillna(0).sum():,.2f} MWh")
    c4.metric("数据库", "Supabase")
    st.info("工作流：创建场站 → 批量上传日清分 → 自动识别/计算 → 月度累计 → 结算费用 → 交易决策。")
    if not daily.empty:
        st.subheader("最近入库")
        st.dataframe(daily.head(20), use_container_width=True, hide_index=True)

elif page == "场站管理":
    st.title("场站管理")
    with st.form("new_station"):
        c1, c2, c3 = st.columns(3)
        name = c1.text_input("场站名称 *")
        short = c2.text_input("简称/文件名关键词")
        typ = c3.selectbox("类型", ["光伏", "风电", "其他"])
        c1, c2, c3, c4 = st.columns(4)
        cap = c1.number_input("装机 MW", 0.0, 10000.0, 50.0)
        upper = c2.number_input("仓位风险上限 %", 0.0, 300.0, 110.0) / 100
        lower = c3.number_input("目标仓位下限 %", 0.0, 300.0, 80.0) / 100
        limit = c4.number_input("单次交易上限 MWh", 0.0, 100000.0, 100.0)
        spread = st.number_input("最小价差阈值 元/MWh", 0.0, 1000.0, 10.0)
        if st.form_submit_button("创建场站", type="primary"):
            if not name.strip():
                st.error("场站名称不能为空")
            else:
                try:
                    create_station({
                        "name": name.strip(), "short_name": short.strip(), "station_type": typ,
                        "capacity_mw": cap, "risk_upper": upper, "target_lower": lower,
                        "trade_limit_mwh": limit, "min_spread": spread, "active": True,
                    })
                    st.success("场站创建成功")
                    st.rerun()
                except Exception as exc:
                    st.error(str(exc))

    sdf = stations(False)
    if not sdf.empty:
        st.subheader("已有场站")
        st.dataframe(sdf, use_container_width=True, hide_index=True)
        with st.expander("编辑场站"):
            edit_name = st.selectbox("选择场站", sdf.name.tolist(), key="edit_station")
            row = sdf[sdf.name == edit_name].iloc[0]
            c1, c2, c3 = st.columns(3)
            short2 = c1.text_input("简称", str(row.get("short_name") or ""))
            cap2 = c2.number_input("装机 MW", 0.0, 10000.0, float(row.get("capacity_mw") or 0))
            active2 = c3.checkbox("启用", bool(row.get("active", True)))
            c1, c2, c3, c4 = st.columns(4)
            up2 = c1.number_input("风险上限 %", 0.0, 300.0, float(row.get("risk_upper") or 1.10) * 100) / 100
            low2 = c2.number_input("目标下限 %", 0.0, 300.0, float(row.get("target_lower") or 0.80) * 100) / 100
            lim2 = c3.number_input("单次上限 MWh", 0.0, 100000.0, float(row.get("trade_limit_mwh") or 100))
            spr2 = c4.number_input("最小价差", 0.0, 1000.0, float(row.get("min_spread") or 10))
            if st.button("保存场站修改"):
                update_station(int(row.id), {
                    "short_name": short2, "capacity_mw": cap2, "active": active2,
                    "risk_upper": up2, "target_lower": low2, "trade_limit_mwh": lim2, "min_spread": spr2,
                })
                st.success("已保存")
                st.rerun()

elif page == "日清分上传":
    st.title("日清分上传与自动识别")
    sdf = stations()
    if sdf.empty:
        st.warning("请先创建场站")
        st.stop()
    names = sdf.name.tolist()
    ids = dict(zip(sdf.name, sdf.id))
    files = st.file_uploader("可一次上传多个 Excel", type=["xlsx", "xlsm"], accept_multiple_files=True)
    overwrite = st.checkbox("覆盖同场站同日期数据")

    for i, f in enumerate(files or []):
        st.divider()
        st.subheader(f.name)
        try:
            data = f.getvalue()
            summary, hourly, start = parse_excel(data, f.name)
            _, auto_name = detect_station(f.name)
            c1, c2 = st.columns(2)
            index = names.index(auto_name) if auto_name in names else 0
            station_name = c1.selectbox("归属场站", names, index=index, key=f"station_{i}")
            summary["trade_date"] = c2.date_input("交易日期", summary["trade_date"] or date.today(), key=f"date_{i}")
            st.caption(f"96点起始行：{start}｜自动识别场站：{auto_name or '未识别'}")
            m = st.columns(5)
            m[0].metric("中长期电量", f"{summary['lt_energy']:.2f}")
            m[1].metric("中长期均价", f"{summary['lt_price']:.2f}")
            m[2].metric("上网电量", f"{summary['actual_energy']:.2f}")
            m[3].metric("现货均价", f"{summary['spot_price']:.2f}")
            m[4].metric("电能量合计", money(summary["energy_total"]))
            with st.expander("查看24时点"):
                st.dataframe(hourly, use_container_width=True, hide_index=True)
            if st.button("确认入库", key=f"save_{i}", type="primary"):
                file_hash = hashlib.md5(data).hexdigest()
                save_daily(int(ids[station_name]), f.name, file_hash, summary, hourly, overwrite)
                st.success("已写入 Supabase")
        except Exception as exc:
            st.error(f"解析失败：{exc}")

elif page == "月度累计":
    st.title("月度累计")
    row, _ = station_selector("month_station")
    if row is None:
        st.stop()
    month = st.text_input("月份 YYYY-MM", date.today().strftime("%Y-%m"))
    df = month_daily(int(row.id), month)
    sm = month_summary(df)
    if df.empty:
        st.info("暂无数据")
    else:
        c = st.columns(7)
        c[0].metric("累计上网", f"{sm['actual']:,.2f}")
        c[1].metric("累计中长期", f"{sm['lt']:,.2f}")
        c[2].metric("仓位", f"{sm['coverage']:.2%}")
        c[3].metric("中长期均价", f"{sm['lt_price']:.2f}")
        c[4].metric("现货均价", f"{sm['spot_price']:.2f}")
        c[5].metric("累计考核", money(sm["assessment"]))
        c[6].metric("最终收益", money(sm["final"]))
        cols = [
            "trade_date", "lt_energy", "lt_price", "actual_energy", "spot_price", "energy_total",
            "assessment", "risk_prevention", "final_revenue", "final_price",
        ]
        show = df[[c for c in cols if c in df.columns]].copy()
        st.dataframe(show, use_container_width=True, hide_index=True)
        st.download_button("导出CSV", show.to_csv(index=False).encode("utf-8-sig"), f"{row['name']}_{month}.csv", "text/csv")

elif page == "结算费用":
    st.title("结算费用录入/修正")
    row, _ = station_selector("fee_station")
    if row is None:
        st.stop()
    month = st.text_input("月份 YYYY-MM", date.today().strftime("%Y-%m"), key="fee_month")
    df = month_daily(int(row.id), month)
    if df.empty:
        st.info("该月暂无日清分")
        st.stop()
    day = st.selectbox("日期", df.trade_date.astype(str).tolist())
    rec = df[df.trade_date.astype(str) == day].iloc[0]
    with st.form("fees"):
        c1, c2, c3, c4 = st.columns(4)
        congestion = c1.number_input("阻塞盈余", value=float(rec.get("congestion") or 0.0))
        assessment = c2.number_input("考核费用（正数扣除）", value=float(rec.get("assessment") or 0.0))
        green_fee = c3.number_input("绿电费用", value=float(rec.get("green_fee") or 0.0))
        regular_fee = c4.number_input("常规费用", value=float(rec.get("regular_fee") or 0.0))
        c1, c2, c3 = st.columns(3)
        mechanism_fee = c1.number_input("机制费用", value=float(rec.get("mechanism_fee") or 0.0))
        risk_prevention = c2.number_input("风险防范（补偿正/回收负）", value=float(rec.get("risk_prevention") or 0.0))
        unit_fee = c3.number_input("机组费用（正数扣除）", value=float(rec.get("unit_fee") or 0.0))
        if st.form_submit_button("保存结算费用", type="primary"):
            energy_total = float(rec.get("energy_total") or 0.0)
            actual = float(rec.get("actual_energy") or 0.0)
            final = energy_total + congestion - assessment + green_fee + regular_fee + mechanism_fee + risk_prevention - unit_fee
            update_settlement(int(rec.id), {
                "congestion": congestion, "assessment": assessment, "green_fee": green_fee,
                "regular_fee": regular_fee, "mechanism_fee": mechanism_fee,
                "risk_prevention": risk_prevention, "unit_fee": unit_fee,
                "final_revenue": final, "final_price": final / actual if actual else 0.0,
            })
            st.success(f"已保存，最终日收益：{money(final)} 元")
            st.rerun()

elif page == "交易决策":
    st.title("月内交易决策")
    row, _ = station_selector("decision_station")
    if row is None:
        st.stop()
    month = st.text_input("月份 YYYY-MM", date.today().strftime("%Y-%m"), key="decision_month")
    df = month_daily(int(row.id), month)
    sm = month_summary(df)
    if not sm:
        sm = {"actual": 0.0, "lt": 0.0, "coverage": 0.0, "lt_price": 0.0, "spot_price": 0.0, "final": 0.0}

    st.subheader("截至目前")
    c = st.columns(5)
    c[0].metric("已发", f"{sm['actual']:,.2f} MWh")
    c[1].metric("累计中长期", f"{sm['lt']:,.2f} MWh")
    c[2].metric("当前仓位", f"{sm['coverage']:.2%}")
    c[3].metric("累计中长期价", f"{sm['lt_price']:.2f}")
    c[4].metric("累计现货价", f"{sm['spot_price']:.2f}")

    st.subheader("剩余月份预测与当前报价")
    c1, c2, c3 = st.columns(3)
    p10 = c1.number_input("剩余P10发电预测 MWh", min_value=0.0, value=700.0)
    p50 = c2.number_input("剩余P50发电预测 MWh", min_value=0.0, value=950.0)
    p90 = c3.number_input("剩余P90发电预测 MWh", min_value=0.0, value=1200.0)
    c1, c2, c3 = st.columns(3)
    remaining_contract = c1.number_input("剩余已签中长期 MWh", min_value=0.0, value=0.0)
    expected_spot = c2.number_input("预计剩余现货均价", value=300.0)
    proposed_price = c3.number_input("当前拟交易中长期价", value=220.0)

    contract_total = sm["lt"] + remaining_contract
    risk_upper = float(row.get("risk_upper") or 1.10)
    target_lower = float(row.get("target_lower") or 0.80)
    trade_limit = float(row.get("trade_limit_mwh") or 100.0)
    min_spread = float(row.get("min_spread") or 10.0)

    scenarios = []
    for label, rem in [("P10低发", p10), ("P50基准", p50), ("P90高发", p90)]:
        total = sm["actual"] + rem
        coverage = contract_total / total if total else 0.0
        limit_energy = total * risk_upper
        room = limit_energy - contract_total
        state = "超限" if room < 0 else ("偏低" if coverage < target_lower else "可控")
        scenarios.append([label, total, contract_total, coverage, limit_energy, room, state])
    scen = pd.DataFrame(scenarios, columns=["情景", "预计月末总电量", "当前+未来已签", "预计月末仓位", "风险上限合同量", "剩余可售空间", "判断"])
    display_scen = scen.copy()
    display_scen["预计月末仓位"] = display_scen["预计月末仓位"].map(lambda x: f"{x:.2%}")
    st.dataframe(display_scen, use_container_width=True, hide_index=True)

    p10_room = float(scen.iloc[0]["剩余可售空间"])
    spread = proposed_price - expected_spot
    sell_qty = min(max(0.0, p10_room), trade_limit) if spread >= min_spread else 0.0
    buyback = min(abs(p10_room), trade_limit) if p10_room < 0 else 0.0
    advice = "优先买回/降仓" if buyback > 0 else ("可考虑卖出" if sell_qty > 0 else ("有仓位但价差不足" if p10_room > 0 else "观望"))

    st.subheader("当前建议")
    c = st.columns(5)
    c[0].metric("拟交易-预计现货价差", f"{spread:.2f} 元/MWh")
    c[1].metric("P10可售空间", f"{p10_room:,.2f} MWh")
    c[2].metric("建议卖出", f"{sell_qty:,.2f} MWh")
    c[3].metric("建议买回", f"{buyback:,.2f} MWh")
    c[4].metric("动作", advice)
    st.caption("决策顺序：先看P10低发仓位 → 再看价差 → 再看单次交易上限。正式考核和风险防范仍以当期规则参数为准。")

elif page == "数据管理":
    st.title("数据管理")
    row, _ = station_selector("data_station", active=False)
    if row is None:
        st.stop()
    df = daily_rows(int(row.id), limit=500)
    if df.empty:
        st.info("暂无数据")
        st.stop()
    st.dataframe(df, use_container_width=True, hide_index=True)
    day = st.selectbox("查看/删除日期", df.trade_date.astype(str).tolist())
    rec = df[df.trade_date.astype(str) == day].iloc[0]
    with st.expander("查看24时点"):
        st.dataframe(hourly_rows(int(row.id), day), use_container_width=True, hide_index=True)
    confirm = st.checkbox("我确认删除该日数据")
    if st.button("删除该日", type="secondary", disabled=not confirm):
        delete_daily(int(rec.id), int(row.id), day)
        st.success("已删除")
        st.rerun()

elif page == "系统状态":
    st.title("系统状态")
    if DB_OK:
        st.success("Supabase 数据库连接正常。")
        st.write("项目 URL 已在后端配置；密钥仅从 Streamlit Secrets 或环境变量读取，不会显示在页面，也不会提交到 GitHub。")
        try:
            st.metric("可读取场站数", len(stations(False)))
        except Exception as exc:
            st.warning(str(exc))
    else:
        st.error("数据库连接失败")
        st.code(DB_ERROR)
        st.markdown(
            """
请检查：

1. Supabase SQL Editor 已执行仓库中的 `supabase_schema.sql`；
2. 本地创建 `.streamlit/secrets.toml`，或在 Streamlit Cloud 的 **Settings → Secrets** 配置：

```toml
SUPABASE_URL = "https://bwhpatzqcklgyzuifvhu.supabase.co"
SUPABASE_KEY = "你的 publishable key"
```

真实 key 不要写入 GitHub 文件。
"""
        )
