from datetime import date

import pandas as pd
import streamlit as st

from supabase_backend import hourly_range, month_daily, stations

st.set_page_config(page_title="价格曲线", page_icon="📈", layout="wide")
st.title("📈 时点价格曲线")
st.caption("按已入库日清分，从当月1日累计到所选截止日期，查看中长期合约价、实时节点电价、全网统一价及中长期-全网统一价差。")

sdf = stations(True)
if sdf.empty:
    st.warning("请先创建并启用场站。")
    st.stop()

c1, c2 = st.columns(2)
station_name = c1.selectbox("场站", sdf["name"].tolist())
month = c2.text_input("月份 YYYY-MM", date.today().strftime("%Y-%m"))
station = sdf[sdf["name"] == station_name].iloc[0]

try:
    daily = month_daily(int(station.id), month)
except Exception as exc:
    st.error(f"月份格式或数据读取失败：{exc}")
    st.stop()

if daily.empty:
    st.info("该场站该月暂无已入库日清分。")
    st.stop()

available_dates = sorted(pd.to_datetime(daily["trade_date"], errors="coerce").dropna().dt.date.unique().tolist())
if not available_dates:
    st.info("未找到有效交易日期。")
    st.stop()

c1, c2, c3 = st.columns(3)
cutoff = c1.selectbox("截止日期", available_dates, index=len(available_dates) - 1, format_func=lambda x: x.strftime("%Y-%m-%d"))
smooth_hours = c2.slider("平滑窗口（小时）", min_value=1, max_value=12, value=3, step=1, help="1=原始24时点曲线；数值越大，曲线越平滑。")
show_raw = c3.checkbox("显示原始数据表", value=False)

start_date = f"{month}-01"
end_date = cutoff.isoformat()
hourly = hourly_range(int(station.id), start_date, end_date)
if hourly.empty:
    st.info("该范围暂无24时点明细。")
    st.stop()

for col in ["hour_no", "lt_price", "spot_price", "unified_price"]:
    hourly[col] = pd.to_numeric(hourly[col], errors="coerce")
hourly["trade_date"] = pd.to_datetime(hourly["trade_date"], errors="coerce")
hourly = hourly.dropna(subset=["trade_date", "hour_no"]).copy()
hourly["时间"] = hourly["trade_date"] + pd.to_timedelta(hourly["hour_no"] - 1, unit="h")
hourly = hourly.sort_values("时间").drop_duplicates(subset=["时间"], keep="last")

hourly["中长期合约价"] = hourly["lt_price"]
hourly["实时节点电价"] = hourly["spot_price"]
hourly["全网统一价"] = hourly["unified_price"]
hourly["中长期-全网统一价差"] = hourly["lt_price"] - hourly["unified_price"]

price_cols = ["中长期合约价", "实时节点电价", "全网统一价"]
plot_df = hourly.set_index("时间")[price_cols].copy()
spread_df = hourly.set_index("时间")[["中长期-全网统一价差"]].copy()

if smooth_hours > 1:
    plot_df = plot_df.rolling(window=smooth_hours, min_periods=1, center=True).mean()
    spread_df = spread_df.rolling(window=smooth_hours, min_periods=1, center=True).mean()

st.subheader(f"{station_name}｜{month}-01 至 {end_date}｜价格平滑曲线")
st.line_chart(plot_df, use_container_width=True, height=420)

st.subheader("中长期合约价 - 全网统一价｜价差曲线")
st.line_chart(spread_df, use_container_width=True, height=300)

spread_raw = hourly["中长期-全网统一价差"].dropna()
price_raw = hourly[price_cols]
if not spread_raw.empty:
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("当前截止时点价差", f"{spread_raw.iloc[-1]:.2f} 元/MWh")
    m2.metric("累计平均价差", f"{spread_raw.mean():.2f} 元/MWh")
    m3.metric("最大正价差", f"{spread_raw.max():.2f} 元/MWh")
    m4.metric("最大负价差", f"{spread_raw.min():.2f} 元/MWh")

st.caption("图中“实时节点电价”读取日清分24时点 spot_price；“全网统一价”读取 unified_price；价差按每个时点 P_LT - P_全网统一 计算，再按所选小时窗口平滑。")

if show_raw:
    raw = hourly[["时间", "中长期合约价", "实时节点电价", "全网统一价", "中长期-全网统一价差"]].copy()
    for col in raw.columns[1:]:
        raw[col] = raw[col].round(4)
    st.subheader("原始时点数据")
    st.dataframe(raw, use_container_width=True, hide_index=True, height=520)
