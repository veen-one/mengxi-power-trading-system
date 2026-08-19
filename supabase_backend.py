import os
from datetime import date, datetime

import pandas as pd
import streamlit as st
from supabase import create_client

# URL 不是密钥，可以安全进入代码库；KEY 只从 Secrets / 环境变量读取。
DEFAULT_SUPABASE_URL = "https://bwhpatzqcklgyzuifvhu.supabase.co"


def _secret(name: str, default=None):
    try:
        return st.secrets.get(name, default)
    except Exception:
        return os.getenv(name, default)


@st.cache_resource
def get_supabase():
    url = _secret("SUPABASE_URL", DEFAULT_SUPABASE_URL)
    key = _secret("SUPABASE_KEY")
    if not key:
        raise RuntimeError(
            "未配置 SUPABASE_KEY。请在 .streamlit/secrets.toml 或部署平台 Secrets 中配置，切勿提交到 GitHub。"
        )
    return create_client(url, key)


def _records(df_or_records):
    if isinstance(df_or_records, pd.DataFrame):
        records = df_or_records.to_dict("records")
    elif isinstance(df_or_records, dict):
        records = [df_or_records]
    else:
        records = list(df_or_records)
    clean = []
    for row in records:
        item = {}
        for k, v in row.items():
            if pd.isna(v):
                item[k] = None
            elif isinstance(v, (date, datetime)):
                item[k] = v.isoformat()
            elif hasattr(v, "item"):
                item[k] = v.item()
            else:
                item[k] = v
        clean.append(item)
    return clean


def stations(active=True):
    q = get_supabase().table("stations").select("*").order("name")
    if active:
        q = q.eq("active", True)
    return pd.DataFrame(q.execute().data or [])


def create_station(payload):
    return get_supabase().table("stations").insert(_records(payload)[0]).execute()


def all_daily():
    return pd.DataFrame(get_supabase().table("daily_summary").select("*").order("trade_date").execute().data or [])


def month_daily(station_id: int, month: str):
    start = f"{month}-01"
    y, m = map(int, month.split("-"))
    end = f"{y + (m == 12):04d}-{1 if m == 12 else m + 1:02d}-01"
    data = (get_supabase().table("daily_summary").select("*")
            .eq("station_id", station_id).gte("trade_date", start).lt("trade_date", end)
            .order("trade_date").execute().data or [])
    return pd.DataFrame(data)


def save_daily(station_id: int, filename: str, file_hash: str, summary: dict, hourly: pd.DataFrame, overwrite=False):
    sb = get_supabase()
    trade_date = summary["trade_date"].isoformat() if hasattr(summary["trade_date"], "isoformat") else str(summary["trade_date"])
    existing = (sb.table("daily_summary").select("id").eq("station_id", station_id)
                .eq("trade_date", trade_date).limit(1).execute().data or [])
    if existing and not overwrite:
        raise ValueError("该场站该日期已存在；如需重算请勾选覆盖。")

    final = float(summary["energy_total"])
    actual = float(summary["actual_energy"])
    daily = {
        "station_id": station_id, "trade_date": trade_date, "source_file": filename, "file_hash": file_hash,
        "lt_energy": float(summary["lt_energy"]), "lt_price": float(summary["lt_price"]),
        "actual_energy": actual, "spot_price": float(summary["spot_price"]),
        "unified_price": float(summary["unified_price"]), "spot_fee": float(summary["spot_fee"]),
        "lt_diff_fee": float(summary["lt_diff_fee"]), "energy_total": final,
        "final_revenue": final, "final_price": final / actual if actual else 0.0,
    }
    sb.table("daily_summary").upsert(daily, on_conflict="station_id,trade_date").execute()

    # 24时点采用同一自然键 upsert，重复上传不会重复累计。
    detail = []
    for _, r in hourly.iterrows():
        detail.append({
            "station_id": station_id, "trade_date": trade_date, "hour_no": int(r["时点"]),
            "lt_energy": float(r["中长期电量"]), "lt_price": float(r["中长期价"]),
            "actual_energy": float(r["上网电量"]), "spot_price": float(r["实时价"]),
            "unified_price": float(r["统一结算点价"]), "spot_fee": float(r["现货电费"]),
            "lt_diff_fee": float(r["中长期差价电费"]), "energy_total": float(r["电能量合计"]),
        })
    sb.table("hourly_detail").upsert(detail, on_conflict="station_id,trade_date,hour_no").execute()


def update_settlement(row_id: int, payload: dict):
    return get_supabase().table("daily_summary").update(_records(payload)[0]).eq("id", row_id).execute()


def delete_daily(row_id: int, station_id: int, trade_date: str):
    sb = get_supabase()
    sb.table("hourly_detail").delete().eq("station_id", station_id).eq("trade_date", trade_date).execute()
    return sb.table("daily_summary").delete().eq("id", row_id).execute()


def healthcheck():
    return get_supabase().table("stations").select("id", count="exact").limit(1).execute()
