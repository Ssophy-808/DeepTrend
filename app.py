# DeepTrend app.py 功能地圖
#
# 1. 資料來源
#    - output/stock_analysis_result.xlsx：主表格與股票雷達的基礎分析結果。
#    - chip.csv：外資、投信與籌碼資料，由 update_chip.py 更新。
#    - groups.csv：股票族群分類，用於族群熱度判斷。
#    - 台灣證交所 / 櫃買中心官方日 K：個股 K 線、回測歷史資料與即時補價。
#    - yfinance：部分行情資料的備援來源。
#
# 2. 資料處理
#    - 讀取 Excel/CSV 後整理欄位型別、格式化數字、補上即時價格與漲跌幅。
#    - 計算技術指標、狀態標籤、量價異常、乖離率與族群強弱。
#    - 將股票代號標準化為 .TW / .TWO，方便不同資料源查詢。
#
# 3. UI 顯示
#    - 股票雷達：以卡片呈現重點股票、狀態、分數、量價訊號與籌碼提示。
#    - 強勢排行榜 / 策略排行榜：用排行方式快速比較股票表現與策略結果。
#    - 股票掃描：顯示完整資料表，搭配狀態、分數、關鍵字篩選。
#    - 個股查詢：顯示個股摘要、技術/籌碼說明與 K 線圖。
#
# 4. 回測功能
#    - 回測條件：5MA > 10MA、成交量大於 5日均量 1.5 倍、KD 黃金交叉。
#    - 支援 6個月 / 1年 / 3年期間，以及不同持有天數比較。
#    - 統計交易次數、勝率、平均報酬、最大回撤、盈虧比與信賴度。
#
# 5. 主程式流程
#    - 設定 Streamlit 頁面與標題。
#    - 讀取並整理股票分析資料。
#    - 顯示更新按鈕、篩選條件與主功能選單。
#    - 依選單切換股票雷達、排行榜、掃描、個股查詢與回測實驗室。
#
# Deep Trend 計算說明
#    - 股票雷達：使用 output/stock_analysis_result.xlsx 內的技術分數、狀態、籌碼欄位與即時補價後的衍生欄位呈現。
#    - 族群熱度：用 groups.csv 將觀察池股票映射到族群，再用技術分數、漲跌、乖離與籌碼資料估算族群強弱。
#    - 回測實驗室：只用官方日 K 重新計算均線、量能、突破條件與後續報酬，不回寫 Deep Trend 主分數。
#    - 觀察池溫度：只統計目前 watchlist / Excel 觀察池內股票，不代表全市場；分數由多頭排列、創高、量增、漲跌停與低價轉強比例組成。
#    - 新聞熱度：使用 Google News RSS 的近 7 天標題做關鍵字統計，只作輔助欄位，不進入主分數。

# =========================
# Imports and constants
# =========================

import math
import re
import subprocess
import sys
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date, datetime, timedelta
from email.utils import parsedate_to_datetime
from pathlib import Path
from textwrap import dedent
from urllib.parse import quote_plus
import xml.etree.ElementTree as ET

import pandas as pd
import plotly.graph_objects as go
import requests
import streamlit as st
import urllib3
import yfinance as yf
from plotly.subplots import make_subplots


BASE_DIR = Path(__file__).resolve().parent
RESULT_FILE = BASE_DIR / "output" / "stock_analysis_result.xlsx"
UNIVERSE_RESULT_FILE = BASE_DIR / "output" / "universe_analysis_result.xlsx"
STOCK_ANALYSIS_HISTORY_FILE = BASE_DIR / "output" / "stock_analysis_history.csv"
CHIP_DAILY_FILE = BASE_DIR / "output" / "chip_daily.csv"
GROUP_FILE = BASE_DIR / "groups.csv"
GROUP_HEAT_HISTORY_FILE = BASE_DIR / "output" / "group_heat_history.csv"
BACKTEST_RECORD_DIR = BASE_DIR / "backtest_records"

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


# =========================
# Basic data helpers
# =========================

def get_series(df, column):
    """Return a single Series from yfinance data, even when columns are MultiIndex/DataFrame-like."""
    data = df[column]
    if isinstance(data, pd.DataFrame):
        return data.iloc[:, 0]
    return data


def normalize_tw_symbol(symbol):
    """Normalize plain Taiwan stock codes into yfinance-style symbols such as 2330.TW."""
    symbol = str(symbol).strip()

    if not symbol:
        return symbol
    if symbol.startswith("^") or "." in symbol or symbol.endswith("=F"):
        return symbol
    if symbol.isdigit():
        return f"{symbol}.TW"

    return symbol


@st.cache_data(ttl=600)
def download_market_data(symbol, period="3mo", interval="1d"):
    """Download market data through yfinance with a short cache to reduce repeated requests."""
    return yf.download(
        symbol,
        period=period,
        interval=interval,
        progress=False,
        auto_adjust=False,
    )


def first_quote_price(value):
    """Extract the first price from TWSE realtime quote strings such as bid/ask values joined by underscores."""
    text = str(value or "").strip()
    if not text or text == "-":
        return 0
    first = text.split("_")[0].replace(",", "").strip()
    try:
        return float(first)
    except ValueError:
        return 0


def to_float(value):
    """Convert exchange-provided numeric strings into float while treating blank/null values as 0."""
    try:
        text = str(value).replace(",", "").replace("%", "").strip()
        if text in ["", "-", "NULL"]:
            return 0
        return float(text)
    except (TypeError, ValueError):
        return 0


def parse_roc_date(value):
    """Convert Taiwan ROC date strings like 113/01/02 into Python date objects."""
    parts = str(value).split("/")
    if len(parts) != 3:
        return None
    try:
        year = int(parts[0]) + 1911
        month = int(parts[1])
        day = int(parts[2])
        return date(year, month, day)
    except ValueError:
        return None


def recent_month_starts(month_count=3):
    """Build YYYYMM01 strings for the latest N months used by TWSE/TPEX historical APIs."""
    today = date.today()
    year = today.year
    month = today.month
    starts = []

    for _ in range(month_count):
        starts.append(f"{year}{month:02d}01")
        month -= 1
        if month == 0:
            month = 12
            year -= 1

    return starts


def build_twse_channel(ticker):
    """Build the TWSE realtime channel id for listed or OTC symbols."""
    text = str(ticker).strip()
    code = text.split(".")[0]

    if not code.isdigit():
        return None

    if text.endswith(".TWO"):
        return f"otc_{code}.tw"

    return f"tse_{code}.tw"


@st.cache_data(ttl=300)
def get_official_daily_history(ticker, month_count=3):
    """Fetch official TWSE/TPEX daily K data and normalize it into 日期/收盤價/最高價/最低價/成交量 columns."""
    text = str(ticker).strip()
    code = text.split(".")[0]

    if not code:
        return pd.DataFrame(columns=["日期", "收盤價", "最高價", "最低價", "成交量"])

    rows = []
    headers = {"User-Agent": "Mozilla/5.0"}

    for month_start in recent_month_starts(month_count):
        try:
            if text.endswith(".TWO"):
                response = requests.get(
                    "https://www.tpex.org.tw/www/zh-tw/afterTrading/tradingStock",
                    params={"date": f"{month_start[:4]}/{month_start[4:6]}/01", "code": code, "response": "json"},
                    timeout=10,
                    verify=False,
                    headers=headers,
                )
                tables = response.json().get("tables", [])
                data_rows = tables[0].get("data", []) if tables else []

                for row in data_rows:
                    trade_date = parse_roc_date(row[0])
                    if not trade_date:
                        continue
                    rows.append(
                        {
                            "日期": trade_date,
                            "成交量": to_float(row[1]) * 1000,
                            "最高價": to_float(row[4]),
                            "最低價": to_float(row[5]),
                            "收盤價": to_float(row[6]),
                        }
                    )
            else:
                response = requests.get(
                    "https://www.twse.com.tw/exchangeReport/STOCK_DAY",
                    params={"response": "json", "date": month_start, "stockNo": code},
                    timeout=10,
                    verify=False,
                    headers=headers,
                )
                payload = response.json()

                if payload.get("stat") != "OK":
                    continue

                for row in payload.get("data", []):
                    trade_date = parse_roc_date(row[0])
                    if not trade_date:
                        continue
                    rows.append(
                        {
                            "日期": trade_date,
                            "成交量": to_float(row[1]),
                            "最高價": to_float(row[4]),
                            "最低價": to_float(row[5]),
                            "收盤價": to_float(row[6]),
                        }
                    )
        except Exception:
            continue

    if not rows:
        return pd.DataFrame(columns=["日期", "收盤價", "最高價", "最低價", "成交量"])

    history = pd.DataFrame(rows)
    history = history.dropna(subset=["日期"])
    history = history.sort_values("日期").drop_duplicates(subset=["日期"], keep="last")
    return history.tail(month_count * 24).reset_index(drop=True)


def official_history_to_kline(history):
    """Convert official TWSE/TPEX daily rows into close-line chart data without fake open prices."""
    if history.empty:
        return pd.DataFrame()

    k_df = history.copy()
    k_df = k_df.rename(
        columns={
            "日期": "Date",
            "收盤價": "Close",
            "最高價": "High",
            "最低價": "Low",
            "成交量": "Volume",
        }
    )
    k_df["Date"] = pd.to_datetime(k_df["Date"])
    k_df = k_df.set_index("Date").sort_index()

    for col in ["Close", "High", "Low", "Volume"]:
        k_df[col] = pd.to_numeric(k_df[col], errors="coerce")

    k_df = k_df.dropna(subset=["High", "Low", "Close"])

    return k_df[["High", "Low", "Close", "Volume"]]


@st.cache_data(ttl=10)
def get_twse_realtime_map(tickers):
    """Fetch TWSE realtime quote snapshots for the current watchlist; failures are skipped per ticker."""
    channels = []
    ticker_by_channel = {}

    for ticker in tickers:
        channel = build_twse_channel(ticker)
        if channel:
            channels.append(channel)
            ticker_by_channel[channel] = ticker

    if not channels:
        return {}

    url = "https://mis.twse.com.tw/stock/api/getStockInfo.jsp"
    headers = {
        "User-Agent": "Mozilla/5.0",
        "Referer": "https://mis.twse.com.tw/stock/fibest.jsp?stock=0050",
    }

    realtime = {}

    for start in range(0, len(channels), 50):
        batch = channels[start : start + 50]

        try:
            response = requests.get(
                url,
                params={"ex_ch": "|".join(batch), "json": "1", "delay": "0"},
                timeout=10,
                verify=False,
                headers=headers,
            )
            items = response.json().get("msgArray", [])
        except Exception:
            continue

        for item in items:
            channel = item.get("ex", "") + "_" + item.get("ch", "")
            ticker = ticker_by_channel.get(channel)

            if not ticker:
                continue

            latest_price = to_float(item.get("z"))

            if latest_price == 0:
                bid = first_quote_price(item.get("b"))
                ask = first_quote_price(item.get("a"))

                if bid and ask:
                    latest_price = (bid + ask) / 2
                elif bid:
                    latest_price = bid
                elif ask:
                    latest_price = ask
                else:
                    latest_price = to_float(item.get("o")) or to_float(item.get("y"))

            if latest_price:
                previous_close = to_float(item.get("y"))
                realtime[ticker] = {
                    "price": latest_price,
                    "previous_close": previous_close,
                    "high": to_float(item.get("h")),
                    "low": to_float(item.get("l")),
                    "volume": to_float(item.get("v")) * 1000,
                    "time": item.get("t") or item.get("%") or "",
                }

    return realtime


def color_status(val):
    """Style the 狀態 column in the detailed stock table based on its emoji label."""
    value = str(val)

    if "🔥" in value:
        return "background-color: #14532d; color: white"
    if "👀" in value:
        return "background-color: #1e3a8a; color: white"
    if "⚠️" in value:
        return "background-color: #92400e; color: white"
    if "❌" in value:
        return "background-color: #7f1d1d; color: white"

    return ""


def format_number(value, decimals=2):
    """Format floats for UI display while keeping missing values blank."""
    if pd.isna(value):
        return ""
    try:
        return f"{float(value):,.{decimals}f}"
    except (TypeError, ValueError):
        return value


def format_integer(value):
    """Format integer values such as volume and chip counts for UI display."""
    if pd.isna(value):
        return ""
    try:
        return f"{int(value):,}"
    except (TypeError, ValueError):
        return value


def format_signed_pct(value):
    """Format returns/changes as signed percentages, e.g. +3.25% or -1.10%."""
    if pd.isna(value):
        return ""
    try:
        return f"{float(value):+,.2f}%"
    except (TypeError, ValueError):
        return value


def value_color(value):
    """Choose the app's red/green/neutral color for positive, negative, or zero values."""
    try:
        number = float(value)
    except (TypeError, ValueError):
        return "#aaaaaa"

    if number > 0:
        return "#ff4b4b"
    if number < 0:
        return "#22c55e"
    return "#aaaaaa"


def stock_code_key(value):
    """Normalize stock identifiers to the numeric code so groups/watchlist/output can be joined."""
    return str(value).strip().split(".")[0]


def build_stock_label_map(df):
    """Create selectbox labels in the form '代號 名稱'."""
    if df.empty:
        return {}

    label_df = df[["股票代號", "股票名稱"]].drop_duplicates(subset=["股票代號"], keep="first")
    return {
        str(row["股票代號"]): f"{row['股票代號']} {row['股票名稱']}"
        for _, row in label_df.iterrows()
    }


@st.cache_data(ttl=600)
def load_stock_analysis_history():
    """Load saved DeepTrend daily snapshots for score component comparison."""
    if not STOCK_ANALYSIS_HISTORY_FILE.exists():
        return pd.DataFrame()

    try:
        history_df = pd.read_csv(STOCK_ANALYSIS_HISTORY_FILE)
    except Exception:
        return pd.DataFrame()

    if history_df.empty or "snapshot_date" not in history_df.columns or "股票代號" not in history_df.columns:
        return pd.DataFrame()

    history_df["snapshot_date"] = pd.to_datetime(history_df["snapshot_date"], errors="coerce")
    history_df["股票代號_key"] = history_df["股票代號"].map(stock_code_key)
    return history_df.dropna(subset=["snapshot_date", "股票代號_key"])


def row_number(row, column):
    """Read a numeric value from a row-like object without raising on missing columns."""
    try:
        return pd.to_numeric(pd.Series([row.get(column)]), errors="coerce").iloc[0]
    except Exception:
        return pd.NA


def score_change_label(previous_value, today_value):
    """Format component score difference for the score composition table."""
    if pd.isna(previous_value) or pd.isna(today_value):
        return ""

    change = today_value - previous_value
    return format_signed_pct(change).replace("%", "")


def describe_technical_component(today_row, previous_row=None):
    """Explain technical score changes with concrete MA and breakout conditions."""
    close = row_number(today_row, "收盤價")
    ma5 = row_number(today_row, "5日線")
    ma10 = row_number(today_row, "10日線")
    ma20 = row_number(today_row, "20日線")
    high20 = row_number(today_row, "20日高點")
    low20 = row_number(today_row, "20日低點")

    notes = []
    if pd.notna(close) and pd.notna(ma5):
        notes.append("站上5MA" if close > ma5 else "跌破5MA")
    if pd.notna(ma5) and pd.notna(ma10) and pd.notna(ma20):
        if ma5 > ma10 > ma20:
            notes.append("5MA>10MA>20MA 多頭排列")
        elif ma5 < ma10 < ma20:
            notes.append("5MA<10MA<20MA 空頭排列")
        else:
            notes.append("均線糾結")
    if pd.notna(close) and pd.notna(high20) and close >= high20:
        notes.append("突破20日高點")
    elif pd.notna(close) and pd.notna(high20) and close >= high20 * 0.98:
        notes.append("接近20日高點")
    if pd.notna(close) and pd.notna(low20) and close <= low20 * 1.05:
        notes.append("接近/跌破20日低點")

    if previous_row is not None:
        prev_close = row_number(previous_row, "收盤價")
        prev_ma5 = row_number(previous_row, "5日線")
        if pd.notna(prev_close) and pd.notna(prev_ma5) and pd.notna(close) and pd.notna(ma5):
            if prev_close > prev_ma5 and close <= ma5:
                notes.insert(0, "由站上5MA轉為跌破5MA")
            elif prev_close <= prev_ma5 and close > ma5:
                notes.insert(0, "由跌破5MA轉為站上5MA")

    return "、".join(notes) if notes else "技術面資料不足"


def describe_chip_component(today_row, previous_row=None):
    """Explain chip score changes with foreign/investment trust and 5-day chip flow."""
    chip_5d = row_number(today_row, "籌碼5日")
    foreign_5d = row_number(today_row, "外資5日")
    investment_5d = row_number(today_row, "投信5日")

    notes = []
    if pd.notna(chip_5d):
        if chip_5d > 0:
            notes.append(f"法人5日買超 {format_integer(chip_5d)}")
        elif chip_5d < 0:
            notes.append(f"法人5日賣超 {format_integer(abs(chip_5d))}")
        else:
            notes.append("法人5日中性")
    if pd.notna(foreign_5d) and pd.notna(investment_5d):
        if foreign_5d > 0 and investment_5d > 0:
            notes.append("外資、投信同步買超")
        elif foreign_5d < 0 and investment_5d < 0:
            notes.append("外資、投信同步賣超")
        elif foreign_5d > 0 and investment_5d < 0:
            notes.append("外資買、投信賣")
        elif foreign_5d < 0 and investment_5d > 0:
            notes.append("投信買、外資賣")

    if previous_row is not None:
        prev_chip_5d = row_number(previous_row, "籌碼5日")
        if pd.notna(prev_chip_5d) and pd.notna(chip_5d):
            if prev_chip_5d <= 0 < chip_5d:
                notes.insert(0, "籌碼由賣壓轉買超")
            elif prev_chip_5d >= 0 > chip_5d:
                notes.insert(0, "籌碼由買超轉賣壓")

    return "、".join(notes) if notes else "籌碼資料不足"


def describe_volume_price_component(today_row, previous_row=None):
    """Explain volume-price score changes with volume ratio and breakout state."""
    close = row_number(today_row, "收盤價")
    ma_volume = row_number(today_row, "5日均量")
    volume = row_number(today_row, "成交量")
    high20 = row_number(today_row, "20日高點")
    signal_text = str(today_row.get("量價異常", ""))

    notes = []
    if pd.notna(volume) and pd.notna(ma_volume) and ma_volume:
        volume_ratio = volume / ma_volume
        if volume_ratio >= 2:
            notes.append(f"成交量約5日均量 {volume_ratio:.1f} 倍")
        elif volume_ratio >= 1.5:
            notes.append(f"量能放大至5日均量 {volume_ratio:.1f} 倍")
        else:
            notes.append("量能未明顯放大")
    if pd.notna(close) and pd.notna(high20) and close >= high20:
        notes.append("價格突破20日高點")
    if signal_text and signal_text != "無明顯異常":
        notes.append(signal_text)

    return "、".join(notes) if notes else "量價無明顯異常"


def build_score_component_change(selected_row):
    """Build a yesterday/today comparison table for DeepTrend score components."""
    score_columns = [
        ("技術面", "技術面分數", describe_technical_component),
        ("籌碼面", "籌碼分數", describe_chip_component),
        ("量價面", "量價分數", describe_volume_price_component),
    ]
    today_key = stock_code_key(selected_row["股票代號"])
    today_values = {
        label: row_number(selected_row, column)
        for label, column, _ in score_columns
    }

    history_df = load_stock_analysis_history()
    previous_row = None
    if not history_df.empty:
        stock_history = history_df[history_df["股票代號_key"] == today_key].sort_values("snapshot_date")
        stock_history = stock_history[stock_history["snapshot_date"] < pd.Timestamp(date.today())]
        if not stock_history.empty:
            previous_row = stock_history.iloc[-1]

    rows = []
    for label, column, describer in score_columns:
        previous_value = None if previous_row is None else row_number(previous_row, column)
        today_value = today_values[label]
        rows.append(
            {
                "項目": label,
                "昨天": "" if pd.isna(previous_value) else format_number(previous_value, 2),
                "今天": "" if pd.isna(today_value) else format_number(today_value, 2),
                "變化": score_change_label(previous_value, today_value),
                "說明": describer(selected_row, previous_row),
            }
        )

    return pd.DataFrame(rows)


@st.cache_data(ttl=600)
def load_group_data():
    """Load and normalize groups.csv; cached because group mappings rarely change during a session."""
    if not GROUP_FILE.exists():
        return pd.DataFrame(columns=["股票代號", "族群", "股票代號_key"])

    group_df = pd.read_csv(GROUP_FILE)
    group_df = group_df.rename(columns={"ticker": "股票代號", "group": "族群"})
    group_df["股票代號"] = group_df["股票代號"].astype(str)
    group_df["股票代號_key"] = group_df["股票代號"].map(stock_code_key)
    group_df["族群"] = group_df["族群"].astype(str).str.strip()
    return group_df[group_df["族群"].str.len() > 0]


def group_heat_status(heat_score, bullish_ratio, member_count):
    """Convert a group heat score and bullish ratio into a compact status label."""
    if member_count >= 2 and bullish_ratio >= 0.8 and heat_score >= 75:
        return "🔥 全面轉強"
    if bullish_ratio >= 0.55 and heat_score >= 55:
        return "🟢 轉強"
    if heat_score >= 35:
        return "👀 觀察"
    return "⚠️ 轉弱"


def load_group_heat_history():
    """Load previously saved daily group heat snapshots."""
    if not GROUP_HEAT_HISTORY_FILE.exists():
        return pd.DataFrame()

    try:
        history_df = pd.read_csv(GROUP_HEAT_HISTORY_FILE)
    except Exception:
        return pd.DataFrame()

    if history_df.empty or "日期" not in history_df.columns or "族群" not in history_df.columns:
        return pd.DataFrame()

    history_df["日期"] = pd.to_datetime(history_df["日期"], errors="coerce")
    history_df["熱度分數"] = pd.to_numeric(history_df.get("熱度分數"), errors="coerce")
    return history_df.dropna(subset=["日期", "族群", "熱度分數"])


def add_group_heat_trend(heat_df):
    """Add 7-day heat change and warming/cooling labels from saved history."""
    heat_df = heat_df.copy()
    heat_df["7日熱度變化"] = pd.NA
    heat_df["溫度趨勢"] = "資料不足"

    history_df = load_group_heat_history()
    if history_df.empty:
        return heat_df

    target_date = pd.Timestamp(date.today() - timedelta(days=7))
    baseline = history_df[history_df["日期"] <= target_date]
    if baseline.empty:
        return heat_df

    baseline = baseline.sort_values("日期").drop_duplicates(subset=["族群"], keep="last")
    baseline = baseline[["族群", "熱度分數"]].rename(columns={"熱度分數": "7日前熱度分數"})
    heat_df = heat_df.merge(baseline, on="族群", how="left")
    heat_df["7日熱度變化"] = heat_df["熱度分數"] - heat_df["7日前熱度分數"]
    heat_df.loc[heat_df["7日熱度變化"] > 0, "溫度趨勢"] = "升溫中"
    heat_df.loc[heat_df["7日熱度變化"] < 0, "溫度趨勢"] = "降溫中"
    heat_df.loc[heat_df["7日熱度變化"] == 0, "溫度趨勢"] = "持平"
    heat_df.loc[heat_df["7日熱度變化"].isna(), "溫度趨勢"] = "資料不足"
    return heat_df.drop(columns=["7日前熱度分數"])


def save_group_heat_history(heat_df):
    """Save one daily group heat snapshot, replacing same-day rows to avoid rerun duplicates."""
    if heat_df.empty:
        return

    history_columns = [
        "日期",
        "族群",
        "熱度分數",
        "檔數",
        "偏多檔數",
        "強勢檔數",
        "偏多比例",
        "今日漲跌幅",
        "平均乖離率",
        "法人5日",
        "外資5日",
        "投信5日",
    ]
    today_text = date.today().isoformat()
    snapshot = heat_df.copy()
    snapshot["日期"] = today_text
    snapshot = snapshot[[col for col in history_columns if col in snapshot.columns]]

    existing = load_group_heat_history()
    if not existing.empty:
        existing["日期"] = existing["日期"].dt.strftime("%Y-%m-%d")
        existing = existing[existing["日期"] != today_text]

    combined = pd.concat([existing, snapshot], ignore_index=True)
    GROUP_HEAT_HISTORY_FILE.parent.mkdir(parents=True, exist_ok=True)
    combined.to_csv(GROUP_HEAT_HISTORY_FILE, index=False, encoding="utf-8-sig")


def build_group_heat(df):
    """Aggregate individual stock signals into group-level heat rankings."""
    group_df = load_group_data()
    if group_df.empty or df.empty:
        return pd.DataFrame()

    stock_df = df.copy()
    stock_df["股票代號_key"] = stock_df["股票代號"].astype(str).map(stock_code_key)
    merged = group_df.merge(stock_df, on="股票代號_key", how="inner", suffixes=("_group", ""))

    if merged.empty:
        return pd.DataFrame()

    numeric_columns = ["技術分數", "今日漲跌幅", "乖離率", "外資5日", "投信5日", "籌碼5日"]
    for col in numeric_columns:
        if col in merged.columns:
            merged[col] = pd.to_numeric(merged[col], errors="coerce").fillna(0)

    rows = []
    for group_name, group_rows in merged.groupby("族群"):
        member_count = len(group_rows)
        if member_count < 3:
            continue

        bullish_mask = group_rows["技術分數"] >= 50
        strong_mask = group_rows["技術分數"] >= 70
        bullish_count = int(bullish_mask.sum())
        strong_count = int(strong_mask.sum())
        bullish_ratio = bullish_count / member_count if member_count else 0
        avg_score = float(group_rows["技術分數"].mean())
        avg_change = float(group_rows["今日漲跌幅"].mean())
        avg_bias = float(group_rows["乖離率"].mean())
        chip_5d = float(group_rows["籌碼5日"].sum()) if "籌碼5日" in group_rows.columns else 0
        foreign_5d = float(group_rows["外資5日"].sum()) if "外資5日" in group_rows.columns else 0
        investment_5d = float(group_rows["投信5日"].sum()) if "投信5日" in group_rows.columns else 0

        # 族群熱度公式：
        # 1. 以族群平均技術分數為底。
        # 2. 加入今日平均漲跌幅、偏多比例。
        # 3. 再用 5 日法人/籌碼方向做小幅加減分。
        heat_score = avg_score
        heat_score += max(min(avg_change * 5, 15), -15)
        heat_score += bullish_ratio * 20
        if chip_5d > 0:
            heat_score += 8
        elif chip_5d < 0:
            heat_score -= 8
        if foreign_5d > 0 and investment_5d > 0:
            heat_score += 8
        elif foreign_5d < 0 and investment_5d < 0:
            heat_score -= 8
        heat_score = max(0, min(100, heat_score))

        top_members = group_rows.sort_values("技術分數", ascending=False).head(4)
        leader_text = "、".join(
            f"{row['股票名稱']}({format_number(row['技術分數'], 0)})"
            for _, row in top_members.iterrows()
        )

        rows.append(
            {
                "族群": group_name,
                "熱度分數": round(heat_score, 1),
                "狀態": group_heat_status(heat_score, bullish_ratio, member_count),
                "檔數": member_count,
                "偏多檔數": bullish_count,
                "強勢檔數": strong_count,
                "偏多比例": bullish_ratio,
                "平均技術分數": round(avg_score, 1),
                "今日漲跌幅": round(avg_change, 2),
                "平均乖離率": round(avg_bias, 2),
                "法人5日": chip_5d,
                "外資5日": foreign_5d,
                "投信5日": investment_5d,
                "領先股": leader_text,
            }
        )

    if not rows:
        return pd.DataFrame()

    heat_df = pd.DataFrame(rows).sort_values(["熱度分數", "偏多比例"], ascending=False)
    heat_df = add_group_heat_trend(heat_df)
    save_group_heat_history(heat_df)
    return heat_df.sort_values(["熱度分數", "偏多比例"], ascending=False)


def add_signal_labels(df):
    """Add derived labels such as moving-average structure and volume breakout alerts."""
    labeled_df = df.copy()

    ma5 = pd.to_numeric(labeled_df.get("5日線"), errors="coerce")
    ma10 = pd.to_numeric(labeled_df.get("10日線"), errors="coerce")
    ma20 = pd.to_numeric(labeled_df.get("20日線"), errors="coerce")
    close = pd.to_numeric(labeled_df.get("收盤價"), errors="coerce")
    volume = pd.to_numeric(labeled_df.get("成交量"), errors="coerce")
    avg_volume_5 = pd.to_numeric(labeled_df.get("5日均量"), errors="coerce")
    high_20 = pd.to_numeric(labeled_df.get("20日高點"), errors="coerce")

    labeled_df["均線型態"] = "中性"
    labeled_df.loc[(ma5 > ma10) & (ma10 > ma20), "均線型態"] = "🔥 多頭排列"
    labeled_df.loc[(ma5 < ma10) & (ma10 < ma20), "均線型態"] = "🔴 空頭排列"

    # 量價異常公式：成交量相對 5 日均量放大，且價格站上/接近 20 日高點。
    strong_breakout = (volume > avg_volume_5 * 2) & (close >= high_20)
    volume_breakout = (volume > avg_volume_5 * 1.5) & (close >= high_20)

    labeled_df["突破警報"] = "無"
    labeled_df.loc[volume_breakout, "突破警報"] = "🟢 帶量突破"
    labeled_df.loc[strong_breakout, "突破警報"] = "🚨 強勢突破"

    return labeled_df


def load_stock_result():
    """Load the generated Excel analysis file that drives the main app screens."""
    try:
        return pd.read_excel(RESULT_FILE)
    except FileNotFoundError:
        st.error(f"找不到分析結果檔案：{RESULT_FILE}")
        st.stop()


@st.cache_data(ttl=600)
def load_universe_result():
    """Load the independent 150-stock market-pool analysis file when available."""
    try:
        return pd.read_excel(UNIVERSE_RESULT_FILE)
    except FileNotFoundError:
        return pd.DataFrame()
    except Exception:
        return pd.DataFrame()


def prepare_stock_data(df):
    """Coerce key columns to numeric types and add derived Deep Trend display labels."""
    numeric_columns = [
        "收盤價",
        "5日線",
        "10日線",
        "20日線",
        "20日高點",
        "20日低點",
        "技術分數",
        "成交量",
        "5日均量",
    ]

    for col in numeric_columns:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    if "今日漲跌幅" not in df.columns:
        df["今日漲跌幅"] = pd.NA

    # 乖離率公式：目前收盤價相對 5 日線的距離，用於強勢排行和雷達排序。
    df["乖離率"] = ((df["收盤價"] - df["5日線"]) / df["5日線"] * 100).round(2)
    return add_signal_labels(df)


def apply_realtime_prices(df):
    """Optionally patch the Excel data with TWSE realtime quotes for intraday display."""
    if "股票代號" not in df.columns or "收盤價" not in df.columns:
        return df

    updated_df = df.copy()
    updated_df["資料時間"] = ""
    tickers = updated_df["股票代號"].astype(str).dropna().unique().tolist()
    realtime = get_twse_realtime_map(tickers)
    histories = {}

    if realtime:
        max_workers = min(8, len(realtime))
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            future_by_ticker = {
                executor.submit(get_official_daily_history, ticker): ticker
                for ticker in realtime
            }

            for future in as_completed(future_by_ticker):
                ticker = future_by_ticker[future]
                try:
                    histories[ticker] = future.result()
                except Exception:
                    histories[ticker] = pd.DataFrame(columns=["日期", "收盤價", "最高價", "最低價", "成交量"])

    for ticker, data in realtime.items():
        mask = updated_df["股票代號"].astype(str) == ticker
        history = histories.get(ticker, pd.DataFrame(columns=["日期", "收盤價", "最高價", "最低價", "成交量"]))
        today = date.today()
        new_close = float(data["price"])
        previous_close = data.get("previous_close") or 0

        if not history.empty:
            today_row = {
                "日期": today,
                "收盤價": new_close,
                "最高價": data.get("high") or new_close,
                "最低價": data.get("low") or new_close,
                "成交量": data.get("volume") or 0,
            }

            history = history[history["日期"] != today]
            history = pd.concat([history, pd.DataFrame([today_row])], ignore_index=True)
            history = history.sort_values("日期").tail(60)

            close_series = pd.to_numeric(history["收盤價"], errors="coerce").dropna()
            high_series = pd.to_numeric(history["最高價"], errors="coerce").dropna()
            low_series = pd.to_numeric(history["最低價"], errors="coerce").dropna()

            updated_df.loc[mask, "收盤價"] = close_series.iloc[-1]

            if "5日線" in updated_df.columns and len(close_series) >= 5:
                updated_df.loc[mask, "5日線"] = close_series.tail(5).mean()

            if "10日線" in updated_df.columns and len(close_series) >= 10:
                updated_df.loc[mask, "10日線"] = close_series.tail(10).mean()

            if "20日線" in updated_df.columns and len(close_series) >= 20:
                updated_df.loc[mask, "20日線"] = close_series.tail(20).mean()

            if "20日高點" in updated_df.columns and len(high_series) >= 1:
                updated_df.loc[mask, "20日高點"] = high_series.tail(20).max()

            if "20日低點" in updated_df.columns and len(low_series) >= 1:
                updated_df.loc[mask, "20日低點"] = low_series.tail(20).min()
        else:
            updated_df.loc[mask, "收盤價"] = new_close

        if "成交量" in updated_df.columns and data.get("volume"):
            updated_df.loc[mask, "成交量"] = data["volume"]

        if previous_close:
            updated_df.loc[mask, "今日漲跌幅"] = round((new_close - previous_close) / previous_close * 100, 2)

        updated_df.loc[mask, "資料時間"] = data.get("time", "")

    if "5日線" in updated_df.columns:
        updated_df["乖離率"] = (
            (updated_df["收盤價"] - updated_df["5日線"]) / updated_df["5日線"] * 100
        ).round(2)

    return add_signal_labels(updated_df)


def render_rank(top_strength):
    """Render the simple strength leaderboard sorted by 乖離率."""
    st.subheader("🚀 強勢股排行榜")

    if top_strength.empty:
        st.info("目前沒有可顯示的排行資料。")
        return

    for i, (_, row) in enumerate(top_strength.iterrows(), 1):
        color = "#ff4b4b" if row["乖離率"] > 0 else "#00c853" if row["乖離率"] < 0 else "#aaaaaa"
        html = dedent(
            f"""
            <div style="padding:12px;margin-bottom:10px;border-radius:12px;background-color:#111111;border:1px solid #333;">
                <span style="font-size:20px;font-weight:bold;color:white;">{i}. {row["股票名稱"]}</span>
                <span style="float:right;font-size:22px;font-weight:bold;color:{color};">{row["乖離率"]:+.2f}%</span>
            </div>
            """
        ).replace("\n", "")
        st.markdown(html, unsafe_allow_html=True)


def render_group_heat(df):
    """Render group-level heat cards and the detailed group heat table."""
    st.subheader("🔥 族群熱度")

    heat_df = build_group_heat(df)
    if heat_df.empty:
        st.info("目前沒有族群資料可顯示。")
        return

    top_cols = st.columns(3)
    for index, (_, row) in enumerate(heat_df.head(3).iterrows()):
        change_color = value_color(row["今日漲跌幅"])
        chip_color = value_color(row["法人5日"])
        trend_color = value_color(row["7日熱度變化"])
        trend_text = row["溫度趨勢"]
        if pd.notna(row["7日熱度變化"]):
            trend_text = f'{trend_text}（{row["7日熱度變化"]:+.1f}）'
        with top_cols[index]:
            html = dedent(
                f"""
                <div style="
                    min-height:190px;
                    padding:18px;
                    border:1px solid #2f3542;
                    border-radius:8px;
                    background:#111827;
                ">
                    <div style="font-size:24px;font-weight:800;color:#ffffff;">{row["族群"]}</div>
                    <div style="margin-top:8px;font-size:15px;color:#d1d5db;">{row["狀態"]}</div>
                    <div style="margin-top:16px;font-size:34px;font-weight:900;color:#ffffff;">{format_number(row["熱度分數"], 1)}</div>
                    <div style="margin-top:4px;font-size:13px;color:{trend_color};">{trend_text}</div>
                    <div style="display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-top:12px;">
                        <div style="font-size:13px;color:#9ca3af;">偏多 <b style="color:#ffffff;">{row["偏多檔數"]}/{row["檔數"]}</b></div>
                        <div style="font-size:13px;color:#9ca3af;">今日 <b style="color:{change_color};">{format_signed_pct(row["今日漲跌幅"])}</b></div>
                        <div style="font-size:13px;color:#9ca3af;">法人5日 <b style="color:{chip_color};">{format_integer(row["法人5日"])}</b></div>
                        <div style="font-size:13px;color:#9ca3af;">強勢 <b style="color:#ffffff;">{row["強勢檔數"]}</b></div>
                    </div>
                    <div style="margin-top:14px;color:#9ca3af;font-size:13px;line-height:1.45;">{row["領先股"]}</div>
                </div>
                """
            ).replace("\n", "")
            st.markdown(html, unsafe_allow_html=True)

    table_df = heat_df.copy()
    table_df["偏多比例"] = table_df["偏多比例"].map(lambda value: f"{value:.0%}")
    for col in ["熱度分數", "平均技術分數", "今日漲跌幅", "平均乖離率", "7日熱度變化"]:
        table_df[col] = table_df[col].map(lambda value: format_number(value, 2))
    for col in ["法人5日", "外資5日", "投信5日"]:
        table_df[col] = table_df[col].map(format_integer)

    st.dataframe(table_df, use_container_width=True, hide_index=True)


def open_stock_detail(stock_code):
    """Switch the app state to the detail page for the clicked radar card."""
    stock_code = str(stock_code)
    st.session_state["pending_detail_stock"] = stock_code
    st.session_state["detail_stock"] = stock_code
    st.session_state["active_view"] = "🔎 個股查詢"


def render_stock_radar(filtered_df):
    """Render the primary stock radar cards. This is the main watchlist dashboard."""
    st.subheader("📊 股票雷達")
    st.caption(f"目前顯示 {len(filtered_df)} 檔股票")

    if filtered_df.empty:
        st.info("目前沒有符合篩選條件的股票。")
        return

    sort_options = {
        "技術分數高到低": (["技術分數", "乖離率"], [False, False]),
        "今日漲跌幅高到低": (["今日漲跌幅", "技術分數"], [False, False]),
        "乖離率高到低": (["乖離率", "技術分數"], [False, False]),
        "收盤價高到低": (["收盤價", "技術分數"], [False, False]),
    }
    selected_sort = st.selectbox("排序方式", list(sort_options.keys()), key="radar_sort")
    sort_columns, sort_ascending = sort_options[selected_sort]
    card_df = filtered_df.sort_values(sort_columns, ascending=sort_ascending, na_position="last")
    columns = st.columns(3)

    for index, (_, row) in enumerate(card_df.iterrows()):
        change = row.get("今日漲跌幅", pd.NA)
        bias = row.get("乖離率", pd.NA)
        change_color = value_color(change)
        bias_color = value_color(bias)
        score = format_number(row.get("技術分數"), 0)
        status = row.get("狀態", "")
        judgement = row.get("綜合判斷", "")
        time_text = row.get("資料時間", "")
        volume_price_signal = row.get("量價異常", "無明顯異常")
        ma_pattern = row.get("均線型態", "中性")
        breakout_alert = row.get("突破警報", "無")
        signal_text = breakout_alert if str(breakout_alert) != "無" else volume_price_signal
        signal_color = "#facc15" if str(signal_text) != "無明顯異常" and str(signal_text) != "無" else "#9ca3af"
        foreign_5d = row.get("外資5日", pd.NA)
        investment_5d = row.get("投信5日", pd.NA)
        foreign_color = value_color(foreign_5d)
        investment_color = value_color(investment_5d)

        html = dedent(
            f"""
            <div style="
                min-height:210px;
                padding:18px;
                margin-bottom:14px;
                border:1px solid #2f3542;
                border-radius:8px;
                background:#111827;
            ">
                <div style="display:flex;justify-content:space-between;gap:12px;align-items:flex-start;">
                    <div>
                        <div style="font-size:22px;font-weight:800;color:#ffffff;line-height:1.2;">{row["股票名稱"]}</div>
                        <div style="font-size:13px;color:#9ca3af;margin-top:4px;">{row["股票代號"]} · {time_text}</div>
                    </div>
                    <div style="font-size:18px;font-weight:800;color:#ffffff;white-space:nowrap;">{score}</div>
                </div>
                <div style="margin-top:14px;font-size:14px;color:#d1d5db;">{status}　{judgement}</div>
                <div style="display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-top:10px;">
                    <div style="font-size:13px;font-weight:700;color:#d1d5db;">{ma_pattern}</div>
                    <div style="font-size:13px;font-weight:700;color:{signal_color};">{signal_text}</div>
                </div>
                <div style="display:grid;grid-template-columns:1fr 1fr;gap:10px;margin-top:12px;">
                    <div style="font-size:13px;color:#d1d5db;">外資5日 <span style="font-weight:800;color:{foreign_color};">{format_integer(foreign_5d)}</span></div>
                    <div style="font-size:13px;color:#d1d5db;">投信5日 <span style="font-weight:800;color:{investment_color};">{format_integer(investment_5d)}</span></div>
                </div>
                <div style="display:grid;grid-template-columns:1fr 1fr 1fr;gap:10px;margin-top:18px;">
                    <div>
                        <div style="font-size:12px;color:#9ca3af;">收盤價</div>
                        <div style="font-size:19px;font-weight:800;color:#ffffff;">{format_number(row.get("收盤價"))}</div>
                    </div>
                    <div>
                        <div style="font-size:12px;color:#9ca3af;">今日漲跌幅</div>
                        <div style="font-size:19px;font-weight:800;color:{change_color};">{format_signed_pct(change)}</div>
                    </div>
                    <div>
                        <div style="font-size:12px;color:#9ca3af;">乖離率</div>
                        <div style="font-size:19px;font-weight:800;color:{bias_color};">{format_signed_pct(bias)}</div>
                    </div>
                </div>
                <div style="
                    margin-top:16px;
                    padding-top:12px;
                    border-top:1px solid #253041;
                    color:#9ca3af;
                    font-size:13px;
                    line-height:1.45;
                ">{row.get("技術面", "")}</div>
            </div>
            """
        ).replace("\n", "")

        with columns[index % 3]:
            st.markdown(html, unsafe_allow_html=True)
            stock_code = str(row["股票代號"])
            stock_name = str(row["股票名稱"])
            st.button(
                f"查看 {stock_name}",
                key=f"open_detail_{stock_code}_{index}",
                on_click=open_stock_detail,
                args=(stock_code,),
                use_container_width=True,
            )

@st.cache_data(ttl=600)
def load_chip_daily_data():
    """Load daily institutional chip detail used for audit and official-report reconciliation."""
    if not CHIP_DAILY_FILE.exists():
        return pd.DataFrame()

    try:
        chip_df = pd.read_csv(CHIP_DAILY_FILE)
    except Exception:
        return pd.DataFrame()

    if "date" not in chip_df.columns or "ticker" not in chip_df.columns:
        return pd.DataFrame()

    chip_df["date"] = pd.to_datetime(chip_df["date"], errors="coerce")
    chip_df = chip_df[chip_df["date"].notna()].copy()
    chip_df["ticker"] = chip_df["ticker"].astype(str).str.strip()
    chip_df["stock_code"] = chip_df["ticker"].str.split(".").str[0]

    numeric_columns = [
        "foreign_buy",
        "foreign_sell",
        "foreign_net",
        "foreign_dealer_buy",
        "foreign_dealer_sell",
        "foreign_dealer_net",
        "investment_buy",
        "investment_sell",
        "investment_net",
        "dealer_buy",
        "dealer_sell",
        "dealer_net",
        "dealer_self_buy",
        "dealer_self_sell",
        "dealer_self_net",
        "dealer_hedge_buy",
        "dealer_hedge_sell",
        "dealer_hedge_net",
        "total_net",
    ]
    for col in numeric_columns:
        if col in chip_df.columns:
            chip_df[col] = pd.to_numeric(chip_df[col], errors="coerce").fillna(0)

    return chip_df.sort_values(["date", "ticker", "source" if "source" in chip_df.columns else "ticker"])


def chip_sum(df, column):
    """Safely sum one chip column when the source file may be missing optional fields."""
    if column not in df.columns:
        return 0
    return int(pd.to_numeric(df[column], errors="coerce").fillna(0).sum())


def render_chip_audit(stock_df):
    """Render an audit page for daily institutional chip data and interval reconciliation."""
    st.subheader("🧾 籌碼查帳")
    st.caption("資料來源：output/chip_daily.csv。可用來對照證交所/櫃買中心的週報、月報區間加總。")

    chip_df = load_chip_daily_data()
    if chip_df.empty:
        st.warning("目前找不到每日籌碼明細。請先執行 update_chip.py 產生 output/chip_daily.csv。")
        return

    stock_options = []
    seen_codes = set()
    if not stock_df.empty and {"股票代號", "股票名稱"}.issubset(stock_df.columns):
        for _, row in stock_df[["股票代號", "股票名稱"]].dropna().iterrows():
            ticker = normalize_tw_symbol(row["股票代號"])
            code = str(ticker).split(".")[0]
            if code and code not in seen_codes:
                stock_options.append((code, f"{ticker} {row['股票名稱']}"))
                seen_codes.add(code)

    for code, name in (
        chip_df[["stock_code", "stock_name"]]
        .dropna()
        .drop_duplicates()
        .sort_values("stock_code")
        .itertuples(index=False, name=None)
    ):
        if code not in seen_codes:
            stock_options.append((str(code), f"{code} {name}"))
            seen_codes.add(code)

    if not stock_options:
        st.info("籌碼明細內目前沒有可選股票。")
        return

    min_date = chip_df["date"].min().date()
    max_date = chip_df["date"].max().date()

    col1, col2, col3 = st.columns([2, 1, 1])
    with col1:
        selected_label = st.selectbox(
            "選擇股票",
            [label for _, label in stock_options],
            index=0,
            key="chip_audit_stock",
        )
    with col2:
        start_date = st.date_input("開始日期", value=min_date, min_value=min_date, max_value=max_date)
    with col3:
        end_date = st.date_input("結束日期", value=max_date, min_value=min_date, max_value=max_date)

    if start_date > end_date:
        st.warning("開始日期不能晚於結束日期。")
        return

    selected_code = stock_options[[label for _, label in stock_options].index(selected_label)][0]
    selected_df = chip_df[
        chip_df["stock_code"].eq(selected_code)
        & (chip_df["date"].dt.date >= start_date)
        & (chip_df["date"].dt.date <= end_date)
    ].copy()

    if selected_df.empty:
        st.info("這個日期區間沒有該股票的籌碼明細。")
        return

    stock_name = selected_df["stock_name"].dropna().astype(str).iloc[-1] if "stock_name" in selected_df.columns else ""
    st.markdown(f"### {selected_code} {stock_name}")
    st.caption(f"區間：{start_date} 至 {end_date}，共 {selected_df['date'].nunique()} 個交易日")

    total_foreign = chip_sum(selected_df, "foreign_net") + chip_sum(selected_df, "foreign_dealer_net")
    total_investment = chip_sum(selected_df, "investment_net")
    total_dealer = chip_sum(selected_df, "dealer_net")
    total_all = chip_sum(selected_df, "total_net")

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("外資買賣超", format_integer(total_foreign))
    m2.metric("投信買賣超", format_integer(total_investment))
    m3.metric("自營商買賣超", format_integer(total_dealer))
    m4.metric("三大法人買賣超", format_integer(total_all))

    summary_rows = [
        {
            "項目": "外資(不含外資自營商)",
            "買進": chip_sum(selected_df, "foreign_buy"),
            "賣出": chip_sum(selected_df, "foreign_sell"),
            "買賣超": chip_sum(selected_df, "foreign_net"),
        },
        {
            "項目": "外資自營商",
            "買進": chip_sum(selected_df, "foreign_dealer_buy"),
            "賣出": chip_sum(selected_df, "foreign_dealer_sell"),
            "買賣超": chip_sum(selected_df, "foreign_dealer_net"),
        },
        {
            "項目": "投信",
            "買進": chip_sum(selected_df, "investment_buy"),
            "賣出": chip_sum(selected_df, "investment_sell"),
            "買賣超": chip_sum(selected_df, "investment_net"),
        },
        {
            "項目": "自營商自行買賣",
            "買進": chip_sum(selected_df, "dealer_self_buy"),
            "賣出": chip_sum(selected_df, "dealer_self_sell"),
            "買賣超": chip_sum(selected_df, "dealer_self_net"),
        },
        {
            "項目": "自營商避險",
            "買進": chip_sum(selected_df, "dealer_hedge_buy"),
            "賣出": chip_sum(selected_df, "dealer_hedge_sell"),
            "買賣超": chip_sum(selected_df, "dealer_hedge_net"),
        },
        {
            "項目": "自營商合計",
            "買進": chip_sum(selected_df, "dealer_buy"),
            "賣出": chip_sum(selected_df, "dealer_sell"),
            "買賣超": chip_sum(selected_df, "dealer_net"),
        },
        {
            "項目": "三大法人合計",
            "買進": "",
            "賣出": "",
            "買賣超": chip_sum(selected_df, "total_net"),
        },
    ]
    summary_df = pd.DataFrame(summary_rows)
    for col in ["買進", "賣出", "買賣超"]:
        summary_df[col] = summary_df[col].map(lambda value: "" if value == "" else format_integer(value))

    st.markdown("#### 區間加總")
    st.dataframe(summary_df, use_container_width=True, hide_index=True)

    display_columns = [
        "date",
        "ticker",
        "stock_name",
        "foreign_buy",
        "foreign_sell",
        "foreign_net",
        "foreign_dealer_net",
        "investment_net",
        "dealer_self_net",
        "dealer_hedge_net",
        "dealer_net",
        "total_net",
        "source",
    ]
    display_columns = [col for col in display_columns if col in selected_df.columns]
    detail_df = selected_df[display_columns].sort_values("date").copy()
    detail_df["date"] = detail_df["date"].dt.strftime("%Y-%m-%d")
    detail_df = detail_df.rename(
        columns={
            "date": "日期",
            "ticker": "股票代號",
            "stock_name": "股票名稱",
            "foreign_buy": "外資買進",
            "foreign_sell": "外資賣出",
            "foreign_net": "外資買賣超",
            "foreign_dealer_net": "外資自營商買賣超",
            "investment_net": "投信買賣超",
            "dealer_self_net": "自營商自行買賣",
            "dealer_hedge_net": "自營商避險",
            "dealer_net": "自營商合計",
            "total_net": "三大法人合計",
            "source": "來源",
        }
    )

    numeric_display_columns = [
        "外資買進",
        "外資賣出",
        "外資買賣超",
        "外資自營商買賣超",
        "投信買賣超",
        "自營商自行買賣",
        "自營商避險",
        "自營商合計",
        "三大法人合計",
    ]
    for col in numeric_display_columns:
        if col in detail_df.columns:
            detail_df[col] = detail_df[col].map(format_integer)

    st.markdown("#### 每日明細")
    st.dataframe(detail_df, use_container_width=True, hide_index=True)

    csv_df = selected_df.sort_values("date").copy()
    csv_df["date"] = csv_df["date"].dt.strftime("%Y-%m-%d")
    st.download_button(
        "下載此區間籌碼明細 CSV",
        data=csv_df.to_csv(index=False, encoding="utf-8-sig"),
        file_name=f"chip_audit_{selected_code}_{start_date}_{end_date}.csv",
        mime="text/csv",
    )


@st.cache_data(ttl=600)
def load_score_history_data():
    """Load saved DeepTrend score snapshots for per-stock trend review."""
    if not STOCK_ANALYSIS_HISTORY_FILE.exists():
        return pd.DataFrame()

    try:
        history_df = pd.read_csv(STOCK_ANALYSIS_HISTORY_FILE)
    except Exception:
        return pd.DataFrame()

    required_columns = {"snapshot_date", "股票代號", "股票名稱"}
    if not required_columns.issubset(history_df.columns):
        return pd.DataFrame()

    history_df = history_df.copy()
    history_df["snapshot_date"] = pd.to_datetime(history_df["snapshot_date"], errors="coerce")
    history_df = history_df[history_df["snapshot_date"].notna()].copy()
    history_df["股票代號"] = history_df["股票代號"].astype(str).str.strip()
    history_df["stock_code"] = history_df["股票代號"].str.split(".").str[0]

    score_columns = ["收盤價", "DeepTrend分數", "技術面分數", "籌碼分數", "量價分數", "技術分數", "分數變化", "分數變化率"]
    for col in score_columns:
        if col in history_df.columns:
            history_df[col] = pd.to_numeric(history_df[col], errors="coerce")

    return history_df.sort_values(["snapshot_date", "股票代號"])


def deeptrend_status_level(score):
    """Translate DeepTrend score into an easy-to-read strength stage."""
    if pd.isna(score):
        return "資料不足"
    score = float(score)
    if score >= 80:
        return "強勢"
    if score >= 60:
        return "轉強"
    if score >= 40:
        return "整理"
    if score >= 20:
        return "轉弱"
    return "避開"


def score_signal_label(row, previous_row=None):
    """Infer a lightweight buy/sell marker from score and basic MA state."""
    score = row_number(row, "DeepTrend分數")
    close = row_number(row, "收盤價")
    ma5 = row_number(row, "5日線")
    prev_score = row_number(previous_row, "DeepTrend分數") if previous_row is not None else pd.NA

    if pd.notna(score) and score >= 60 and (pd.isna(prev_score) or prev_score < 60):
        return "▲ 轉強訊號"
    if pd.notna(score) and score < 40 and pd.notna(prev_score) and prev_score >= 40:
        return "▼ 風險訊號"
    if pd.notna(close) and pd.notna(ma5) and close < ma5 and pd.notna(prev_score) and prev_score >= 60:
        return "▼ 風險訊號"
    return ""


def score_point_rows(row):
    """Build readable score component details from one saved snapshot."""
    close = row_number(row, "收盤價")
    ma5 = row_number(row, "5日線")
    ma10 = row_number(row, "10日線")
    ma20 = row_number(row, "20日線")
    high20 = row_number(row, "20日高點")
    volume = row_number(row, "成交量")
    avg_volume = row_number(row, "5日均量")
    chip_5d = row_number(row, "籌碼5日")
    foreign_5d = row_number(row, "外資5日")
    investment_5d = row_number(row, "投信5日")

    rows = []

    def add(category, condition, points, active):
        rows.append(
            {
                "項目": category,
                "條件": f"{'✓' if active else '☐'} {condition}",
                "分數": points if active else 0,
            }
        )

    add("技術面", "站上5MA", 10, pd.notna(close) and pd.notna(ma5) and close > ma5)
    add("技術面", "站上10MA", 10, pd.notna(close) and pd.notna(ma10) and close > ma10)
    add("技術面", "站上20MA", 10, pd.notna(close) and pd.notna(ma20) and close > ma20)
    add("技術面", "5MA > 10MA > 20MA", 20, pd.notna(ma5) and pd.notna(ma10) and pd.notna(ma20) and ma5 > ma10 > ma20)
    add("技術面", "接近20日高點", 15, pd.notna(close) and pd.notna(high20) and close >= high20 * 0.98)
    add("技術面", "突破20日高點", 20, pd.notna(close) and pd.notna(high20) and close >= high20)
    add("量價面", "量能放大", 10, pd.notna(volume) and pd.notna(avg_volume) and avg_volume > 0 and volume >= avg_volume * 1.5)
    add("籌碼面", "5日法人買超", 15, pd.notna(chip_5d) and chip_5d > 0)
    add("籌碼面", "外資5日買超", 10, pd.notna(foreign_5d) and foreign_5d > 0)
    add("籌碼面", "投信5日買超", 10, pd.notna(investment_5d) and investment_5d > 0)
    add(
        "籌碼面",
        "外資投信同步買超",
        15,
        pd.notna(foreign_5d) and pd.notna(investment_5d) and foreign_5d > 0 and investment_5d > 0,
    )

    return pd.DataFrame(rows)


def render_score_history(stock_df):
    """Render DeepTrend and component score history for one stock."""
    st.subheader("📈 分數歷史")
    st.caption("資料來源：output/stock_analysis_history.csv。用來觀察 DeepTrend 分數是否持續轉強或轉弱。")
    st.caption("階段小標：避開（<20） → 轉弱（20-39） → 整理（40-59） → 轉強（60-79） → 強勢（80+）")
    st.info(
        "買賣訊號標準："
        "▲ 轉強訊號 = DeepTrend 首次站上 60 分，或由 60 分以下重新轉強；"
        "▼ 風險訊號 = DeepTrend 跌破 40 分，或原本已轉強後收盤跌破 5MA。"
        "訊號是風險提示，仍需搭配 K 線、籌碼與市場環境判斷。"
    )

    history_df = load_score_history_data()
    if history_df.empty:
        st.warning("目前沒有可讀取的分數歷史資料。請先確認 output/stock_analysis_history.csv 是否存在且欄位完整。")
        return

    stock_options = []
    seen_codes = set()
    if not stock_df.empty and {"股票代號", "股票名稱"}.issubset(stock_df.columns):
        for _, row in stock_df[["股票代號", "股票名稱"]].dropna().iterrows():
            ticker = normalize_tw_symbol(row["股票代號"])
            code = str(ticker).split(".")[0]
            if code and code not in seen_codes:
                stock_options.append((code, f"{ticker} {row['股票名稱']}"))
                seen_codes.add(code)

    for code, name in (
        history_df[["stock_code", "股票名稱"]]
        .dropna()
        .drop_duplicates()
        .sort_values("stock_code")
        .itertuples(index=False, name=None)
    ):
        if code not in seen_codes:
            stock_options.append((str(code), f"{code} {name}"))
            seen_codes.add(code)

    if not stock_options:
        st.info("分數歷史內目前沒有可選股票。")
        return

    selected_label = st.selectbox(
        "選擇股票",
        [label for _, label in stock_options],
        index=0,
        key="score_history_stock",
    )
    selected_code = stock_options[[label for _, label in stock_options].index(selected_label)][0]

    selected_df = history_df[history_df["stock_code"].eq(selected_code)].copy()
    if selected_df.empty:
        st.info("這檔股票目前沒有分數歷史資料。")
        return

    stock_name = selected_df["股票名稱"].dropna().astype(str).iloc[-1] if "股票名稱" in selected_df.columns else ""
    selected_df = selected_df.drop_duplicates(subset=["snapshot_date", "股票代號"], keep="last")
    selected_df = selected_df.sort_values("snapshot_date")

    score_columns = [col for col in ["DeepTrend分數", "技術面分數", "籌碼分數", "量價分數"] if col in selected_df.columns]
    available_score_columns = [col for col in score_columns if selected_df[col].notna().any()]
    previous_rows = selected_df.shift(1)
    selected_df["狀態程度"] = selected_df["DeepTrend分數"].map(deeptrend_status_level) if "DeepTrend分數" in selected_df.columns else "資料不足"
    selected_df["買賣訊號"] = [
        score_signal_label(row, previous_rows.iloc[index] if index > 0 else None)
        for index, (_, row) in enumerate(selected_df.iterrows())
    ]

    st.markdown(f"### {selected_code} {stock_name}")
    st.caption(
        f"歷史區間：{selected_df['snapshot_date'].min().date()} 至 {selected_df['snapshot_date'].max().date()}，"
        f"共 {selected_df['snapshot_date'].nunique()} 筆快照"
    )

    latest_row = selected_df.iloc[-1]
    latest_deeptrend = latest_row.get("DeepTrend分數", pd.NA)
    latest_change = latest_row.get("分數變化", pd.NA)
    latest_change_rate = latest_row.get("分數變化率", pd.NA)

    metric_cols = st.columns(5)
    metric_cols[0].metric("最新 DeepTrend", format_number(latest_deeptrend, 1))
    metric_cols[1].metric("技術面", format_number(latest_row.get("技術面分數", pd.NA), 1))
    metric_cols[2].metric("籌碼面", format_number(latest_row.get("籌碼分數", pd.NA), 1))
    metric_cols[3].metric(
        "分數變化",
        format_number(latest_change, 1),
        delta=f"{format_number(latest_change_rate, 2)}%" if pd.notna(latest_change_rate) else None,
    )
    metric_cols[4].metric("狀態程度", latest_row.get("狀態程度", "資料不足"))

    if not available_score_columns:
        st.info("這檔股票目前還沒有 DeepTrend 分數組成歷史，之後每日更新會逐步累積。")
    else:
        fig = make_subplots(specs=[[{"secondary_y": True}]])
        color_map = {
            "DeepTrend分數": "#ef4444",
            "技術面分數": "#38bdf8",
            "籌碼分數": "#22c55e",
            "量價分數": "#facc15",
        }
        for col in available_score_columns:
            fig.add_trace(
                go.Scatter(
                    x=selected_df["snapshot_date"],
                    y=selected_df[col],
                    mode="lines+markers",
                    name=col,
                    line=dict(color=color_map.get(col)),
                ),
                secondary_y=False,
            )
        if "收盤價" in selected_df.columns and selected_df["收盤價"].notna().any():
            fig.add_trace(
                go.Scatter(
                    x=selected_df["snapshot_date"],
                    y=selected_df["收盤價"],
                    mode="lines",
                    name="股價",
                    line=dict(color="#a3a3a3", dash="dash"),
                ),
                secondary_y=True,
            )

        if "DeepTrend分數" in selected_df.columns:
            buy_df = selected_df[selected_df["買賣訊號"].eq("▲ 轉強訊號")]
            sell_df = selected_df[selected_df["買賣訊號"].eq("▼ 風險訊號")]
            if not buy_df.empty:
                fig.add_trace(
                    go.Scatter(
                        x=buy_df["snapshot_date"],
                        y=buy_df["DeepTrend分數"],
                        mode="markers",
                        name="轉強訊號",
                        marker=dict(color="#22c55e", symbol="triangle-up", size=13),
                    ),
                    secondary_y=False,
                )
            if not sell_df.empty:
                fig.add_trace(
                    go.Scatter(
                        x=sell_df["snapshot_date"],
                        y=sell_df["DeepTrend分數"],
                        mode="markers",
                        name="風險訊號",
                        marker=dict(color="#ef4444", symbol="triangle-down", size=13),
                    ),
                    secondary_y=False,
                )
        fig.update_layout(
            height=460,
            margin=dict(l=10, r=10, t=30, b=10),
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="left", x=0),
            xaxis_title="日期",
        )
        fig.update_yaxes(title_text="分數", secondary_y=False)
        fig.update_yaxes(title_text="股價", secondary_y=True, showgrid=False)
        st.plotly_chart(fig, use_container_width=True)

    detail_columns = [
        "snapshot_date",
        "股票代號",
        "股票名稱",
        "收盤價",
        "DeepTrend分數",
        "技術面分數",
        "籌碼分數",
        "量價分數",
        "前次分數",
        "分數變化",
        "分數變化率",
        "狀態程度",
        "買賣訊號",
        "狀態",
        "綜合判斷",
    ]
    detail_columns = [col for col in detail_columns if col in selected_df.columns]
    detail_df = selected_df[detail_columns].copy()
    detail_df["snapshot_date"] = detail_df["snapshot_date"].dt.strftime("%Y-%m-%d")
    detail_df = detail_df.rename(columns={"snapshot_date": "日期"})

    for col in ["收盤價", "DeepTrend分數", "技術面分數", "籌碼分數", "量價分數", "前次分數", "分數變化", "分數變化率"]:
        if col in detail_df.columns:
            detail_df[col] = detail_df[col].map(lambda value: "" if pd.isna(value) else format_number(value, 2))

    st.markdown("### 分數歷史明細")
    st.dataframe(detail_df.sort_values("日期", ascending=False), use_container_width=True, hide_index=True)

    st.markdown("### 分數組成明細")
    date_options = selected_df["snapshot_date"].dt.strftime("%Y-%m-%d").tolist()
    selected_date_label = st.selectbox(
        "選擇快照日期",
        date_options,
        index=len(date_options) - 1,
        key=f"score_breakdown_date_{selected_code}",
    )
    selected_snapshot = selected_df[selected_df["snapshot_date"].dt.strftime("%Y-%m-%d").eq(selected_date_label)].iloc[-1]
    st.caption(
        f"{selected_date_label}｜狀態：{selected_snapshot.get('狀態程度', '資料不足')}｜"
        f"{selected_snapshot.get('買賣訊號', '') or '無明確轉強/風險訊號'}"
    )

    breakdown_df = score_point_rows(selected_snapshot)
    if breakdown_df.empty:
        st.info("這筆快照目前沒有足夠欄位可拆解分數。")
    else:
        st.dataframe(breakdown_df, use_container_width=True, hide_index=True)

    csv_df = selected_df.copy()
    csv_df["snapshot_date"] = csv_df["snapshot_date"].dt.strftime("%Y-%m-%d")
    st.download_button(
        "下載此股票分數歷史 CSV",
        data=csv_df.to_csv(index=False, encoding="utf-8-sig"),
        file_name=f"score_history_{selected_code}.csv",
        mime="text/csv",
    )


SCORE_VALIDATION_BUCKETS = [
    ("DT < 0", None, 0),
    ("0 <= DT < 20", 0, 20),
    ("20 <= DT < 40", 20, 40),
    ("40 <= DT < 50", 40, 50),
    ("50 <= DT < 60", 50, 60),
    ("60 <= DT < 70", 60, 70),
    ("70 <= DT < 80", 70, 80),
    ("DT >= 80", 80, None),
]


def prepare_score_validation_forward_returns(history_df):
    """Use saved daily snapshots to calculate forward returns without downloading prices."""
    required_columns = {"snapshot_date", "股票代號", "股票名稱", "收盤價", "DeepTrend分數"}
    if history_df.empty or not required_columns.issubset(history_df.columns):
        return pd.DataFrame()

    work_df = history_df.copy()
    work_df["收盤價"] = pd.to_numeric(work_df["收盤價"], errors="coerce")
    work_df["DeepTrend分數"] = pd.to_numeric(work_df["DeepTrend分數"], errors="coerce")
    work_df = work_df.dropna(subset=["snapshot_date", "股票代號", "收盤價", "DeepTrend分數"])
    if work_df.empty:
        return pd.DataFrame()

    rows = []
    for _, stock_rows in work_df.groupby("股票代號"):
        stock_rows = stock_rows.sort_values("snapshot_date").drop_duplicates(subset=["snapshot_date"], keep="last")
        if stock_rows.empty:
            continue

        stock_rows = stock_rows.reset_index(drop=True)
        for horizon in [1, 3, 5, 10, 20]:
            future_close = stock_rows["收盤價"].shift(-horizon)
            stock_rows[f"{horizon}日後收盤價"] = future_close
            stock_rows[f"{horizon}日後報酬率"] = (future_close - stock_rows["收盤價"]) / stock_rows["收盤價"] * 100

        rows.append(stock_rows)

    if not rows:
        return pd.DataFrame()

    return pd.concat(rows, ignore_index=True)


def build_score_validation_result(history_df, score_threshold):
    """Calculate forward returns after each DeepTrend score threshold event."""
    work_df = prepare_score_validation_forward_returns(history_df)
    if work_df.empty:
        return pd.DataFrame()

    result_df = work_df[work_df["DeepTrend分數"] >= score_threshold].copy()
    if result_df.empty:
        return pd.DataFrame()

    return result_df.sort_values(["snapshot_date", "DeepTrend分數"], ascending=[False, False])


def score_bucket_mask(score_series, lower, upper):
    """Return a boolean mask for one DeepTrend score interval."""
    if lower is None:
        return score_series < upper
    if upper is None:
        return score_series >= lower
    return score_series.ge(lower) & score_series.lt(upper)


def score_bucket_label(score):
    """Assign one DeepTrend score to the configured validation interval label."""
    if pd.isna(score):
        return pd.NA
    for label, lower, upper in SCORE_VALIDATION_BUCKETS:
        if lower is None and score < upper:
            return label
        if upper is None and score >= lower:
            return label
        if lower is not None and upper is not None and lower <= score < upper:
            return label
    return pd.NA


def filter_score_interval_events(work_df):
    """Keep only the first snapshot when a stock enters a different DeepTrend score interval."""
    if work_df.empty:
        return work_df

    event_rows = []
    for _, stock_rows in work_df.groupby("股票代號"):
        stock_rows = stock_rows.sort_values("snapshot_date").drop_duplicates(subset=["snapshot_date"], keep="last")
        stock_rows = stock_rows.copy()
        stock_rows["分數區間"] = stock_rows["DeepTrend分數"].map(score_bucket_label)
        stock_rows = stock_rows.dropna(subset=["分數區間"])
        if stock_rows.empty:
            continue

        previous_bucket = stock_rows["分數區間"].shift(1)
        event_rows.append(stock_rows[stock_rows["分數區間"].ne(previous_bucket)])

    if not event_rows:
        return pd.DataFrame(columns=work_df.columns)
    return pd.concat(event_rows, ignore_index=True)


def summarize_score_validation_slice(signal_df):
    """Summarize a validation slice by signal count, unique stocks, returns, and win rates."""
    total_signals = len(signal_df)
    unique_stocks = signal_df["股票代號"].nunique() if "股票代號" in signal_df.columns else 0
    summary = {
        "總訊號數": total_signals,
        "獨立股票數": unique_stocks,
        "平均每檔觸發": total_signals / unique_stocks if unique_stocks else 0,
    }
    for horizon in [1, 3, 5, 10, 20]:
        col = f"{horizon}日後報酬率"
        valid_returns = pd.to_numeric(signal_df.get(col, pd.Series(dtype=float)), errors="coerce").dropna()
        summary[f"{horizon}日平均報酬"] = valid_returns.mean() if not valid_returns.empty else pd.NA
        summary[f"{horizon}日上漲率"] = valid_returns.gt(0).mean() * 100 if not valid_returns.empty else pd.NA
    return summary


def build_score_interval_validation(history_df, event_mode=False):
    """Compare forward returns across fixed DeepTrend score intervals."""
    work_df = prepare_score_validation_forward_returns(history_df)
    if work_df.empty:
        return pd.DataFrame()

    if event_mode:
        work_df = filter_score_interval_events(work_df)
        if work_df.empty:
            return pd.DataFrame()

    score_series = pd.to_numeric(work_df["DeepTrend分數"], errors="coerce")
    rows = []
    for label, lower, upper in SCORE_VALIDATION_BUCKETS:
        bucket_df = work_df[score_bucket_mask(score_series, lower, upper)].copy()
        summary = summarize_score_validation_slice(bucket_df)
        rows.append(
            {
                "區間": label,
                **summary,
            }
        )
    return pd.DataFrame(rows)


def validation_summary(result_df):
    """Summarize score validation signals into average returns and win rates."""
    if result_df.empty:
        return {}

    summary = {
        "總訊號數": len(result_df),
        "獨立股票數": result_df["股票代號"].nunique() if "股票代號" in result_df.columns else 0,
    }
    for horizon in [1, 3, 5, 10, 20]:
        col = f"{horizon}日後報酬率"
        if col not in result_df.columns:
            summary[f"{horizon}日樣本數"] = 0
            summary[f"{horizon}日平均報酬"] = pd.NA
            summary[f"{horizon}日上漲機率"] = pd.NA
            continue

        valid_returns = pd.to_numeric(result_df[col], errors="coerce").dropna()
        summary[f"{horizon}日樣本數"] = len(valid_returns)
        summary[f"{horizon}日平均報酬"] = valid_returns.mean() if not valid_returns.empty else pd.NA
        summary[f"{horizon}日上漲機率"] = (valid_returns.gt(0).mean() * 100) if not valid_returns.empty else pd.NA

    return summary


def build_signal_tracking(history_df, result_df):
    """Build an in-progress tracking table for recent DeepTrend threshold signals."""
    required_columns = {"snapshot_date", "股票代號", "股票名稱", "收盤價"}
    if history_df.empty or result_df.empty or not required_columns.issubset(history_df.columns):
        return pd.DataFrame()

    rows = []
    history_df = history_df.copy()
    history_df["收盤價"] = pd.to_numeric(history_df["收盤價"], errors="coerce")

    for _, signal in result_df.iterrows():
        stock_id = signal.get("股票代號")
        trigger_date = signal.get("snapshot_date")
        if pd.isna(stock_id) or pd.isna(trigger_date):
            continue

        stock_history = history_df[history_df["股票代號"].eq(stock_id)].copy()
        stock_history = stock_history.sort_values("snapshot_date").drop_duplicates(subset=["snapshot_date"], keep="last")
        stock_history = stock_history.reset_index(drop=True)
        if stock_history.empty:
            continue

        matched_indexes = stock_history.index[stock_history["snapshot_date"].eq(trigger_date)].tolist()
        if not matched_indexes:
            continue

        trigger_index = matched_indexes[-1]
        latest_index = len(stock_history) - 1
        elapsed = latest_index - trigger_index
        if elapsed < 0 or elapsed > 20:
            continue

        latest_row = stock_history.iloc[latest_index]
        trigger_close = pd.to_numeric(signal.get("收盤價"), errors="coerce")
        latest_close = pd.to_numeric(latest_row.get("收盤價"), errors="coerce")
        current_return = (
            (latest_close - trigger_close) / trigger_close * 100
            if pd.notna(trigger_close) and pd.notna(latest_close) and trigger_close
            else pd.NA
        )

        row = {
            "股票代號": stock_id,
            "股票名稱": signal.get("股票名稱", ""),
            "觸發日": trigger_date.strftime("%Y-%m-%d") if hasattr(trigger_date, "strftime") else str(trigger_date),
            "追蹤天數": elapsed,
            "觸發價": trigger_close,
            "目前價": latest_close,
            "目前報酬": current_return,
            "DeepTrend分數": signal.get("DeepTrend分數", pd.NA),
        }
        for horizon in [3, 5, 10, 20]:
            return_col = f"{horizon}日後報酬率"
            if elapsed >= horizon:
                value = pd.to_numeric(signal.get(return_col), errors="coerce")
                row[f"{horizon}日"] = format_signed_pct(value) if pd.notna(value) else "尚無資料"
            else:
                row[f"{horizon}日"] = "追蹤中"
        rows.append(row)

    if not rows:
        return pd.DataFrame()

    tracking_df = pd.DataFrame(rows)
    return tracking_df.sort_values(["觸發日", "DeepTrend分數"], ascending=[False, False]).head(30)


def render_signal_tracking(history_df, result_df):
    """Render currently active DeepTrend signal follow-up tracking."""
    st.subheader("📌 訊號後續追蹤")
    st.caption("追蹤最近 20 個快照內的達標訊號，觀察 3 / 5 / 10 / 20 個快照後的表現。")

    tracking_df = build_signal_tracking(history_df, result_df)
    if tracking_df.empty:
        st.info("目前沒有正在追蹤中的達標訊號。")
        return

    display_df = tracking_df.copy()
    for col in ["觸發價", "目前價", "DeepTrend分數"]:
        if col in display_df.columns:
            display_df[col] = display_df[col].map(lambda value: "" if pd.isna(value) else format_number(value, 2))
    if "目前報酬" in display_df.columns:
        display_df["目前報酬"] = display_df["目前報酬"].map(format_signed_pct)

    st.dataframe(display_df, use_container_width=True, hide_index=True)


def render_score_validation(stock_df):
    """Render validation statistics for DeepTrend threshold signals."""
    st.subheader("✅ 分數驗證")
    st.caption("用歷史快照驗證：DeepTrend 分數達標後，後續 1 / 3 / 5 / 10 / 20 個快照的平均表現。")

    history_df = load_score_history_data()
    if history_df.empty:
        st.warning("目前沒有可讀取的分數歷史資料，無法驗證 DeepTrend 分數。")
        return

    required_columns = {"snapshot_date", "股票代號", "股票名稱", "收盤價", "DeepTrend分數"}
    missing_columns = sorted(required_columns - set(history_df.columns))
    if missing_columns:
        st.warning(f"分數驗證缺少必要欄位：{', '.join(missing_columns)}")
        return

    valid_score_count = pd.to_numeric(history_df["DeepTrend分數"], errors="coerce").notna().sum()
    if valid_score_count == 0:
        st.info("目前歷史資料中還沒有 DeepTrend 分數。之後每日更新累積後，就能開始驗證。")
        return

    validation_mode = st.radio(
        "驗證模式",
        ["達標模式", "區間模式"],
        horizontal=True,
        help="達標模式看 DeepTrend 高於某個門檻後的表現；區間模式比較不同分數帶的後續表現。",
    )

    if validation_mode == "區間模式":
        st.info(
            "區間模式用固定 DeepTrend 分數帶比較後續表現，"
            "可以觀察分數是否越高越好，或是否存在 40~60 這類甜蜜區。"
        )
        interval_count_mode = st.radio(
            "區間統計方式",
            ["快照模式（目前）", "事件模式（建議）"],
            horizontal=True,
            help="快照模式會計算每一天符合區間的快照；事件模式只計算同一檔股票第一次進入該區間。",
        )
        use_event_mode = interval_count_mode.startswith("事件模式")
        if use_event_mode:
            st.caption("事件模式：同一檔股票連續待在同一區間只算一次，只有第一次進入該區間才觸發。")
        else:
            st.caption("快照模式：每一筆每日快照都列入統計，適合觀察分數狀態分布，但可能重複計算同一檔股票。")

        interval_df = build_score_interval_validation(history_df, event_mode=use_event_mode)
        if interval_df.empty:
            st.info("目前沒有足夠的分數歷史資料可做區間驗證。")
            return

        display_df = interval_df.copy()
        display_df["總訊號數"] = display_df["總訊號數"].map(format_integer)
        display_df["獨立股票數"] = display_df["獨立股票數"].map(format_integer)
        display_df["平均每檔觸發"] = display_df["平均每檔觸發"].map(lambda value: format_number(value, 2))
        for horizon in [1, 3, 5, 10, 20]:
            avg_col = f"{horizon}日平均報酬"
            win_col = f"{horizon}日上漲率"
            display_df[avg_col] = display_df[avg_col].map(
                lambda value: f"{format_number(value, 2)}%" if pd.notna(value) else "N/A"
            )
            display_df[win_col] = display_df[win_col].map(
                lambda value: f"{format_number(value, 1)}%" if pd.notna(value) else "N/A"
            )

        st.markdown("### 分數區間比較表")
        st.dataframe(display_df, use_container_width=True, hide_index=True)

        chart_source = interval_df[["區間", "1日平均報酬", "3日平均報酬", "5日平均報酬"]].copy()
        chart_source = chart_source.melt(id_vars="區間", var_name="觀察天數", value_name="平均報酬率")
        chart_source = chart_source.dropna(subset=["平均報酬率"])
        if chart_source.empty:
            st.info("目前區間資料還不足，暫不繪製平均報酬圖。")
        else:
            fig = go.Figure()
            for horizon_label in ["1日平均報酬", "3日平均報酬", "5日平均報酬"]:
                horizon_df = chart_source[chart_source["觀察天數"].eq(horizon_label)]
                fig.add_trace(
                    go.Bar(
                        x=horizon_df["區間"],
                        y=horizon_df["平均報酬率"],
                        name=horizon_label.replace("平均報酬", "日後"),
                    )
                )
            fig.update_layout(
                height=360,
                margin=dict(l=10, r=10, t=30, b=10),
                yaxis_title="平均報酬率 (%)",
                barmode="group",
            )
            st.plotly_chart(fig, use_container_width=True)

        st.download_button(
            "下載分數區間驗證 CSV",
            data=interval_df.to_csv(index=False, encoding="utf-8-sig"),
            file_name=f"score_interval_validation_{'event' if use_event_mode else 'snapshot'}.csv",
            mime="text/csv",
        )
        return

    col1, col2 = st.columns([1, 2])
    with col1:
        score_threshold = st.number_input(
            "DeepTrend 達標分數",
            min_value=0.0,
            value=60.0,
            step=5.0,
            help="每筆歷史快照只要 DeepTrend 分數大於等於此門檻，就列入驗證樣本。",
        )
    with col2:
        st.info("這裡使用已儲存的每日快照計算，不重新下載行情。若歷史資料還少，10日/20日後樣本數會先偏少。")

    result_df = build_score_validation_result(history_df, score_threshold)
    if result_df.empty:
        st.info("目前沒有符合這個分數門檻的歷史訊號。可以先降低達標分數，或等待歷史資料累積。")
        return

    summary = validation_summary(result_df)
    total_signals = summary.get("總訊號數", 0)
    unique_stocks = summary.get("獨立股票數", 0)
    avg_per_stock = total_signals / unique_stocks if unique_stocks else 0

    top_cols = st.columns(3)
    top_cols[0].metric("總訊號數", format_integer(total_signals))
    top_cols[1].metric("獨立股票數", format_integer(unique_stocks))
    top_cols[2].metric("平均每檔觸發", format_number(avg_per_stock, 2))

    summary_rows = []
    for horizon in [1, 3, 5, 10, 20]:
        summary_rows.append(
            {
                "觀察天數": f"{horizon}個快照後",
                "有效樣本數": summary.get(f"{horizon}日樣本數", 0),
                "平均報酬率": summary.get(f"{horizon}日平均報酬", pd.NA),
                "上漲機率": summary.get(f"{horizon}日上漲機率", pd.NA),
            }
        )

    summary_df = pd.DataFrame(summary_rows)
    summary_display = summary_df.copy()
    summary_display["有效樣本數"] = summary_display["有效樣本數"].map(format_integer)
    summary_display["平均報酬率"] = summary_display["平均報酬率"].map(lambda value: f"{format_number(value, 2)}%" if pd.notna(value) else "")
    summary_display["上漲機率"] = summary_display["上漲機率"].map(lambda value: f"{format_number(value, 1)}%" if pd.notna(value) else "")

    st.markdown("### 驗證摘要")
    st.dataframe(summary_display, use_container_width=True, hide_index=True)

    max_sample_count = int(pd.to_numeric(summary_df["有效樣本數"], errors="coerce").fillna(0).max())
    if max_sample_count < 10:
        st.warning(
            "目前分數歷史樣本還太少，統計結果僅供觀察。"
            "建議至少累積 10 筆以上有效樣本，再判斷 DeepTrend 分數是否可靠。"
        )

    chart_df = summary_df.dropna(subset=["平均報酬率"]).copy()
    chart_df = chart_df[pd.to_numeric(chart_df["有效樣本數"], errors="coerce").fillna(0) >= 3]
    if chart_df.empty:
        st.info("目前各觀察天數的有效樣本少於 3 筆，暫不繪製平均報酬圖，避免視覺誤導。")
    else:
        fig = go.Figure()
        fig.add_trace(
            go.Bar(
                x=chart_df["觀察天數"],
                y=chart_df["平均報酬率"],
                name="平均報酬率",
                marker_color="#38bdf8",
            )
        )
        fig.update_layout(
            height=320,
            margin=dict(l=10, r=10, t=30, b=10),
            yaxis_title="平均報酬率 (%)",
            showlegend=False,
        )
        st.plotly_chart(fig, use_container_width=True)

    detail_columns = [
        "snapshot_date",
        "股票代號",
        "股票名稱",
        "收盤價",
        "DeepTrend分數",
        "1日後報酬率",
        "3日後報酬率",
        "5日後報酬率",
        "10日後報酬率",
        "20日後報酬率",
        "狀態",
        "綜合判斷",
    ]
    detail_columns = [col for col in detail_columns if col in result_df.columns]
    detail_df = result_df[detail_columns].copy()
    detail_df["snapshot_date"] = detail_df["snapshot_date"].dt.strftime("%Y-%m-%d")
    detail_df = detail_df.rename(columns={"snapshot_date": "觸發日期"})

    for col in ["收盤價", "DeepTrend分數", "1日後報酬率", "3日後報酬率", "5日後報酬率", "10日後報酬率", "20日後報酬率"]:
        if col in detail_df.columns:
            suffix = "%" if "報酬率" in col else ""
            detail_df[col] = detail_df[col].map(lambda value: f"{format_number(value, 2)}{suffix}" if pd.notna(value) else "")

    with st.expander("查看達標訊號明細（最多顯示前 300 筆）"):
        st.dataframe(detail_df.head(300), use_container_width=True, hide_index=True)

    render_signal_tracking(history_df, result_df)

    csv_df = result_df.copy()
    csv_df["snapshot_date"] = csv_df["snapshot_date"].dt.strftime("%Y-%m-%d")
    st.download_button(
        "下載分數驗證結果 CSV",
        data=csv_df.to_csv(index=False, encoding="utf-8-sig"),
        file_name=f"score_validation_{int(score_threshold)}.csv",
        mime="text/csv",
    )


def file_modified_text(path):
    """Return a readable modified timestamp for a local data file."""
    if not path.exists():
        return ""
    return datetime.fromtimestamp(path.stat().st_mtime).strftime("%Y-%m-%d %H:%M:%S")


def latest_csv_date(path, column):
    """Read a CSV date column safely and return the latest date string."""
    if not path.exists():
        return ""
    try:
        file_df = pd.read_csv(path, usecols=[column])
    except Exception:
        return ""
    dates = pd.to_datetime(file_df[column], errors="coerce").dropna()
    if dates.empty:
        return ""
    return dates.max().strftime("%Y-%m-%d")


def render_data_health(stock_df):
    """Render local data freshness checks for the daily update pipeline."""
    st.subheader("🩺 資料健康檢查")
    st.caption("檢查本地/雲端資料檔是否有更新，方便確認 DeepTrend 今天的資料狀態。")

    today_text = date.today().isoformat()
    result_modified = file_modified_text(RESULT_FILE)
    result_modified_date = result_modified[:10] if result_modified else ""
    chip_latest_date = latest_csv_date(CHIP_DAILY_FILE, "date")
    history_latest_date = latest_csv_date(STOCK_ANALYSIS_HISTORY_FILE, "snapshot_date")

    rows = [
        {
            "檢查項目": "stock_analysis_result.xlsx 是否今天更新",
            "目前狀態": "✅ 今天已更新" if result_modified_date == today_text else "⚠️ 不是今天",
            "最新日期/時間": result_modified or "找不到檔案",
        },
        {
            "檢查項目": "chip_daily.csv 是否有今天資料",
            "目前狀態": "✅ 今天有資料" if chip_latest_date == today_text else "⚠️ 尚未看到今天",
            "最新日期/時間": chip_latest_date or "找不到日期",
        },
        {
            "檢查項目": "stock_analysis_history.csv 是否有今天快照",
            "目前狀態": "✅ 今天有快照" if history_latest_date == today_text else "⚠️ 尚未看到今天",
            "最新日期/時間": history_latest_date or "找不到日期",
        },
    ]
    st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)

    watchlist_file = BASE_DIR / "watchlist.csv"
    if not watchlist_file.exists():
        st.warning("找不到 watchlist.csv，無法比對缺漏股票。")
        return
    if stock_df.empty or "股票代號" not in stock_df.columns:
        st.warning("目前股票分析結果為空，無法比對缺漏股票。")
        return

    try:
        watchlist_df = pd.read_csv(watchlist_file)
    except Exception as exc:
        st.warning(f"watchlist.csv 讀取失敗：{exc}")
        return

    if "ticker" not in watchlist_df.columns:
        st.warning("watchlist.csv 缺少 ticker 欄位，無法比對缺漏股票。")
        return

    expected_codes = watchlist_df["ticker"].astype(str).map(stock_code_key)
    actual_codes = stock_df["股票代號"].astype(str).map(stock_code_key)
    missing_mask = ~expected_codes.isin(set(actual_codes.dropna()))
    missing_df = watchlist_df.loc[missing_mask].copy()

    st.markdown("### 抓不到資料股票")
    if missing_df.empty:
        st.success("目前 watchlist 內股票都有出現在 stock_analysis_result.xlsx。")
    else:
        display_cols = [col for col in ["ticker", "name", "group"] if col in missing_df.columns]
        st.warning(f"目前有 {len(missing_df)} 檔 watchlist 股票沒有出現在分析結果。")
        st.dataframe(missing_df[display_cols] if display_cols else missing_df, use_container_width=True, hide_index=True)


def render_scan_table(filtered_df):
    """Render the full detailed stock table with formatting and status coloring."""
    st.subheader("📋 詳細表格")
    st.write(f"目前顯示 {len(filtered_df)} 檔股票")

    if filtered_df.empty:
        st.info("目前沒有符合篩選條件的股票。")
        return

    display_df = filtered_df.copy()

    if "資料時間" in display_df.columns:
        updated_count = display_df["資料時間"].astype(str).str.len().gt(0).sum()
        latest_times = sorted(
            display_df["資料時間"].dropna().astype(str).loc[
                display_df["資料時間"].dropna().astype(str).str.len() > 0
            ].unique().tolist()
        )
        latest_time_text = latest_times[-1] if latest_times else "尚未取得"
        st.caption(f"即時資料更新：{updated_count}/{len(display_df)} 檔，最新時間 {latest_time_text}")

    display_df = display_df.drop(columns=["資料時間"], errors="ignore")

    front_columns = [
        "狀態",
        "股票代號",
        "股票名稱",
        "均線型態",
        "突破警報",
        "收盤價",
        "今日漲跌幅",
        "乖離率",
        "量價異常",
        "外資5日",
        "投信5日",
    ]
    ordered_columns = [col for col in front_columns if col in display_df.columns]
    ordered_columns += [col for col in display_df.columns if col not in ordered_columns]
    display_df = display_df[ordered_columns]

    price_columns = ["收盤價", "今日漲跌幅", "5日線", "10日線", "20日線", "20日高點", "20日低點", "乖離率"]
    chip_columns = [
        "籌碼1日",
        "籌碼3日",
        "籌碼5日",
        "籌碼10日",
        "外資1日",
        "外資3日",
        "外資5日",
        "外資10日",
        "投信1日",
        "投信3日",
        "投信5日",
        "投信10日",
        "自營商1日",
        "自營商3日",
        "自營商5日",
        "自營商10日",
    ]

    for col in chip_columns:
        if col in display_df.columns:
            display_df[col] = display_df[col].map(format_integer)

    for col in price_columns:
        if col in display_df.columns:
            display_df[col] = display_df[col].map(format_number)

    for col in ["成交量", "5日均量"]:
        if col in display_df.columns:
            display_df[col] = display_df[col].map(format_integer)

    styled_df = display_df.style.map(color_status, subset=["狀態"])
    st.dataframe(styled_df, use_container_width=True, hide_index=True)


def debug_kline_data(k_df):
    """Print and export recent OHLCV rows used by the K-line chart for inspection."""
    debug_df = k_df.copy().sort_index()
    original_columns = list(debug_df.columns)
    print(debug_df.columns.tolist())
    debug_df = debug_df.reset_index()

    column_aliases = {
        "Date": ["Date", "Datetime", "日期", "index"],
        "Open": ["Open", "open", "開盤價", "開盤"],
        "High": ["High", "high", "最高價", "最高"],
        "Low": ["Low", "low", "最低價", "最低"],
        "Close": ["Close", "close", "收盤價", "收盤", "Close_Price"],
        "Volume": ["Volume", "volume", "成交量"],
    }
    column_map = {}
    for target, aliases in column_aliases.items():
        matched_column = next((col for col in aliases if col in debug_df.columns), None)
        if matched_column is None and target == "Date":
            matched_column = debug_df.columns[0]
        if matched_column is not None:
            column_map[matched_column] = target

    debug_df = debug_df.rename(columns=column_map)
    debug_columns = ["Date", "Open", "High", "Low", "Close", "Volume"]
    missing_columns = [col for col in debug_columns if col not in debug_df.columns]
    if missing_columns:
        print("K線 debug 原始欄位：", original_columns)
        print("K線 debug reset_index 後欄位：", list(debug_df.columns))
        print("K線 debug 已對應欄位：", column_map)
        print("K線 debug 缺少欄位：", missing_columns)

    available_columns = [col for col in debug_columns if col in debug_df.columns]
    debug_df = debug_df[available_columns]
    if "Date" in debug_df.columns:
        debug_df["Date"] = pd.to_datetime(debug_df["Date"], errors="coerce").dt.strftime("%Y-%m-%d")

    print(debug_df.tail(10)[available_columns])
    debug_path = BASE_DIR / "debug_kline.csv"
    debug_csv = debug_df.tail(30).to_csv(index=False, encoding="utf-8-sig")
    debug_path.write_text(debug_csv, encoding="utf-8-sig")
    st.caption(f"K線 debug 檔已產生：{debug_path}")
    st.download_button(
        "下載 debug_kline.csv",
        data=debug_csv.encode("utf-8-sig"),
        file_name="debug_kline.csv",
        mime="text/csv",
    )


def render_detail(filtered_df):
    """Render single-stock summary cards plus the K-line chart."""
    if filtered_df.empty:
        st.info("請調整篩選條件後再查看個股。")
        return

    stock_options = filtered_df["股票代號"].astype(str).tolist()
    preferred_stock = str(st.session_state.pop("pending_detail_stock", ""))
    current_detail_stock = str(st.session_state.get("detail_stock", ""))

    if preferred_stock in stock_options:
        st.session_state["detail_stock"] = preferred_stock
    elif current_detail_stock not in stock_options:
        st.session_state["detail_stock"] = stock_options[0]

    selected_stock = st.selectbox(
        "選擇股票查看K線",
        stock_options,
        format_func=build_stock_label_map(filtered_df).get,
        key="detail_stock",
    )
    st.session_state["selected_detail_stock"] = selected_stock
    selected_row = filtered_df[filtered_df["股票代號"].astype(str) == selected_stock].iloc[0]

    st.markdown("## 🔎 個股分析摘要")

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("股票名稱", selected_row["股票名稱"])
    col2.metric("技術分數", selected_row["技術分數"])
    col3.metric("狀態", selected_row["狀態"])
    col4.metric("綜合判斷", selected_row["綜合判斷"])

    signal_col1, signal_col2 = st.columns(2)
    signal_col1.metric("均線型態", selected_row.get("均線型態", "中性"))
    signal_col2.metric("突破警報", selected_row.get("突破警報", "無"))

    st.markdown("### 📈 分數組成變化")
    score_change_df = build_score_component_change(selected_row)
    st.dataframe(score_change_df, use_container_width=True, hide_index=True)
    if score_change_df["昨天"].eq("").all():
        st.caption("目前尚無前一個新版 DeepTrend 快照可比較，累積一天後會顯示昨天分數。")

    st.markdown("### 📌 技術面")
    st.info(selected_row["技術面"])

    st.markdown("### 💰 籌碼面")
    st.success(selected_row["籌碼面"])

    symbol = normalize_tw_symbol(selected_stock)
    k_df = download_market_data(symbol)
    chart_mode = "candlestick"

    if k_df.empty:
        history = get_official_daily_history(symbol)
        k_df = official_history_to_kline(history)
        chart_mode = "close_line"

    if k_df.empty:
        st.warning(f"抓不到 {symbol} 的K線資料。")
        return

    close_series = get_series(k_df, "Close")

    k_df["MA5"] = close_series.rolling(5).mean()
    k_df["MA10"] = close_series.rolling(10).mean()
    k_df["MA20"] = close_series.rolling(20).mean()

    delta = close_series.diff()
    gain = delta.clip(lower=0)
    loss = -delta.clip(upper=0)
    avg_gain = gain.rolling(14).mean()
    avg_loss = loss.rolling(14).mean()
    rs = avg_gain / avg_loss
    k_df["RSI"] = 100 - (100 / (1 + rs))
    k_df = k_df[k_df["MA20"].notna() & k_df["RSI"].notna()]

    if k_df.empty:
        st.warning("K線資料不足，無法計算均線與RSI。")
        return

    debug_kline_data(k_df)
    render_k_chart(k_df, chart_mode=chart_mode)


def run_ma_backtest(history, holding_days=5, volume_multiplier=1.5):
    """Run the older MA/KD strategy used by the strategy ranking page."""
    if history.empty or len(history) < 30:
        return pd.DataFrame()

    test_df = history.copy().sort_values("日期").reset_index(drop=True)
    test_df["收盤價"] = pd.to_numeric(test_df["收盤價"], errors="coerce")
    test_df["最高價"] = pd.to_numeric(test_df["最高價"], errors="coerce")
    test_df["最低價"] = pd.to_numeric(test_df["最低價"], errors="coerce")
    test_df["成交量"] = pd.to_numeric(test_df["成交量"], errors="coerce")
    test_df["MA5"] = test_df["收盤價"].rolling(5).mean()
    test_df["MA10"] = test_df["收盤價"].rolling(10).mean()
    test_df["5日均量"] = test_df["成交量"].rolling(5).mean()
    low_9 = test_df["最低價"].rolling(9).min()
    high_9 = test_df["最高價"].rolling(9).max()
    price_range = high_9 - low_9
    test_df["RSV"] = ((test_df["收盤價"] - low_9) / price_range.replace(0, pd.NA) * 100).fillna(50)
    test_df["K值"] = test_df["RSV"].ewm(alpha=1 / 3, adjust=False).mean()
    test_df["D值"] = test_df["K值"].ewm(alpha=1 / 3, adjust=False).mean()
    test_df["prev_K值"] = test_df["K值"].shift(1)
    test_df["prev_D值"] = test_df["D值"].shift(1)
    # 舊策略進場公式：5MA > 10MA + 量能放大 + KD 黃金交叉。
    test_df["signal"] = (
        (test_df["MA5"] > test_df["MA10"])
        & (test_df["成交量"] > test_df["5日均量"] * volume_multiplier)
        & (test_df["prev_K值"] <= test_df["prev_D值"])
        & (test_df["K值"] > test_df["D值"])
    )

    trades = []
    for index, row in test_df[test_df["signal"]].iterrows():
        exit_index = index + holding_days
        if exit_index >= len(test_df):
            continue

        entry_price = float(row["收盤價"])
        exit_price = float(test_df.loc[exit_index, "收盤價"])
        if not entry_price:
            continue

        holding_window = test_df.loc[index:exit_index, "收盤價"].dropna()
        running_peak = holding_window.cummax()
        drawdown = (holding_window - running_peak) / running_peak * 100
        max_drawdown = float(drawdown.min()) if not drawdown.empty else 0

        trades.append(
            {
                "進場日": row["日期"],
                "進場價": entry_price,
                "成交量倍率": round(float(row["成交量"] / row["5日均量"]), 2) if row["5日均量"] else 0,
                "K值": round(float(row["K值"]), 2),
                "D值": round(float(row["D值"]), 2),
                "出場日": test_df.loc[exit_index, "日期"],
                "出場價": exit_price,
                "報酬率": round((exit_price - entry_price) / entry_price * 100, 2),
                "最大回撤": round(max_drawdown, 2),
            }
        )

    return pd.DataFrame(trades)


def backtest_confidence(trade_count):
    """Classify backtest reliability based on sample count."""
    if trade_count < 10:
        return "🔴 低信賴"
    if trade_count < 30:
        return "🟡 中信賴"
    return "🟢 高信賴"


def summarize_backtest_trades(trades):
    """Summarize older strategy trades into win rate, average return, drawdown, and payoff ratio."""
    if trades.empty:
        return {
            "交易次數": 0,
            "勝率": 0,
            "平均報酬": 0,
            "最佳報酬": 0,
            "最差報酬": 0,
            "平均回撤": 0,
            "最大回撤": 0,
            "平均獲利": 0,
            "平均虧損": 0,
            "盈虧比": None,
            "信賴度": backtest_confidence(0),
        }

    trade_count = len(trades)
    winning_returns = trades.loc[trades["報酬率"] > 0, "報酬率"]
    losing_returns = trades.loc[trades["報酬率"] < 0, "報酬率"]
    avg_win = winning_returns.mean() if not winning_returns.empty else 0
    avg_loss = losing_returns.mean() if not losing_returns.empty else 0

    return {
        "交易次數": trade_count,
        "勝率": (trades["報酬率"] > 0).mean() * 100,
        "平均報酬": trades["報酬率"].mean(),
        "最佳報酬": trades["報酬率"].max(),
        "最差報酬": trades["報酬率"].min(),
        "平均回撤": trades["最大回撤"].mean(),
        "最大回撤": trades["最大回撤"].min(),
        "平均獲利": avg_win,
        "平均虧損": avg_loss,
        "盈虧比": (avg_win / abs(avg_loss)) if avg_loss < 0 else None,
        "信賴度": backtest_confidence(trade_count),
    }


def backtest_period_months(period_label):
    """Translate UI period labels into month counts for historical data fetching."""
    return {
        "6個月": 6,
        "1年": 12,
        "2年": 24,
        "3年": 36,
        "5年": 60,
    }.get(period_label, 6)


def get_current_entry_status(history, market_is_bullish, volume_multiplier=1.5):
    """Check whether the latest K-line state is close to the current entry-condition checklist."""
    if history.empty or len(history) < 10:
        return None

    latest_df = history.copy().sort_values("日期").reset_index(drop=True)
    latest_df["收盤價"] = pd.to_numeric(latest_df["收盤價"], errors="coerce")
    latest_df["最高價"] = pd.to_numeric(latest_df["最高價"], errors="coerce")
    latest_df["最低價"] = pd.to_numeric(latest_df["最低價"], errors="coerce")
    latest_df["成交量"] = pd.to_numeric(latest_df["成交量"], errors="coerce")
    latest_df["MA5"] = latest_df["收盤價"].rolling(5).mean()
    latest_df["MA10"] = latest_df["收盤價"].rolling(10).mean()
    latest_df["5日均量"] = latest_df["成交量"].rolling(5).mean()
    low_9 = latest_df["最低價"].rolling(9).min()
    high_9 = latest_df["最高價"].rolling(9).max()
    price_range = high_9 - low_9
    latest_df["RSV"] = ((latest_df["收盤價"] - low_9) / price_range.replace(0, pd.NA) * 100).fillna(50)
    latest_df["K值"] = latest_df["RSV"].ewm(alpha=1 / 3, adjust=False).mean()
    latest_df["D值"] = latest_df["K值"].ewm(alpha=1 / 3, adjust=False).mean()
    latest_df["prev_K值"] = latest_df["K值"].shift(1)
    latest_df["prev_D值"] = latest_df["D值"].shift(1)

    latest = latest_df.iloc[-1]
    volume_ratio = latest["成交量"] / latest["5日均量"] if latest["5日均量"] else 0
    ma_ok = bool(latest["MA5"] > latest["MA10"])
    volume_ok = bool(volume_ratio >= volume_multiplier)
    kd_ok = bool(latest["prev_K值"] <= latest["prev_D值"] and latest["K值"] > latest["D值"])

    return {
        "ma_ok": ma_ok,
        "volume_ok": volume_ok,
        "kd_ok": kd_ok,
        "market_ok": bool(market_is_bullish),
        "volume_ratio": volume_ratio,
    }


def render_entry_status_card(stock_name, status):
    """Render the checklist card that explains which entry conditions are currently satisfied."""
    if not status:
        st.info("目前日 K 資料不足，無法判斷是否接近歷史進場條件。")
        return

    rows = [
        ("量能達標" if status["volume_ok"] else "成交量不足", status["volume_ok"]),
        ("MA突破", status["ma_ok"]),
        ("KD黃金交叉", status["kd_ok"]),
        ("台指偏多", status["market_ok"]),
    ]
    checklist = "<br>".join(f"{'☑' if is_ok else '☐'} {label}" for label, is_ok in rows)
    matched_count = sum(is_ok for _, is_ok in rows)

    st.markdown(
        f"""
        <div style="
            padding:16px;
            border:1px solid #2f3542;
            border-radius:8px;
            background:#111827;
            margin:12px 0;
            color:#d1d5db;
        ">
            <div style="font-size:16px;font-weight:800;color:#ffffff;margin-bottom:8px;">
                目前是否接近歷史最佳買點
            </div>
            <div style="font-size:15px;line-height:1.8;">
                <b>{stock_name}</b><br>
                目前狀態：距離歷史回測進場條件還差：<br>
                {checklist}
            </div>
            <div style="margin-top:8px;color:#9ca3af;font-size:13px;">
                已符合 {matched_count} / {len(rows)} 項，成交量目前約 {status["volume_ratio"]:.2f} 倍。
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def prepare_breakout_history(history):
    """Prepare official daily K data with moving averages and 20-day average volume for breakout tests."""
    if history.empty:
        return pd.DataFrame()

    required_columns = {"日期", "收盤價", "最高價", "最低價", "成交量"}
    if not required_columns.issubset(history.columns):
        return pd.DataFrame()

    prepared = history.copy().sort_values("日期").reset_index(drop=True)
    for col in ["收盤價", "最高價", "最低價", "成交量"]:
        prepared[col] = pd.to_numeric(prepared[col], errors="coerce")

    prepared["MA5"] = prepared["收盤價"].rolling(5).mean()
    prepared["MA10"] = prepared["收盤價"].rolling(10).mean()
    prepared["MA20"] = prepared["收盤價"].rolling(20).mean()
    prepared["MA120"] = prepared["收盤價"].rolling(120).mean()
    prepared["20日均量"] = prepared["成交量"].rolling(20).mean()
    return prepared.dropna(subset=["日期", "收盤價"]).reset_index(drop=True)


@st.cache_data(ttl=900)
def get_market_temperature_history(ticker):
    """Return a lightweight 120-trading-day history for the observation-pool thermometer."""
    history = get_official_daily_history(ticker, month_count=6)
    history = prepare_breakout_history(history)
    if history.empty:
        return history
    return history.tail(120).reset_index(drop=True)


POSITIVE_NEWS_KEYWORDS = [
    "AI",
    "NVIDIA",
    "GB300",
    "伺服器",
    "記憶體",
    "漲價",
    "擴產",
    "接單",
    "轉機",
    "營收成長",
    "法說",
]

RISK_NEWS_KEYWORDS = [
    "虧損",
    "下修",
    "衰退",
    "裁員",
    "停工",
    "違約",
    "訴訟",
    "調查",
    "減資",
    "處分",
]


def empty_news_signal():
    """Return the neutral news signal used when RSS is disabled or no titles are found."""
    return {
        "新聞熱度": "無近期新聞",
        "news_count": 0,
        "news_titles": "無近期新聞",
        "positive_keywords": "無",
        "risk_keywords": "無",
        "news_score": 0,
    }


@st.cache_data(ttl=3600)
def fetch_recent_news_titles(ticker, name):
    """Fetch recent Google News RSS titles for a stock without API keys or full-text analysis."""
    code = str(ticker).split(".")[0]
    query = quote_plus(f"{name} {code} 台股")
    url = f"https://news.google.com/rss/search?q={query}&hl=zh-TW&gl=TW&ceid=TW:zh-Hant"
    cutoff = datetime.now() - timedelta(days=7)

    try:
        response = requests.get(url, timeout=8, headers={"User-Agent": "Mozilla/5.0"})
        if response.status_code != 200:
            return []
        root = ET.fromstring(response.content)
        titles = []
        for item in root.findall(".//item"):
            title = (item.findtext("title") or "").strip()
            pub_date_text = (item.findtext("pubDate") or "").strip()
            if not title:
                continue
            try:
                pub_date = parsedate_to_datetime(pub_date_text).replace(tzinfo=None)
                if pub_date < cutoff:
                    continue
            except Exception:
                pass
            titles.append(title)
        return titles[:20]
    except Exception:
        return []


def analyze_stock_news(ticker, name, enable_news=False):
    """Count recent RSS titles and match positive/risk keywords to produce a lightweight news score."""
    if not enable_news:
        return empty_news_signal()

    titles = fetch_recent_news_titles(ticker, name)
    if not titles:
        return empty_news_signal()

    text = " ".join(titles).upper()
    positive_hits = [keyword for keyword in POSITIVE_NEWS_KEYWORDS if keyword.upper() in text]
    risk_hits = [keyword for keyword in RISK_NEWS_KEYWORDS if keyword.upper() in text]
    score = min(10, max(0, len(positive_hits) * 2 + min(len(titles), 5) - len(risk_hits) * 2))

    return {
        "新聞熱度": f"{score}/10（{len(titles)}則）",
        "news_count": len(titles),
        "news_titles": "｜".join(titles[:5]),
        "positive_keywords": "、".join(positive_hits) if positive_hits else "無",
        "risk_keywords": "、".join(risk_hits) if risk_hits else "無",
        "news_score": score,
    }


def get_eps_value(row):
    """Sum EPS-like columns if present; returns None when EPS data is unavailable."""
    eps_columns = [col for col in row.index if "EPS" in str(col).upper()]
    if not eps_columns:
        return None

    values = pd.to_numeric(row[eps_columns], errors="coerce")
    if values.dropna().empty:
        return None

    return float(values.dropna().sum())


@st.cache_data(ttl=3600)
def get_eps_columns(columns):
    """Cache discovery of EPS columns because most app runs have no EPS data yet."""
    return [col for col in columns if "EPS" in str(col).upper()]


def days_until_break_ma5(history, start_index, max_days=20):
    """Find how many trading days after entry the close first falls below MA5."""
    for offset in range(1, max_days + 1):
        check_index = start_index + offset
        if check_index >= len(history):
            break
        close = history.loc[check_index, "收盤價"]
        ma5 = history.loc[check_index, "MA5"]
        if pd.notna(close) and pd.notna(ma5) and close < ma5:
            return offset
    return "20日內未跌破"


def calc_forward_return(history, start_index, days):
    """Calculate fixed-horizon return after a trigger date."""
    target_index = start_index + days
    if target_index >= len(history):
        return None

    entry_price = history.loc[start_index, "收盤價"]
    exit_price = history.loc[target_index, "收盤價"]
    if not entry_price or pd.isna(entry_price) or pd.isna(exit_price):
        return None

    return round((exit_price - entry_price) / entry_price * 100, 2)


def run_breakout_backtest_for_stock(row, settings):
    """Scan one stock's history for breakout triggers and collect forward-return outcomes."""
    ticker = str(row["股票代號"])
    name = str(row["股票名稱"])
    history = get_official_daily_history(ticker, month_count=settings["month_count"])
    history = prepare_breakout_history(history)

    required_days = max(settings["breakout_days"], 20) + 21
    if settings.get("require_above_ma120"):
        required_days = max(required_days, 141)
    if history.empty or len(history) < required_days:
        return []

    eps_value = get_eps_value(row)
    if settings["require_eps"] and (eps_value is None or eps_value <= 0):
        return []

    results = []
    breakout_days = settings["breakout_days"]
    breakout_method = settings.get("breakout_method", "收盤創高")
    price_min = settings["price_min"]
    price_max = settings["price_max"]
    volume_multiplier = settings["volume_multiplier"]
    cooldown_days = int(settings.get("cooldown_days", 20))
    last_trigger_index = None
    news_signal = analyze_stock_news(ticker, name, enable_news=settings.get("enable_news", False))

    # 新版回測進場公式：
    # 1. 收盤/盤中價格突破指定日數高點。
    # 2. 依使用者勾選，檢查多頭排列、量能放大、EPS、半年線與股價區間。
    # 3. 觸發後只追蹤固定 5/10/20 日報酬與 20 日內風險/潛力。
    start_index = max(breakout_days, 120 if settings.get("require_above_ma120") else 20)
    for index in range(start_index, len(history) - 20):
        if last_trigger_index is not None and index <= last_trigger_index + cooldown_days:
            continue

        current = history.loc[index]
        close = current["收盤價"]
        previous_high_col = "最高價" if breakout_method == "盤中創高" else "收盤價"
        trigger_price = current["最高價"] if breakout_method == "盤中創高" else close
        previous_high = history.loc[index - breakout_days : index - 1, previous_high_col].max()

        if pd.isna(close) or pd.isna(trigger_price) or trigger_price <= previous_high:
            continue
        if close < price_min or close > price_max:
            continue
        if settings["require_ma_alignment"] and not (current["MA5"] > current["MA10"] > current["MA20"]):
            continue
        if settings["require_volume"] and not (current["成交量"] > current["20日均量"] * volume_multiplier):
            continue
        if settings.get("require_above_ma120") and not (pd.notna(current["MA120"]) and close > current["MA120"]):
            continue

        return_5d = calc_forward_return(history, index, 5)
        return_10d = calc_forward_return(history, index, 10)
        return_20d = calc_forward_return(history, index, 20)
        if return_5d is None or return_10d is None or return_20d is None:
            continue

        forward_20 = history.loc[index + 1 : index + 20]
        if forward_20.empty or len(forward_20) < 20:
            continue

        max_rise = (forward_20["最高價"].max() - close) / close * 100 if not forward_20.empty else None
        max_pullback = (forward_20["最低價"].min() - close) / close * 100 if not forward_20.empty else None
        max_high_20 = forward_20["最高價"].max() if not forward_20.empty else None

        results.append(
            {
                "股票代號": ticker,
                "股票名稱": name,
                "觸發日期": current["日期"],
                "觸發收盤價": round(float(close), 2),
                "5日後報酬率": return_5d,
                "10日後報酬率": return_10d,
                "20日後報酬率": return_20d,
                "幾天後跌破5日線": days_until_break_ma5(history, index, max_days=20),
                "20日內最高價": round(float(max_high_20), 2) if pd.notna(max_high_20) else None,
                "觸發後20日內最大漲幅": round(float(max_rise), 2) if pd.notna(max_rise) else None,
                "觸發後20日內最大回檔": round(float(max_pullback), 2) if pd.notna(max_pullback) else None,
                **news_signal,
            }
        )
        last_trigger_index = index

    return results


@st.cache_data(ttl=900)
def build_breakout_backtest(stock_records, settings):
    """Run the breakout scanner across the selected stock records."""
    rows = []
    for record in stock_records:
        row = pd.Series({"股票代號": record[0], "股票名稱": record[1], "EPS合計": record[2]})
        rows.extend(run_breakout_backtest_for_stock(row, settings))
    return pd.DataFrame(rows)


def summarize_breakout_result(result_df):
    """Summarize breakout backtest rows while safely handling empty results and missing columns."""
    if result_df.empty:
        return {
            "總訊號數": 0,
            "獨立股票數": 0,
            "平均每檔股票觸發次數": 0,
            "5日後上漲機率": "",
            "10日後上漲機率": "",
            "20日後上漲機率": "",
            "5日平均報酬率": "",
            "10日平均報酬率": "",
            "20日平均報酬率": "",
            "20日內平均最大漲幅": "",
            "20日內平均最大回檔": "",
        }

    total_signals = len(result_df)
    unique_stocks = result_df["股票代號"].nunique() if "股票代號" in result_df.columns else 0
    avg_triggers = total_signals / unique_stocks if unique_stocks else 0

    def numeric_column(column):
        if column not in result_df.columns:
            return pd.Series(dtype="float64")
        return pd.to_numeric(result_df[column], errors="coerce")

    return_5d = numeric_column("5日後報酬率")
    return_10d = numeric_column("10日後報酬率")
    return_20d = numeric_column("20日後報酬率")
    max_gain = numeric_column("觸發後20日內最大漲幅")
    max_drawdown = numeric_column("觸發後20日內最大回檔")

    return {
        "總訊號數": total_signals,
        "獨立股票數": unique_stocks,
        "平均每檔股票觸發次數": round(avg_triggers, 2),
        "5日後上漲機率": f"{(return_5d > 0).mean() * 100:.1f}%" if not return_5d.empty else "",
        "10日後上漲機率": f"{(return_10d > 0).mean() * 100:.1f}%" if not return_10d.empty else "",
        "20日後上漲機率": f"{(return_20d > 0).mean() * 100:.1f}%" if not return_20d.empty else "",
        "5日平均報酬率": format_signed_pct(return_5d.mean()),
        "10日平均報酬率": format_signed_pct(return_10d.mean()),
        "20日平均報酬率": format_signed_pct(return_20d.mean()),
        "20日內平均最大漲幅": format_signed_pct(max_gain.mean()),
        "20日內平均最大回檔": format_signed_pct(max_drawdown.mean()),
    }


def build_strategy_comparison(stock_records, base_settings, eps_available):
    """Compare common breakout condition sets using the current stock list and cached price history."""
    variants = [
        ("A. 只突破高點", {"require_ma_alignment": False, "require_volume": False, "require_eps": False}),
        ("B. 突破高點 + 均線多頭", {"require_ma_alignment": True, "require_volume": False, "require_eps": False}),
        ("C. 突破高點 + 均線多頭 + 量增", {"require_ma_alignment": True, "require_volume": True, "require_eps": False}),
        ("D. 突破高點 + 均線多頭 + 量增 + EPS > 0", {"require_ma_alignment": True, "require_volume": True, "require_eps": True}),
    ]

    rows = []
    for strategy_name, overrides in variants:
        if overrides.get("require_eps") and not eps_available:
            rows.append({"策略名稱": strategy_name, "備註": "EPS資料不足，無法比較"})
            continue

        variant_settings = dict(base_settings)
        variant_settings.update(overrides)
        variant_settings["enable_news"] = False
        variant_settings["strategy_variant"] = strategy_name
        variant_result = build_breakout_backtest(stock_records, variant_settings)
        summary = summarize_breakout_result(variant_result)
        rows.append(
            {
                "策略名稱": strategy_name,
                "總訊號數": summary["總訊號數"],
                "獨立股票數": summary["獨立股票數"],
                "5日後上漲機率": summary["5日後上漲機率"],
                "10日後上漲機率": summary["10日後上漲機率"],
                "20日後上漲機率": summary["20日後上漲機率"],
                "5日平均報酬率": summary["5日平均報酬率"],
                "10日平均報酬率": summary["10日平均報酬率"],
                "20日平均報酬率": summary["20日平均報酬率"],
                "20日內平均最大漲幅": summary["20日內平均最大漲幅"],
                "20日內平均最大回檔": summary["20日內平均最大回檔"],
                "備註": "",
            }
        )

    return pd.DataFrame(rows)


def latest_market_snapshot(row, enable_news=False):
    """Build the latest observation-pool thermometer signals for one stock."""
    ticker = str(row["股票代號"])
    name = str(row["股票名稱"])
    history = get_market_temperature_history(ticker)

    if history.empty or len(history) < 61:
        return None

    latest = history.iloc[-1]
    prev = history.iloc[-2]
    prev20_high = history["收盤價"].shift(1).rolling(20).max().iloc[-1]
    prev60_high = history["收盤價"].shift(1).rolling(60).max().iloc[-1]
    change_pct = (latest["收盤價"] - prev["收盤價"]) / prev["收盤價"] * 100 if prev["收盤價"] else 0

    return {
        "股票代號": ticker,
        "股票名稱": name,
        "ma_bull": bool(latest["MA5"] > latest["MA10"] > latest["MA20"]),
        "high20": bool(latest["收盤價"] > prev20_high),
        "high60": bool(latest["收盤價"] > prev60_high),
        "volume_surge": bool(latest["成交量"] > latest["20日均量"] * 1.5),
        "limit_up": bool(change_pct >= 9.5),
        "limit_down": bool(change_pct <= -9.5),
        "under30_turning": bool(latest["收盤價"] <= 30 and latest["MA5"] > latest["MA10"] and latest["收盤價"] > latest["MA20"]),
        **analyze_stock_news(ticker, name, enable_news=enable_news),
    }


def market_temperature_state(score):
    """Map a 0-100 observation-pool temperature score to a market-state label."""
    if score <= 20:
        return "極冷"
    if score <= 40:
        return "偏冷"
    if score <= 60:
        return "正常"
    if score <= 80:
        return "偏熱"
    return "過熱"


def attack_temperature_advice(score):
    """Return the action label and practical suggestions for an attack-temperature score."""
    if score < 20:
        return "❄ 極冷", ["降低持股比例", "以觀察為主", "等待轉強訊號"]
    if score < 40:
        return "🌥 偏冷", ["可少量試單", "避免追價", "優先觀察強勢族群"]
    if score < 60:
        return "🌤 中性", ["可少量布局", "分批進場", "關注分數歷史轉強個股"]
    if score < 80:
        return "🔥 偏熱", ["積極尋找轉強股", "可提高持股比例", "留意熱門族群"]
    return "🚀 全面攻擊", ["市場資金活躍", "可提高持股比例", "優先布局強勢股", "注意過熱風險"]


def sync_level(correlation):
    """Translate correlation into a readable market-synchronization label."""
    if pd.isna(correlation):
        return "資料不足"
    if correlation >= 0.9:
        return "高度同步"
    if correlation >= 0.7:
        return "中度同步"
    return "低同步"


def trend_structure_state(score):
    """Map the moving-average breadth score to a trend-structure label."""
    if score >= 70:
        return "偏多"
    if score >= 50:
        return "轉強"
    if score >= 30:
        return "中性"
    return "偏弱"


def market_temperature_summary(trend_state, attack_state):
    """Explain the two-layer observation-pool reading in plain language."""
    if trend_state == "偏多" and attack_state in ["極冷", "偏冷"]:
        return "結構偏多，攻擊尚未全面升溫"
    if trend_state in ["偏多", "轉強"] and attack_state in ["偏熱", "過熱"]:
        return "結構轉強，攻擊溫度偏高"
    if trend_state in ["偏弱", "中性"] and attack_state in ["偏熱", "過熱"]:
        return "短線攻擊偏熱，但結構尚未全面轉強"
    if trend_state == "偏弱" and attack_state in ["極冷", "偏冷"]:
        return "結構偏弱，攻擊也偏冷"
    return "結構與攻擊溫度大致同步"


def group_heat_level(strong_ratio):
    """Translate group strong-ratio percentage into a fire-level label."""
    if strong_ratio >= 80:
        return "🔥🔥🔥🔥🔥"
    if strong_ratio >= 60:
        return "🔥🔥🔥🔥"
    if strong_ratio >= 40:
        return "🔥🔥🔥"
    if strong_ratio >= 20:
        return "🔥🔥"
    return "🔥"


def temperature_ratio_vector(snapshot_df):
    """Build a comparable ratio vector from one temperature snapshot table."""
    if snapshot_df.empty:
        return pd.Series(dtype=float)
    total = len(snapshot_df)
    if total == 0:
        return pd.Series(dtype=float)
    return pd.Series(
        {
            "多頭排列": snapshot_df["ma_bull"].sum() / total,
            "創20日高": snapshot_df["high20"].sum() / total,
            "創60日高": snapshot_df["high60"].sum() / total,
            "量能放大": snapshot_df["volume_surge"].sum() / total,
            "漲停": snapshot_df["limit_up"].sum() / total,
            "跌停": snapshot_df["limit_down"].sum() / total,
            "低價轉強": snapshot_df["under30_turning"].sum() / total,
        },
        dtype=float,
    )


def calculate_pool_sync(observation_snapshot_df, market_snapshot_df):
    """Compare observation-pool and market-pool structure using their temperature factor ratios."""
    obs_vector = temperature_ratio_vector(observation_snapshot_df)
    market_vector = temperature_ratio_vector(market_snapshot_df)
    if obs_vector.empty or market_vector.empty:
        return pd.NA
    comparison_df = pd.concat([obs_vector, market_vector], axis=1).dropna()
    comparison_df.columns = ["觀察池", "市場池"]
    if len(comparison_df) < 2:
        return pd.NA
    if comparison_df["觀察池"].nunique() <= 1 or comparison_df["市場池"].nunique() <= 1:
        return pd.NA
    return float(comparison_df["觀察池"].corr(comparison_df["市場池"]))


@st.cache_data(ttl=900)
def build_market_temperature(stock_records, enable_news=False, save_group_history=True):
    """Calculate observation-pool temperature and group rankings from lightweight latest K data."""
    snapshots = []
    for record in stock_records:
        snapshot = latest_market_snapshot(
            pd.Series({"股票代號": record[0], "股票名稱": record[1]}),
            enable_news=enable_news,
        )
        if snapshot:
            snapshots.append(snapshot)

    snapshot_df = pd.DataFrame(snapshots)
    if snapshot_df.empty:
        stats = {
            "統計股票數": 0,
            "5MA > 10MA > 20MA": 0,
            "收盤創20日高": 0,
            "收盤創60日高": 0,
            "成交量大於20日均量1.5倍": 0,
            "漲停家數": 0,
            "跌停家數": 0,
            "股價30元以下且轉強": 0,
            "觀察池溫度分數": 0,
            "觀察池狀態": market_temperature_state(0),
            "趨勢結構分數": 0,
            "趨勢結構狀態": trend_structure_state(0),
            "攻擊溫度分數": 0,
            "攻擊溫度狀態": market_temperature_state(0),
            "綜合判斷": "目前沒有可統計股票",
        }
        return stats, pd.DataFrame(), pd.DataFrame()

    total = len(snapshot_df)
    stats = {
        "統計股票數": total,
        "5MA > 10MA > 20MA": int(snapshot_df["ma_bull"].sum()),
        "收盤創20日高": int(snapshot_df["high20"].sum()),
        "收盤創60日高": int(snapshot_df["high60"].sum()),
        "成交量大於20日均量1.5倍": int(snapshot_df["volume_surge"].sum()),
        "漲停家數": int(snapshot_df["limit_up"].sum()),
        "跌停家數": int(snapshot_df["limit_down"].sum()),
        "股價30元以下且轉強": int(snapshot_df["under30_turning"].sum()),
    }

    bull_ratio = stats["5MA > 10MA > 20MA"] / total
    high20_ratio = stats["收盤創20日高"] / total
    high60_ratio = stats["收盤創60日高"] / total
    volume_ratio = stats["成交量大於20日均量1.5倍"] / total
    limit_up_ratio = stats["漲停家數"] / total
    limit_down_ratio = stats["跌停家數"] / total
    low_price_strong_ratio = stats["股價30元以下且轉強"] / total

    # 趨勢結構公式：只看觀察池中多頭排列的比例，讓「結構是否偏多」獨立呈現。
    trend_score = bull_ratio * 100

    # 攻擊溫度公式：
    # 各訊號先除以 total 轉成比例，再乘上百分制權重；最後限制在 0～100。
    # 注意：權重本身已是百分制，所以這裡不再額外 * 100。
    attack_score = (
        bull_ratio * 25
        + high20_ratio * 20
        + high60_ratio * 20
        + volume_ratio * 15
        + limit_up_ratio * 15
        + low_price_strong_ratio * 5
        - limit_down_ratio * 20
    )
    stats["趨勢結構分數"] = round(max(0, min(100, trend_score)), 1)
    stats["趨勢結構狀態"] = trend_structure_state(stats["趨勢結構分數"])
    stats["攻擊溫度分數"] = round(max(0, min(100, attack_score)), 1)
    stats["攻擊溫度狀態"] = market_temperature_state(stats["攻擊溫度分數"])
    stats["觀察池溫度分數"] = stats["攻擊溫度分數"]
    stats["觀察池狀態"] = stats["攻擊溫度狀態"]
    stats["綜合判斷"] = market_temperature_summary(stats["趨勢結構狀態"], stats["攻擊溫度狀態"])

    group_rank = pd.DataFrame()
    group_df = load_group_data()
    if not group_df.empty:
        snapshot_df["股票代號_key"] = snapshot_df["股票代號"].map(stock_code_key)
        merged = group_df.merge(snapshot_df, on="股票代號_key", how="inner")
        if not merged.empty:
            group_rows = []
            for group_name, group_rows_df in merged.groupby("族群"):
                ma_count = int(group_rows_df["ma_bull"].sum())
                high20_count = int(group_rows_df["high20"].sum())
                volume_count = int(group_rows_df["volume_surge"].sum())
                stock_count = len(group_rows_df)
                if stock_count < 3:
                    continue

                # 強勢族群公式：多頭排列數 + 創20日高數 + 量增數。
                # 排行先用強勢比例衡量族群內部強度，再乘上 log(股票數+1) 避免 1 檔族群 100% 失真。
                strong_count = ma_count + high20_count + volume_count
                strong_ratio = strong_count / (stock_count * 3) * 100 if stock_count else 0
                weighted_raw_score = strong_ratio * math.log(stock_count + 1)
                group_rows.append(
                    {
                        "族群": group_name,
                        "股票數": stock_count,
                        "多頭排列數": ma_count,
                        "創20日高數": high20_count,
                        "量增數": volume_count,
                        "強勢分數": strong_count,
                        "強勢比例": strong_ratio,
                        "加權原始分數": weighted_raw_score,
                        "熱度等級": group_heat_level(strong_ratio),
                        "檔數": stock_count,
                        "偏多檔數": ma_count,
                        "強勢檔數": strong_count,
                    }
                )
            group_rank = pd.DataFrame(group_rows)
            if not group_rank.empty:
                max_raw_score = float(group_rank["加權原始分數"].max())
                group_rank["熱度分數"] = (
                    group_rank["加權原始分數"] / max_raw_score * 100 if max_raw_score > 0 else 0
                )
                group_rank = add_group_heat_trend(group_rank)
                if save_group_history:
                    save_group_heat_history(group_rank)
                group_rank = group_rank.sort_values(
                    ["熱度分數", "強勢比例", "股票數"],
                    ascending=[False, False, False],
                )

    return stats, snapshot_df, group_rank


@st.cache_data(ttl=900)
def build_strategy_rank(stock_records, month_count, holding_days):
    """Build the older strategy ranking table for all selected stocks."""
    rows = []

    for ticker, name in stock_records:
        ticker = str(ticker)
        name = str(name)
        history = get_official_daily_history(ticker, month_count=month_count)
        trades = run_ma_backtest(history, holding_days=holding_days, volume_multiplier=1.5)

        if trades.empty:
            continue

        summary = summarize_backtest_trades(trades)

        rows.append(
            {
                "股票": name,
                "代號": ticker,
                "交易數": summary["交易次數"],
                "勝率": summary["勝率"],
                "平均報酬": summary["平均報酬"],
                "最大回撤": summary["最大回撤"],
                "信賴度": summary["信賴度"],
            }
        )

    if not rows:
        return pd.DataFrame()

    rank_df = pd.DataFrame(rows)
    return rank_df.sort_values(["平均報酬", "勝率", "交易數"], ascending=[False, False, False])


def render_strategy_rank(df):
    """Render the older strategy leaderboard page."""
    st.subheader("🏆 策略排行榜")

    if df.empty:
        st.info("目前沒有股票資料可排行。")
        return

    col1, col2 = st.columns(2)
    with col1:
        period_label = st.selectbox("回測期間", ["6個月", "1年", "3年"], index=1, key="strategy_rank_period")
    with col2:
        holding_days = st.selectbox("持有天數", [3, 5, 10], index=1, key="strategy_rank_holding_days")

    month_count = backtest_period_months(period_label)
    stock_records = tuple(
        (str(row["股票代號"]), str(row["股票名稱"]))
        for _, row in df[["股票代號", "股票名稱"]].drop_duplicates(subset=["股票代號"]).iterrows()
    )

    with st.spinner("正在計算策略排行榜..."):
        rank_df = build_strategy_rank(stock_records, month_count, holding_days)

    if rank_df.empty:
        st.info("目前沒有股票符合這組策略條件，策略排行榜暫無資料。")
        return

    rank_display = rank_df.head(20).copy()
    rank_display["勝率"] = rank_display["勝率"].map(lambda value: f"{value:.1f}%")
    rank_display["平均報酬"] = rank_display["平均報酬"].map(format_signed_pct)
    rank_display["最大回撤"] = rank_display["最大回撤"].map(format_signed_pct)
    st.dataframe(
        rank_display[["股票", "代號", "交易數", "勝率", "平均報酬", "最大回撤", "信賴度"]],
        use_container_width=True,
        hide_index=True,
    )


def render_backtest_metric_grid(metrics):
    """Render compact metric cards used by backtest and observation-pool pages."""
    cards = "\n".join(
        f"""
        <div class="backtest-metric-card">
            <div class="backtest-metric-label">{label}</div>
            <div class="backtest-metric-value">{value}</div>
        </div>
        """
        for label, value in metrics
    )
    st.markdown(
        f"""
        <style>
        .backtest-metric-grid {{
            display: grid;
            grid-template-columns: repeat(3, minmax(0, 1fr));
            gap: 10px;
            margin: 12px 0 8px;
        }}
        .backtest-metric-card {{
            min-width: 0;
            padding: 12px 14px;
            border: 1px solid #2f3542;
            border-radius: 8px;
            background: #111827;
        }}
        .backtest-metric-label {{
            color: #aeb4c0;
            font-size: 13px;
            font-weight: 700;
            margin-bottom: 6px;
            white-space: nowrap;
        }}
        .backtest-metric-value {{
            color: #f9fafb;
            font-size: clamp(22px, 4vw, 32px);
            font-weight: 800;
            line-height: 1.15;
            overflow-wrap: anywhere;
        }}
        @media (max-width: 640px) {{
            .backtest-metric-grid {{
                grid-template-columns: repeat(2, minmax(0, 1fr));
            }}
            .backtest-metric-card {{
                padding: 10px 12px;
            }}
            .backtest-metric-value {{
                font-size: 22px;
            }}
        }}
        @media (max-width: 360px) {{
            .backtest-metric-grid {{
                grid-template-columns: 1fr;
            }}
        }}
        </style>
        <div class="backtest-metric-grid">
            {cards}
        </div>
        """,
        unsafe_allow_html=True,
    )


def ensure_backtest_record_dir():
    """Ensure the local backtest_records directory exists before writing or reading CSV files."""
    BACKTEST_RECORD_DIR.mkdir(parents=True, exist_ok=True)
    return BACKTEST_RECORD_DIR


def build_backtest_record_df(result_df, settings, strategy_name, run_time):
    """Convert the current backtest result into the normalized CSV record schema."""
    record_df = pd.DataFrame(
        {
            "run_time": run_time,
            "strategy_name": strategy_name,
            "stock_id": result_df["股票代號"],
            "stock_name": result_df["股票名稱"],
            "trigger_date": pd.to_datetime(result_df["觸發日期"]).dt.strftime("%Y-%m-%d"),
            "trigger_close": result_df["觸發收盤價"],
            "breakout_days": settings["breakout_days"],
            "require_ma_alignment": settings["require_ma_alignment"],
            "volume_ratio_threshold": settings["volume_multiplier"],
            "price_min": settings["price_min"],
            "price_max": settings["price_max"],
            "require_eps_positive": settings["require_eps"],
            "require_above_120ma": settings.get("require_above_ma120", False),
            "cooldown_days": settings.get("cooldown_days", 20),
            "strategy_variant": settings.get("strategy_variant", strategy_name),
            "return_5d": result_df["5日後報酬率"],
            "return_10d": result_df["10日後報酬率"],
            "return_20d": result_df["20日後報酬率"],
            "days_to_break_5ma": result_df["幾天後跌破5日線"],
            "max_gain_20d": result_df["觸發後20日內最大漲幅"],
            "max_drawdown_20d": result_df["觸發後20日內最大回檔"],
        }
    )
    return record_df


def save_backtest_record(result_df, settings, strategy_name):
    """Persist a non-empty backtest result to CSV and avoid duplicate writes during Streamlit reruns."""
    if result_df.empty:
        return None, None

    run_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    record_signature = (
        strategy_name,
        tuple(sorted(settings.items())),
        len(result_df),
        str(result_df["觸發日期"].max()),
    )

    if st.session_state.get("last_backtest_record_signature") == record_signature:
        return st.session_state.get("last_backtest_record_path"), st.session_state.get("last_backtest_record_csv")

    record_df = build_backtest_record_df(result_df, settings, strategy_name, run_time)
    record_dir = ensure_backtest_record_dir()
    file_time = datetime.now().strftime("%Y%m%d_%H%M%S")
    record_path = record_dir / f"backtest_{file_time}.csv"
    csv_text = record_df.to_csv(index=False, encoding="utf-8-sig")
    record_path.write_text(csv_text, encoding="utf-8-sig")

    st.session_state["last_backtest_record_signature"] = record_signature
    st.session_state["last_backtest_record_path"] = str(record_path)
    st.session_state["last_backtest_record_csv"] = csv_text
    return str(record_path), csv_text


def list_backtest_record_files():
    """List saved backtest CSV files from newest to oldest."""
    record_dir = ensure_backtest_record_dir()
    return sorted(record_dir.glob("backtest_*.csv"), key=lambda path: path.stat().st_mtime, reverse=True)


def summarize_backtest_record(record_df):
    """Summarize a saved backtest CSV for the historical record viewer."""
    if record_df.empty:
        return []

    def numeric_record_column(column):
        if column not in record_df.columns:
            return pd.Series(dtype="float64")
        return pd.to_numeric(record_df[column], errors="coerce")

    return_5d = numeric_record_column("return_5d")
    return_10d = numeric_record_column("return_10d")
    return_20d = numeric_record_column("return_20d")
    max_gain_20d = numeric_record_column("max_gain_20d")

    return [
        ("符合筆數", len(record_df)),
        ("5日平均報酬率", format_signed_pct(return_5d.mean())),
        ("10日平均報酬率", format_signed_pct(return_10d.mean())),
        ("20日平均報酬率", format_signed_pct(return_20d.mean())),
        ("20日後上漲機率", f"{(return_20d > 0).mean() * 100:.1f}%"),
        ("20日平均最大漲幅", format_signed_pct(max_gain_20d.mean())),
    ]


def render_backtest_record_history():
    """Render the historical backtest record browser and CSV preview area."""
    st.markdown("### 歷史紀錄區")
    try:
        record_files = list_backtest_record_files()
    except Exception as exc:
        st.error(f"讀取歷史紀錄失敗：{exc}")
        return

    if not record_files:
        st.info("目前尚無歷史回測紀錄。")
        return

    recent_files = record_files[:10]
    st.dataframe(
        pd.DataFrame(
            {
                "檔案名稱": [path.name for path in recent_files],
                "修改時間": [
                    datetime.fromtimestamp(path.stat().st_mtime).strftime("%Y-%m-%d %H:%M:%S")
                    for path in recent_files
                ],
            }
        ),
        use_container_width=True,
        hide_index=True,
    )

    selected_name = st.selectbox("讀取歷史紀錄", [path.name for path in recent_files])
    selected_path = next(path for path in recent_files if path.name == selected_name)

    try:
        selected_df = pd.read_csv(selected_path)
    except Exception as exc:
        st.error(f"CSV 讀取失敗：{exc}")
        return

    render_backtest_metric_grid(summarize_backtest_record(selected_df))
    with st.expander("查看歷史紀錄明細"):
        st.dataframe(selected_df.head(200), use_container_width=True, hide_index=True)


def render_backtest_lab(df):
    """Render the breakout backtest lab: controls, summary, top runners, CSV record, and detail table."""
    st.subheader("🧪 回測實驗室")

    if df.empty:
        st.info("目前沒有股票資料可回測。")
        return

    st.caption("依條件掃描目前清單股票的歷史觸發點，並統計觸發後 5 / 10 / 20 日表現。")

    turnaround_mode = st.checkbox("轉機股模式", value=False)

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        period_label = st.selectbox("回測期間", ["1年", "2年", "3年", "5年"], index=0)
        breakout_days = st.selectbox("突破幾日高點", [5, 20, 60, 120], index=1, format_func=lambda value: f"{value}日")
        breakout_method = st.selectbox("突破判斷方式", ["收盤創高", "盤中創高"], index=0)
    with col2:
        require_ma_alignment = st.checkbox("要求 5MA > 10MA > 20MA", value=True)
        require_volume = st.checkbox("要求成交量放大", value=True)
        require_above_ma120 = st.checkbox("要求站上半年線", value=False)
    with col3:
        volume_multiplier = st.number_input("成交量 / 20日均量倍數", min_value=1.0, max_value=5.0, value=1.5, step=0.1)
        require_eps = st.checkbox("最近四季 EPS 合計 > 0", value=False)
        enable_news = st.checkbox("啟用新聞熱度", value=False, key="backtest_enable_news")
    with col4:
        max_stock_label = st.selectbox("最多回測股票數", ["10", "20", "30", "全部"], index=1)
        cooldown_days = st.selectbox("同股票訊號冷卻期", [0, 5, 10, 20, 30], index=3, format_func=lambda value: f"{value}日")

    price_col1, price_col2 = st.columns(2)
    with price_col1:
        price_min = st.number_input("股價下限", min_value=0.0, value=5.0, step=1.0)
    with price_col2:
        price_max = st.number_input("股價上限", min_value=1.0, value=30.0, step=10.0)

    if turnaround_mode:
        # 轉機股模式：把常用低價轉強條件一次套用，降低手動設定成本。
        breakout_days = 20
        breakout_method = "收盤創高"
        require_ma_alignment = True
        require_volume = True
        require_eps = True
        require_above_ma120 = True
        volume_multiplier = 1.5
        price_min = 5.0
        price_max = 30.0
        st.info("轉機股模式已套用：股價 5～30、EPS > 0、多頭排列、量能 1.5 倍、站上半年線、收盤創 20 日高。")

    all_eps_columns = get_eps_columns(tuple(df.columns))
    if require_eps and not all_eps_columns:
        st.warning("目前資料表沒有 EPS 欄位，EPS 條件會先略過；之後若補進 EPS 資料即可啟用。")
        require_eps = False

    settings = {
        "month_count": backtest_period_months(period_label),
        "breakout_days": breakout_days,
        "breakout_method": breakout_method,
        "require_ma_alignment": require_ma_alignment,
        "require_volume": require_volume,
        "require_above_ma120": require_above_ma120,
        "volume_multiplier": volume_multiplier,
        "price_min": price_min,
        "price_max": price_max,
        "require_eps": require_eps,
        "enable_news": enable_news,
        "cooldown_days": cooldown_days,
        "strategy_variant": "自訂策略",
    }
    eps_columns = all_eps_columns
    unique_stocks = df[["股票代號", "股票名稱", *eps_columns]].drop_duplicates(subset=["股票代號"])
    if max_stock_label != "全部":
        unique_stocks = unique_stocks.head(int(max_stock_label))

    stock_records = tuple(
        (
            str(row["股票代號"]),
            str(row["股票名稱"]),
            get_eps_value(row[["股票代號", "股票名稱", *eps_columns]]) if eps_columns else None,
        )
        for _, row in unique_stocks.iterrows()
    )

    with st.spinner("正在掃描歷史觸發點..."):
        result_df = build_breakout_backtest(stock_records, settings)

    strategy_name = f"{'轉機股模式' if turnaround_mode else '突破回測'}-{breakout_method}-{breakout_days}日"

    st.markdown("### 回測條件")
    st.markdown(
        f"""
        - 突破條件：{breakout_method}，突破前 {breakout_days} 日高點。
        - 本次回測股票數：{len(stock_records)} 檔（預設限制 20 檔，避免網站過慢）。
        - 多頭排列：{"需要" if require_ma_alignment else "不要求"}。
        - 成交量條件：{"成交量 > 20日均量 " + f"{volume_multiplier:.1f}" + " 倍" if require_volume else "不要求"}。
        - 半年線條件：{"需要收盤價站上 MA120" if require_above_ma120 else "不要求"}。
        - 股價區間：{price_min:.2f} ~ {price_max:.2f}。
        - EPS 條件：{"最近四季 EPS 合計 > 0" if require_eps else "不要求或目前無 EPS 欄位"}。
        - 同股票訊號冷卻期：{cooldown_days} 個交易日。
        """
    )

    if result_df.empty:
        st.info("這組條件目前沒有找到可完成 20 日追蹤的歷史觸發紀錄。")
        st.info("本次沒有符合條件的結果，未建立紀錄檔")
    else:
        try:
            record_path, csv_text = save_backtest_record(result_df, settings, strategy_name)
            if record_path and csv_text:
                st.success(f"本次回測紀錄已儲存：{record_path}")
                st.download_button(
                    "下載本次 CSV",
                    data=csv_text.encode("utf-8-sig"),
                    file_name=Path(record_path).name,
                    mime="text/csv",
                )
        except Exception as exc:
            st.error(f"回測紀錄儲存失敗：{exc}")

    st.markdown("### 統計摘要")
    summary = summarize_breakout_result(result_df)
    render_backtest_metric_grid(
        [
            ("總訊號數", summary["總訊號數"]),
            ("獨立股票數", summary["獨立股票數"]),
            ("平均每檔觸發", summary["平均每檔股票觸發次數"]),
            ("5日後上漲機率", summary["5日後上漲機率"]),
            ("10日後上漲機率", summary["10日後上漲機率"]),
            ("20日後上漲機率", summary["20日後上漲機率"]),
            ("5日平均報酬率", summary["5日平均報酬率"]),
            ("10日平均報酬率", summary["10日平均報酬率"]),
            ("20日平均報酬率", summary["20日平均報酬率"]),
            ("20日平均最大漲幅", summary["20日內平均最大漲幅"]),
            ("20日平均最大回檔", summary["20日內平均最大回檔"]),
        ]
    )

    st.markdown("### 條件組合績效比較")
    with st.spinner("正在比較條件組合績效..."):
        comparison_df = build_strategy_comparison(stock_records, settings, eps_available=bool(all_eps_columns))
    st.dataframe(comparison_df, use_container_width=True, hide_index=True)

    st.markdown("### 最大飆股案例")
    top_runner_mode = st.selectbox(
        "最大飆股案例顯示方式",
        ["每檔股票只顯示最大漲幅最高的一筆", "顯示全部訊號"],
        index=0,
    )
    if result_df.empty:
        st.info("目前沒有最大飆股案例可顯示。")
        render_backtest_record_history()
        return

    top_runner_source = result_df.copy()
    if top_runner_mode == "每檔股票只顯示最大漲幅最高的一筆":
        top_runner_source = top_runner_source.sort_values("觸發後20日內最大漲幅", ascending=False).drop_duplicates(
            subset=["股票代號"],
            keep="first",
        )
    top_runners = top_runner_source.sort_values("觸發後20日內最大漲幅", ascending=False).head(10).copy()
    top_runners["觸發日期"] = pd.to_datetime(top_runners["觸發日期"]).dt.strftime("%Y-%m-%d")
    top_runners["觸發收盤價"] = top_runners["觸發收盤價"].map(lambda value: format_number(value, 2))
    top_runners["20日內最高價"] = top_runners["20日內最高價"].map(lambda value: format_number(value, 2))
    top_runners["20日內最大漲幅"] = top_runners["觸發後20日內最大漲幅"].map(lambda value: "" if pd.isna(value) else format_signed_pct(value))
    st.dataframe(
        top_runners[["股票代號", "股票名稱", "觸發日期", "觸發收盤價", "20日內最高價", "20日內最大漲幅"]],
        use_container_width=True,
        hide_index=True,
    )

    display_df = result_df.sort_values("觸發日期", ascending=False).copy()
    display_df["觸發日期"] = pd.to_datetime(display_df["觸發日期"]).dt.strftime("%Y-%m-%d")
    display_df["觸發收盤價"] = display_df["觸發收盤價"].map(lambda value: format_number(value, 2))
    display_df["20日內最高價"] = display_df["20日內最高價"].map(lambda value: "" if pd.isna(value) else format_number(value, 2))
    for col in ["5日後報酬率", "10日後報酬率", "20日後報酬率", "觸發後20日內最大漲幅", "觸發後20日內最大回檔"]:
        display_df[col] = display_df[col].map(lambda value: "" if pd.isna(value) else format_signed_pct(value))
    display_df["news_score"] = display_df["news_score"].map(lambda value: format_number(value, 0))

    with st.expander("查看回測明細（最多顯示前 200 筆）"):
        st.dataframe(display_df.head(200), use_container_width=True, hide_index=True)

    st.caption("報酬率未扣除手續費、交易稅、滑價與股利；歷史觸發不代表未來保證重演。")
    render_backtest_record_history()


def render_market_temperature(
    df,
    title="🌡️ Deep Trend 觀察池溫度",
    source_label="觀察池",
    scope_note="此分數只代表目前 Deep Trend 觀察清單，不代表全市場。",
):
    """Render the Deep Trend observation-pool temperature page and strong-group ranking."""
    st.subheader(title)

    if df.empty:
        st.info("目前沒有股票資料可統計。")
        return

    enable_news = st.checkbox("啟用新聞熱度", value=False, key="temperature_enable_news")
    stock_records = tuple(
        (str(row["股票代號"]), str(row["股票名稱"]))
        for _, row in df[["股票代號", "股票名稱"]].drop_duplicates(subset=["股票代號"]).iterrows()
    )

    with st.spinner("正在統計觀察池溫度..."):
        stats, snapshot_df, group_rank = build_market_temperature(stock_records, enable_news=enable_news)

    trend_col, attack_col = st.columns(2)
    with trend_col:
        st.metric("趨勢結構分數", f"{stats['趨勢結構分數']:.1f} / 100", stats["趨勢結構狀態"])
    with attack_col:
        st.metric("攻擊溫度分數", f"{stats['攻擊溫度分數']:.1f} / 100", stats["攻擊溫度狀態"])
    st.caption(
        f"{source_label}判斷：{stats['綜合判斷']}。趨勢結構看多頭排列比例；攻擊溫度看創高、量增、漲跌停與低價轉強。{scope_note}"
    )

    advice_label, advice_items = attack_temperature_advice(stats["攻擊溫度分數"])
    with st.container(border=True):
        st.markdown(f"### 操作建議｜{advice_label}")
        for item in advice_items:
            st.markdown(f"• {item}")

    metric_items = [
        ("統計股票數", stats["統計股票數"]),
        ("多頭排列股票數", stats["5MA > 10MA > 20MA"]),
        ("創20日高股票數", stats["收盤創20日高"]),
        ("創60日高股票數", stats["收盤創60日高"]),
        ("量增股票數", stats["成交量大於20日均量1.5倍"]),
        ("漲停家數", stats["漲停家數"]),
        ("跌停家數", stats["跌停家數"]),
        ("低價轉強股數", stats["股價30元以下且轉強"]),
    ]
    render_backtest_metric_grid(metric_items)

    if stats["統計股票數"] == 0:
        st.info("目前沒有可統計股票。")
        return

    if source_label == "觀察池":
        universe_df = load_universe_result()
        if universe_df.empty:
            st.info("市場同步性：尚未產生市場池資料，無法比較觀察池與市場池。")
        else:
            universe_records = tuple(
                (str(row["股票代號"]), str(row["股票名稱"]))
                for _, row in universe_df[["股票代號", "股票名稱"]].drop_duplicates(subset=["股票代號"]).iterrows()
            )
            market_stats, market_snapshot_df, _ = build_market_temperature(
                universe_records,
                enable_news=False,
                save_group_history=False,
            )
            corr = calculate_pool_sync(snapshot_df, market_snapshot_df)
            st.markdown("### 市場同步性")
            sync_cols = st.columns(3)
            sync_cols[0].metric("相關係數", "N/A" if pd.isna(corr) else f"{corr:.2f}")
            sync_cols[1].metric("同步程度", sync_level(corr))
            sync_cols[2].metric("市場池攻擊溫度", f"{market_stats['攻擊溫度分數']:.1f} / 100")
            st.caption("此處先比較觀察池與市場池的目前溫度因子比例；之後累積每日溫度歷史後，可升級成跨日相關係數。")

    st.markdown("### 強勢族群排行")
    if group_rank.empty:
        st.info("尚無符合條件的產業分類資料。族群至少需要 3 檔股票才列入排行。")
    else:
        st.markdown("#### 今日最強族群")
        medals = ["🥇", "🥈", "🥉"]
        top_groups = group_rank.head(3)
        for index, (_, row) in enumerate(top_groups.iterrows()):
            st.markdown(f"{medals[index]} {row['族群']}：{row['強勢比例']:.1f}%｜熱度 {row['熱度分數']:.1f}")

        display_group = group_rank.head(10).copy()
        if "加權原始分數" in display_group.columns:
            display_group = display_group.drop(columns=["加權原始分數"])
        display_group["熱度分數"] = display_group["熱度分數"].map(lambda value: f"{value:.1f}")
        display_group["強勢比例"] = display_group["強勢比例"].map(lambda value: f"{value:.1f}%")
        if "7日熱度變化" in display_group.columns:
            display_group["7日熱度變化"] = display_group["7日熱度變化"].map(
                lambda value: "" if pd.isna(value) else f"{value:+.1f}"
            )
        st.dataframe(display_group, use_container_width=True, hide_index=True)

    with st.expander("查看個股統計明細"):
        display_snapshot = snapshot_df.copy()
        rename_map = {
            "ma_bull": "5MA>10MA>20MA",
            "high20": "創20日高",
            "high60": "創60日高",
            "volume_surge": "量能放大",
            "limit_up": "漲停",
            "limit_down": "跌停",
            "under30_turning": "30元以下轉強",
        }
        display_snapshot = display_snapshot.rename(columns=rename_map)
        st.dataframe(
            display_snapshot[
                [
                    "股票代號",
                    "股票名稱",
                    *rename_map.values(),
                    "新聞熱度",
                    "news_count",
                    "news_titles",
                    "positive_keywords",
                    "risk_keywords",
                    "news_score",
                ]
            ],
            use_container_width=True,
            hide_index=True,
        )


def render_market_pool_temperature(universe_df):
    """Render the neutral 150-stock market-pool temperature without changing the core radar."""
    if universe_df.empty:
        st.subheader("🌡️ 市場池溫度")
        st.info("尚未產生 output/universe_analysis_result.xlsx。請先執行更新流程產生市場池分析。")
        return

    render_market_temperature(
        universe_df,
        title="🌡️ 市場池溫度",
        source_label="市場池",
        scope_note="此分數代表 150 檔中性市場池，不等同全市場，但比個人觀察池更接近大環境。",
    )


def render_deeptrend_candidates(universe_df):
    """Show the top DeepTrend candidates from the neutral market pool only."""
    st.subheader("🔭 DeepTrend 候選股")
    st.caption("從 150 檔市場池中找出新鮮轉強候選：重視分數上升、突破、量能與轉強訊號，不只是照 DeepTrend 分數排序。")

    if universe_df.empty:
        st.info("尚未產生 output/universe_analysis_result.xlsx。請先執行更新流程產生市場池分析。")
        return

    candidate_df = universe_df.copy()
    for column in ["DeepTrend分數", "分數變化率", "分數變化", "技術面分數", "籌碼分數", "量價分數"]:
        if column in candidate_df.columns:
            candidate_df[column] = pd.to_numeric(candidate_df[column], errors="coerce")

    def candidate_number(column):
        if column not in candidate_df.columns:
            return pd.Series(0, index=candidate_df.index)
        return pd.to_numeric(candidate_df[column], errors="coerce").fillna(0)

    def candidate_text(column):
        if column not in candidate_df.columns:
            return pd.Series("", index=candidate_df.index)
        return candidate_df[column].fillna("").astype(str)

    text_source = (
        candidate_text("狀態")
        + "｜"
        + candidate_text("綜合判斷")
        + "｜"
        + candidate_text("技術面")
        + "｜"
        + candidate_text("籌碼面")
    )
    deeptrend_score = candidate_number("DeepTrend分數")
    score_change = candidate_number("分數變化")
    score_change_rate = candidate_number("分數變化率")
    volume_price_score = candidate_number("量價分數")

    candidate_df["分數達標"] = deeptrend_score >= 40
    candidate_df["分數上升"] = score_change > 0
    candidate_df["明顯升溫"] = score_change >= 10
    candidate_df["快速升溫"] = score_change >= 20
    candidate_df["接近20日高"] = text_source.str.contains("接近20日高", na=False)
    candidate_df["突破20日高"] = text_source.str.contains("突破20日高|帶量突破20日高", na=False)
    candidate_df["成交量放大"] = text_source.str.contains("爆量|成交量放大|量能溫和放大|量能", na=False)
    candidate_df["多頭排列"] = text_source.str.contains("多頭排列", na=False)
    candidate_df["轉強訊號"] = text_source.str.contains("轉強|強勢|偏多", na=False)
    candidate_df["高分鈍化"] = (deeptrend_score >= 60) & (score_change <= 2)

    candidate_df["候選分數"] = (
        candidate_df["分數達標"].astype(int) * 20
        + candidate_df["分數上升"].astype(int) * 20
        + candidate_df["明顯升溫"].astype(int) * 15
        + candidate_df["快速升溫"].astype(int) * 20
        + candidate_df["接近20日高"].astype(int) * 10
        + candidate_df["突破20日高"].astype(int) * 20
        + candidate_df["成交量放大"].astype(int) * 15
        + candidate_df["多頭排列"].astype(int) * 10
        + candidate_df["轉強訊號"].astype(int) * 10
        + deeptrend_score.clip(lower=0) * 0.08
        + volume_price_score.clip(lower=0) * 0.08
        - candidate_df["高分鈍化"].astype(int) * 15
    )

    def candidate_reason(row):
        reasons = []
        if row.get("快速升溫"):
            reasons.append("快速升溫")
        elif row.get("明顯升溫"):
            reasons.append("分數明顯上升")
        elif row.get("分數上升"):
            reasons.append("分數上升")
        if row.get("突破20日高"):
            reasons.append("突破20日高")
        elif row.get("接近20日高"):
            reasons.append("接近20日高")
        if row.get("成交量放大"):
            reasons.append("成交量放大")
        if row.get("多頭排列"):
            reasons.append("多頭排列")
        if row.get("轉強訊號"):
            reasons.append("轉強訊號")
        if row.get("高分鈍化"):
            reasons.append("高分但升溫放緩")
        return "、".join(reasons) if reasons else "分數達標"

    candidate_df["候選理由"] = candidate_df.apply(candidate_reason, axis=1)
    candidate_df = candidate_df[candidate_df["分數達標"]].copy()
    candidate_df = candidate_df.sort_values(
        ["候選分數", "DeepTrend分數"],
        ascending=[False, False],
    ).head(30)

    display_cols = [
        "股票代號",
        "股票名稱",
        "候選分數",
        "候選理由",
        "DeepTrend分數",
        "分數變化",
        "分數變化率",
        "狀態",
        "綜合判斷",
        "技術面",
        "籌碼面",
    ]
    display_cols = [col for col in display_cols if col in candidate_df.columns]

    display_df = candidate_df[display_cols].copy()
    if "候選分數" in display_df.columns:
        display_df["候選分數"] = display_df["候選分數"].map(lambda value: format_number(value, 1))
    if "分數變化" in display_df.columns:
        display_df["分數變化"] = display_df["分數變化"].map(lambda value: "" if pd.isna(value) else f"{value:+.2f}")
    if "分數變化率" in display_df.columns:
        display_df["分數變化率"] = display_df["分數變化率"].map(
            lambda value: "" if pd.isna(value) else f"{value:+.2f}%"
        )
    st.dataframe(display_df, use_container_width=True, hide_index=True)


def render_k_chart(k_df, chart_mode="candlestick"):
    """Render candlestick, volume, and RSI subplots with category x-axis spacing."""
    has_real_ohlc = chart_mode == "candlestick" and all(col in k_df.columns for col in ["Open", "High", "Low", "Close"])
    close_series = get_series(k_df, "Close")
    volume_series = get_series(k_df, "Volume")
    x_values = pd.to_datetime(k_df.index).strftime("%Y-%m-%d")

    fig = make_subplots(
        rows=3,
        cols=1,
        shared_xaxes=True,
        vertical_spacing=0.03,
        row_heights=[0.6, 0.2, 0.2],
    )

    if has_real_ohlc:
        open_series = get_series(k_df, "Open")
        high_series = get_series(k_df, "High")
        low_series = get_series(k_df, "Low")
        fig.add_trace(
            go.Candlestick(
                x=x_values,
                open=open_series,
                high=high_series,
                low=low_series,
                close=close_series,
                name="K線",
                increasing_line_color="#ef4444",
                increasing_fillcolor="#ef4444",
                decreasing_line_color="#22c55e",
                decreasing_fillcolor="#22c55e",
            ),
            row=1,
            col=1,
        )
    else:
        fig.add_trace(
            go.Scatter(x=x_values, y=close_series, mode="lines", name="收盤價", line=dict(color="#ef4444")),
            row=1,
            col=1,
        )
        st.caption("目前使用官方資料備援，官方資料缺少真實 Open，因此改以收盤價折線圖顯示。")

    for ma_name in ["MA5", "MA10", "MA20"]:
        fig.add_trace(
            go.Scatter(x=x_values, y=k_df[ma_name], mode="lines", name=ma_name),
            row=1,
            col=1,
        )

    if has_real_ohlc:
        volume_colors = [
            "#ef4444" if close_series.iloc[i] >= open_series.iloc[i] else "#22c55e"
            for i in range(len(k_df))
        ]
    else:
        prev_close = close_series.shift(1)
        volume_colors = [
            "#ef4444" if pd.isna(prev_close.iloc[i]) or close_series.iloc[i] >= prev_close.iloc[i] else "#22c55e"
            for i in range(len(k_df))
        ]

    fig.add_trace(
        go.Bar(x=x_values, y=volume_series, name="成交量（紅漲綠跌）", marker_color=volume_colors),
        row=2,
        col=1,
    )

    fig.add_trace(
        go.Scatter(x=x_values, y=k_df["RSI"], mode="lines", name="RSI", line=dict(color="#facc15")),
        row=3,
        col=1,
    )

    fig.add_hline(y=70, line_dash="dash", line_color="red", row=3, col=1)
    fig.add_hline(y=30, line_dash="dash", line_color="green", row=3, col=1)

    fig.update_layout(
        height=700,
        xaxis_rangeslider_visible=False,
        showlegend=True,
        xaxis=dict(type="category"),
        xaxis2=dict(type="category"),
        xaxis3=dict(type="category"),
    )

    st.caption("🔴 紅量 = 收漲　🟢 綠量 = 收跌")
    st.plotly_chart(fig, use_container_width=True)


# =========================
# Streamlit main flow
# =========================
# Streamlit 會從這裡開始由上而下執行：
# 1. 設定頁面。
# 2. 讀取並整理 output Excel。
# 3. 顯示更新按鈕與篩選條件。
# 4. 依功能選單呼叫對應 render_* 頁面函式。

st.set_page_config(page_title="DeepTrend", page_icon="🔥", layout="wide")

st.title("🔥 DeepTrend")
st.caption("AI Quant Trading Radar")

df = apply_realtime_prices(prepare_stock_data(load_stock_result()))

status_options = ["全部"] + sorted(df["狀態"].dropna().unique().tolist())
min_score_value = int(df["技術分數"].min())
max_score_value = int(df["技術分數"].max())

with st.container(border=True):
    update_col, status_col, score_col, search_col = st.columns([1.1, 1.2, 1.6, 2.1])

    with update_col:
        st.caption("資料")
        if st.button("🔄 更新市場資料", use_container_width=True):
            with st.spinner("正在更新資料，請稍等..."):
                subprocess.run([sys.executable, str(BASE_DIR / "update_chip.py")], check=False)
                main_result = subprocess.run([sys.executable, str(BASE_DIR / "main.py")], check=False)
                if main_result.returncode == 0:
                    subprocess.run([sys.executable, str(BASE_DIR / "update_history.py")], check=False)
            st.cache_data.clear()
            if main_result.returncode == 0:
                st.success("更新完成！")
            else:
                st.warning("主分析結果不完整，已保留上一個完整交易日資料。")
            st.rerun()

    with status_col:
        selected_status = st.selectbox("狀態", status_options)

    with score_col:
        min_score = st.slider(
            "最低技術分數",
            min_value=min_score_value,
            max_value=max_score_value,
            value=min_score_value,
        )

    with search_col:
        keyword = st.text_input("搜尋股票名稱或代號")

filtered_df = df.copy()

if selected_status != "全部":
    filtered_df = filtered_df[filtered_df["狀態"] == selected_status]

filtered_df = filtered_df[filtered_df["技術分數"] >= min_score]

if keyword:
    filtered_df = filtered_df[
        filtered_df["股票名稱"].astype(str).str.contains(keyword, case=False, na=False)
        | filtered_df["股票代號"].astype(str).str.contains(keyword, case=False, na=False)
    ]

view_options = [
    "📊 股票雷達",
    "📈 分數歷史",
    "✅ 分數驗證",
    "🔎 個股查詢",
    "🌡️ 觀察池溫度",
    "🌡️ 市場池溫度",
    "🔭 DeepTrend 候選股",
    "🧾 籌碼查帳",
    "📋 詳細表格",
    "🩺 資料健康檢查",
]
if "active_view" not in st.session_state or st.session_state["active_view"] not in view_options:
    st.session_state["active_view"] = view_options[0]

active_view = st.radio(
    "功能選單",
    view_options,
    horizontal=True,
    label_visibility="collapsed",
    key="active_view",
)

if active_view == "📊 股票雷達":
    render_stock_radar(filtered_df)
elif active_view == "📋 詳細表格":
    render_scan_table(filtered_df)
elif active_view == "🔎 個股查詢":
    render_detail(filtered_df)
elif active_view == "🧾 籌碼查帳":
    render_chip_audit(df)
elif active_view == "📈 分數歷史":
    render_score_history(df)
elif active_view == "✅ 分數驗證":
    render_score_validation(df)
elif active_view == "🩺 資料健康檢查":
    render_data_health(df)
elif active_view == "🌡️ 觀察池溫度":
    render_market_temperature(df)
elif active_view == "🌡️ 市場池溫度":
    universe_raw_df = load_universe_result()
    universe_df = apply_realtime_prices(prepare_stock_data(universe_raw_df)) if not universe_raw_df.empty else universe_raw_df
    render_market_pool_temperature(universe_df)
else:
    universe_raw_df = load_universe_result()
    universe_df = apply_realtime_prices(prepare_stock_data(universe_raw_df)) if not universe_raw_df.empty else universe_raw_df
    render_deeptrend_candidates(universe_df)
