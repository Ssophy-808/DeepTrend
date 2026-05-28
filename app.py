import re
import subprocess
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date
from pathlib import Path
from textwrap import dedent

import pandas as pd
import plotly.graph_objects as go
import requests
import streamlit as st
import urllib3
import yfinance as yf
from plotly.subplots import make_subplots


BASE_DIR = Path(__file__).resolve().parent
RESULT_FILE = BASE_DIR / "output" / "stock_analysis_result.xlsx"
AUTO_REFRESH_SECONDS = 180

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


def get_series(df, column):
    data = df[column]
    if isinstance(data, pd.DataFrame):
        return data.iloc[:, 0]
    return data


def normalize_tw_symbol(symbol):
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
    return yf.download(
        symbol,
        period=period,
        interval=interval,
        progress=False,
        auto_adjust=False,
    )


def empty_market_signal(symbol, name, reason="抓不到資料"):
    return {
        "名稱": name,
        "代號": symbol,
        "收盤價": 0,
        "漲跌": 0,
        "漲跌幅": 0,
        "5MA": 0,
        "60MA": 0,
        "訊號": "無資料",
        "原因": reason,
    }


def get_market_signal(symbol, name):
    market_df = download_market_data(symbol)

    if market_df.empty or "Close" not in market_df.columns:
        return empty_market_signal(symbol, name)

    close_series = get_series(market_df, "Close").dropna()

    if len(close_series) < 60:
        return empty_market_signal(symbol, name, "資料不足，無法計算60MA")

    ma5 = close_series.rolling(5).mean()
    ma60 = close_series.rolling(60).mean()

    latest_close = float(close_series.iloc[-1])
    prev_close = float(close_series.iloc[-2])
    latest_ma5 = float(ma5.iloc[-1])
    latest_ma60 = float(ma60.iloc[-1])
    prev_ma5 = float(ma5.iloc[-2])
    prev_ma60 = float(ma60.iloc[-2])

    change = latest_close - prev_close
    change_pct = (change / prev_close) * 100 if prev_close else 0

    signal = "無訊號"
    reason = "目前未出現明確突破或跌破"

    if latest_close > latest_ma60 and prev_ma5 <= prev_ma60 and latest_ma5 > latest_ma60:
        signal = "🟢 買進訊號"
        reason = "收盤價站上60MA，且5MA上穿60MA"
    elif latest_close < latest_ma60 and prev_ma5 >= prev_ma60 and latest_ma5 < latest_ma60:
        signal = "🔴 賣出訊號"
        reason = "收盤價跌破60MA，且5MA下穿60MA"

    return {
        "名稱": name,
        "代號": symbol,
        "價格標籤": "收盤價",
        "收盤價": round(latest_close, 2),
        "漲跌": round(change, 2),
        "漲跌幅": round(change_pct, 2),
        "5MA": round(latest_ma5, 2),
        "60MA": round(latest_ma60, 2),
        "訊號": signal,
        "原因": reason,
    }


def first_quote_price(value):
    text = str(value or "").strip()
    if not text or text == "-":
        return 0
    first = text.split("_")[0].replace(",", "").strip()
    try:
        return float(first)
    except ValueError:
        return 0


def to_float(value):
    try:
        text = str(value).replace(",", "").replace("%", "").strip()
        if text in ["", "-", "NULL"]:
            return 0
        return float(text)
    except (TypeError, ValueError):
        return 0


def parse_roc_date(value):
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


def get_twse_realtime_signal(ex_ch, symbol, name):
    url = "https://mis.twse.com.tw/stock/api/getStockInfo.jsp"
    headers = {
        "User-Agent": "Mozilla/5.0",
        "Referer": "https://mis.twse.com.tw/stock/fibest.jsp?stock=0050",
    }

    try:
        response = requests.get(
            url,
            params={"ex_ch": ex_ch, "json": "1", "delay": "0"},
            timeout=10,
            verify=False,
            headers=headers,
        )
        item = response.json()["msgArray"][0]
    except Exception:
        return empty_market_signal(symbol, name, "證交所即時資料抓取失敗")

    previous_close = to_float(item.get("y"))
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
            latest_price = to_float(item.get("o")) or previous_close

    change = latest_price - previous_close if previous_close else 0
    change_pct = (change / previous_close) * 100 if previous_close else 0

    signal = "無訊號"
    if change > 0:
        signal = "🟢 偏多"
    elif change < 0:
        signal = "🔴 偏空"

    trade_time = item.get("t") or item.get("%") or ""

    return {
        "名稱": name,
        "代號": symbol,
        "價格標籤": "收盤/即時價",
        "收盤價": round(latest_price, 2),
        "漲跌": round(change, 2),
        "漲跌幅": round(change_pct, 2),
        "5MA": 0,
        "60MA": 0,
        "訊號": signal,
        "原因": f"證交所即時資料 {trade_time}",
    }


def build_twse_channel(ticker):
    text = str(ticker).strip()
    code = text.split(".")[0]

    if not code.isdigit():
        return None

    if text.endswith(".TWO"):
        return f"otc_{code}.tw"

    return f"tse_{code}.tw"


@st.cache_data(ttl=300)
def get_official_daily_history(ticker):
    text = str(ticker).strip()
    code = text.split(".")[0]

    if not code:
        return pd.DataFrame(columns=["日期", "收盤價", "最高價", "最低價", "成交量"])

    rows = []
    headers = {"User-Agent": "Mozilla/5.0"}

    for month_start in recent_month_starts(3):
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
    return history.tail(60).reset_index(drop=True)


def official_history_to_kline(history):
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

    k_df["Open"] = k_df["Close"].shift(1)
    k_df["Open"] = k_df["Open"].fillna(k_df["Close"])
    k_df = k_df.dropna(subset=["Open", "High", "Low", "Close"])

    return k_df[["Open", "High", "Low", "Close", "Volume"]]


@st.cache_data(ttl=10)
def get_twse_realtime_map(tickers):
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


def get_txff_signal(name="台指近月"):
    page_url = "https://www.cmoney.tw/finance/futuresnearbytxf.aspx?key=TXF1PM"
    api_url = "https://www.cmoney.tw/finance/ashx/FuturesData.ashx"
    headers = {
        "User-Agent": "Mozilla/5.0",
        "Referer": page_url,
    }

    try:
        page = requests.get(page_url, timeout=10, verify=False, headers=headers).text
        cmkeys = []
        for match in re.finditer(r"<a [^>]*futuresnearbytxf\.aspx\?key=[^']+'[^>]*>", page):
            anchor = match.group(0)
            key_match = re.search(r"cmkey='([^']+)'", anchor)
            if key_match and key_match.group(1) not in cmkeys:
                cmkeys.append(key_match.group(1))

        if not cmkeys:
            cmkeys = re.findall(r"cmkey='([^']+)'", page)

        for fallback_key in ("BDeHVSgtX1n5LWX2LS3UhQ==", "KC0RXxR2JlTTibQFJiOCOg=="):
            if fallback_key not in cmkeys:
                cmkeys.append(fallback_key)

        info = None
        for cmkey in cmkeys:
            for futures_key in ("TXF1PM", "TXF"):
                data = requests.get(
                    api_url,
                    params={
                        "action": "GetNearFutureInstantData",
                        "key": futures_key,
                        "cmkey": cmkey,
                    },
                    timeout=10,
                    verify=False,
                    headers=headers,
                ).json()
                info = data.get("RealInfo")
                if info and float(info.get("SalePr") or 0) > 0:
                    break
            if info and float(info.get("SalePr") or 0) > 0:
                break

        if not info:
            return empty_market_signal("TXFF", name, "CMoney 台指近月未回傳即時資料")
    except Exception:
        return empty_market_signal("TXFF", name, "CMoney 即時資料抓取失敗")

    latest_close = float(info.get("SalePr") or 0)
    change = float(info.get("PriceDifference") or 0)
    change_pct = float(info.get("MagnitudeOfPrice") or 0)
    trade_time = str(info.get("SaleTe") or "")
    code = str(info.get("Commkey") or "TXFF")

    if latest_close == 0:
        return empty_market_signal("TXFF", name, "CMoney 台指近月回傳價格為 0")

    signal = "無訊號"
    reason = f"{code} 即時報價 {trade_time}"

    if change > 0:
        signal = "🟢 偏多"
        reason += "，上漲"
    elif change < 0:
        signal = "🔴 偏空"
        reason += "，下跌"

    return {
        "名稱": name,
        "代號": code,
        "價格標籤": "收盤/即時價",
        "收盤價": round(latest_close, 2),
        "漲跌": round(change, 2),
        "漲跌幅": round(change_pct, 2),
        "5MA": 0,
        "60MA": 0,
        "訊號": signal,
        "原因": reason,
    }


def render_market_card(market):
    change = market["漲跌"]
    change_color = "#ff4b4b" if change > 0 else "#00c853" if change < 0 else "#aaaaaa"
    arrow = "▲" if change > 0 else "▼" if change < 0 else "－"

    html = dedent(
        f"""
        <div style="
            padding:25px;
            border:1px solid #333;
            border-radius:20px;
            background-color:#0e1117;
            min-height:300px;
        ">
            <h2>{market["名稱"]}</h2>
            <div style="font-size:18px;color:#aaaaaa;margin-top:20px;">{market.get("價格標籤", "現價")}</div>
            <div style="
                font-size:42px;
                line-height:1.1;
                font-weight:bold;
                color:white;
                margin-top:10px;
                white-space:nowrap;
                overflow:hidden;
                text-overflow:ellipsis;
            ">
                {market["收盤價"]:,.2f}
            </div>
            <div style="
                font-size:24px;
                line-height:1.2;
                font-weight:bold;
                color:{change_color};
                margin-top:10px;
                white-space:nowrap;
                overflow:hidden;
                text-overflow:ellipsis;
            ">
                {arrow} {abs(market["漲跌"]):,.2f} ({market["漲跌幅"]:+.2f}%)
            </div>
            <div style="margin-top:20px;color:#cccccc;">訊號：{market["訊號"]}</div>
            <div style="color:#888888;margin-top:8px;">{market["原因"]}</div>
        </div>
        """
    ).replace("\n", "")

    st.markdown(html, unsafe_allow_html=True)


def color_status(val):
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
    if pd.isna(value):
        return ""
    try:
        return f"{float(value):,.{decimals}f}"
    except (TypeError, ValueError):
        return value


def format_integer(value):
    if pd.isna(value):
        return ""
    try:
        return f"{int(value):,}"
    except (TypeError, ValueError):
        return value


def format_signed_pct(value):
    if pd.isna(value):
        return ""
    try:
        return f"{float(value):+,.2f}%"
    except (TypeError, ValueError):
        return value


def value_color(value):
    try:
        number = float(value)
    except (TypeError, ValueError):
        return "#aaaaaa"

    if number > 0:
        return "#ff4b4b"
    if number < 0:
        return "#22c55e"
    return "#aaaaaa"


def load_stock_result():
    try:
        return pd.read_excel(RESULT_FILE)
    except FileNotFoundError:
        st.error(f"找不到分析結果檔案：{RESULT_FILE}")
        st.stop()


def prepare_stock_data(df):
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

    df["乖離率"] = ((df["收盤價"] - df["5日線"]) / df["5日線"] * 100).round(2)
    return df


def apply_realtime_prices(df):
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

    return updated_df


def render_rank(top_strength):
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


def open_stock_detail(stock_code):
    stock_code = str(stock_code)
    st.session_state["pending_detail_stock"] = stock_code
    st.session_state["detail_stock"] = stock_code
    st.session_state["active_view"] = "🔎 個股查詢"


def render_stock_radar(filtered_df):
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
        signal_color = "#facc15" if str(volume_price_signal) != "無明顯異常" else "#9ca3af"
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
                <div style="margin-top:10px;font-size:13px;font-weight:700;color:{signal_color};">{volume_price_signal}</div>
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


def render_scan_table(filtered_df):
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

    front_columns = ["狀態", "股票代號", "股票名稱", "收盤價", "今日漲跌幅", "乖離率", "量價異常", "外資5日", "投信5日"]
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


def render_detail(filtered_df):
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

    st.markdown("### 📌 技術面")
    st.info(selected_row["技術面"])

    st.markdown("### 💰 籌碼面")
    st.success(selected_row["籌碼面"])

    symbol = normalize_tw_symbol(selected_stock)
    history = get_official_daily_history(symbol)
    k_df = official_history_to_kline(history)

    if k_df.empty:
        k_df = download_market_data(symbol)

    if k_df.empty:
        st.warning(f"抓不到 {symbol} 的K線資料。")
        return

    open_series = get_series(k_df, "Open")
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

    render_k_chart(k_df)


def render_k_chart(k_df):
    open_series = get_series(k_df, "Open")
    high_series = get_series(k_df, "High")
    low_series = get_series(k_df, "Low")
    close_series = get_series(k_df, "Close")
    volume_series = get_series(k_df, "Volume")

    fig = make_subplots(
        rows=3,
        cols=1,
        shared_xaxes=True,
        vertical_spacing=0.03,
        row_heights=[0.6, 0.2, 0.2],
    )

    fig.add_trace(
        go.Candlestick(
            x=k_df.index,
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

    for ma_name in ["MA5", "MA10", "MA20"]:
        fig.add_trace(
            go.Scatter(x=k_df.index, y=k_df[ma_name], mode="lines", name=ma_name),
            row=1,
            col=1,
        )

    volume_colors = [
        "#ef4444" if close_series.iloc[i] >= open_series.iloc[i] else "#22c55e"
        for i in range(len(k_df))
    ]

    fig.add_trace(
        go.Bar(x=k_df.index, y=volume_series, name="成交量（紅漲綠跌）", marker_color=volume_colors),
        row=2,
        col=1,
    )

    fig.add_trace(
        go.Scatter(x=k_df.index, y=k_df["RSI"], mode="lines", name="RSI", line=dict(color="#facc15")),
        row=3,
        col=1,
    )

    fig.add_hline(y=70, line_dash="dash", line_color="red", row=3, col=1)
    fig.add_hline(y=30, line_dash="dash", line_color="green", row=3, col=1)

    fig.update_layout(
        height=700,
        xaxis_rangeslider_visible=False,
        showlegend=True,
        xaxis=dict(rangebreaks=[dict(bounds=["sat", "mon"])]),
    )

    st.caption("🔴 紅量 = 收漲　🟢 綠量 = 收跌")
    st.plotly_chart(fig, use_container_width=True)


st.set_page_config(page_title="DeepTrend", page_icon="🔥", layout="wide")

st.markdown(
    f"""
    <meta http-equiv="refresh" content="{AUTO_REFRESH_SECONDS}">
    """,
    unsafe_allow_html=True,
)

st.title("🔥 DeepTrend")
st.caption("AI Quant Trading Radar")

with st.container(border=True):
    st.markdown("## 📡 市場方向觀察")
    st.caption("大盤訊號僅供參考，不納入個股評分")

    markets = [
        get_twse_realtime_signal("tse_t00.tw", "t00", "加權指數"),
        get_twse_realtime_signal("tse_0050.tw", "0050", "0050 ETF"),
        get_txff_signal("台指近月"),
    ]

    for col, market in zip(st.columns(3), markets):
        with col:
            render_market_card(market)

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
                subprocess.run(["python", str(BASE_DIR / "update_chip.py")], check=False)
                subprocess.run(["python", str(BASE_DIR / "main.py")], check=False)
            st.cache_data.clear()
            st.success("更新完成！")
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
top_strength = filtered_df.sort_values(by="乖離率", ascending=False).head(5)

if selected_status != "全部":
    filtered_df = filtered_df[filtered_df["狀態"] == selected_status]

filtered_df = filtered_df[filtered_df["技術分數"] >= min_score]

if keyword:
    filtered_df = filtered_df[
        filtered_df["股票名稱"].astype(str).str.contains(keyword, case=False, na=False)
        | filtered_df["股票代號"].astype(str).str.contains(keyword, case=False, na=False)
    ]

view_options = ["📊 股票雷達", "📋 詳細表格", "🚀 強勢排行榜", "🔎 個股查詢"]
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
elif active_view == "🚀 強勢排行榜":
    render_rank(top_strength)
else:
    render_detail(filtered_df)
