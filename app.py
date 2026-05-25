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
                realtime[ticker] = {
                    "price": latest_price,
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
        match = re.search(
            r"futuresnearbytxf\.aspx\?key=TXF1PM'[^>]+cmkey='([^']+)'",
            page,
        )
        cmkey = match.group(1) if match else "oqRHZNsm42MVMSmRgj7fRQ=="

        data = requests.get(
            api_url,
            params={
                "action": "GetNearFutureInstantData",
                "key": "TXF1PM",
                "cmkey": cmkey,
            },
            timeout=10,
            verify=False,
            headers=headers,
        ).json()
        info = data["RealInfo"]
    except Exception:
        return empty_market_signal("TXFF", name, "CMoney 即時資料抓取失敗")

    latest_close = float(info.get("SalePr") or 0)
    change = float(info.get("PriceDifference") or 0)
    change_pct = float(info.get("MagnitudeOfPrice") or 0)
    trade_time = str(info.get("SaleTe") or "")

    signal = "無訊號"
    reason = f"TXF1PM 即時報價 {trade_time}"

    if change > 0:
        signal = "🟢 偏多"
        reason += "，上漲"
    elif change < 0:
        signal = "🔴 偏空"
        reason += "，下跌"

    return {
        "名稱": name,
        "代號": "TXFF",
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

    df["漲幅%"] = ((df["收盤價"] - df["5日線"]) / df["5日線"] * 100).round(2)
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

        updated_df.loc[mask, "資料時間"] = data.get("time", "")

    if "5日線" in updated_df.columns:
        updated_df["漲幅%"] = (
            (updated_df["收盤價"] - updated_df["5日線"]) / updated_df["5日線"] * 100
        ).round(2)

    return updated_df


def render_rank(top_strength):
    st.subheader("🚀 強勢股排行榜")

    if top_strength.empty:
        st.info("目前沒有可顯示的排行資料。")
        return

    for i, (_, row) in enumerate(top_strength.iterrows(), 1):
        color = "#ff4b4b" if row["漲幅%"] > 0 else "#00c853" if row["漲幅%"] < 0 else "#aaaaaa"
        html = dedent(
            f"""
            <div style="padding:12px;margin-bottom:10px;border-radius:12px;background-color:#111111;border:1px solid #333;">
                <span style="font-size:20px;font-weight:bold;color:white;">{i}. {row["股票名稱"]}</span>
                <span style="float:right;font-size:22px;font-weight:bold;color:{color};">{row["漲幅%"]:+.2f}%</span>
            </div>
            """
        ).replace("\n", "")
        st.markdown(html, unsafe_allow_html=True)


def render_scan_table(filtered_df):
    st.subheader("📊 股票掃描結果")
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

    front_columns = ["股票代號", "股票名稱", "資料時間"]
    ordered_columns = [col for col in front_columns if col in display_df.columns]
    ordered_columns += [col for col in display_df.columns if col not in ordered_columns]
    display_df = display_df[ordered_columns]

    price_columns = ["收盤價", "5日線", "10日線", "20日線", "20日高點", "20日低點", "漲幅%"]
    chip_columns = ["籌碼1日", "籌碼3日", "籌碼5日", "籌碼10日"]

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

    selected_stock = st.selectbox("選擇股票查看K線", filtered_df["股票代號"].astype(str))
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

if st.button("🔄 更新市場資料"):
    with st.spinner("正在更新資料，請稍等..."):
        subprocess.run(["python", str(BASE_DIR / "main.py")], check=False)
    st.cache_data.clear()
    st.success("更新完成！")
    st.rerun()

df = apply_realtime_prices(prepare_stock_data(load_stock_result()))

st.sidebar.header("篩選條件")

status_options = ["全部"] + sorted(df["狀態"].dropna().unique().tolist())
selected_status = st.sidebar.selectbox("狀態", status_options)

min_score = st.sidebar.slider(
    "最低技術分數",
    min_value=int(df["技術分數"].min()),
    max_value=int(df["技術分數"].max()),
    value=int(df["技術分數"].min()),
)

keyword = st.sidebar.text_input("搜尋股票名稱或代號")

filtered_df = df.copy()
top_strength = filtered_df.sort_values(by="漲幅%", ascending=False).head(5)

if selected_status != "全部":
    filtered_df = filtered_df[filtered_df["狀態"] == selected_status]

filtered_df = filtered_df[filtered_df["技術分數"] >= min_score]

if keyword:
    filtered_df = filtered_df[
        filtered_df["股票名稱"].astype(str).str.contains(keyword, case=False, na=False)
        | filtered_df["股票代號"].astype(str).str.contains(keyword, case=False, na=False)
    ]

tab_scan, tab_rank, tab_detail = st.tabs(["📊 股票掃描", "🚀 強勢排行榜", "🔎 個股查詢"])

with tab_scan:
    render_scan_table(filtered_df)

with tab_rank:
    render_rank(top_strength)

with tab_detail:
    render_detail(filtered_df)
