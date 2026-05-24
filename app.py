import re
import subprocess
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
        "收盤價": round(latest_close, 2),
        "漲跌": round(change, 2),
        "漲跌幅": round(change_pct, 2),
        "5MA": round(latest_ma5, 2),
        "60MA": round(latest_ma60, 2),
        "訊號": signal,
        "原因": reason,
    }


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
            <div style="font-size:18px;color:#aaaaaa;margin-top:20px;">收盤價</div>
            <div style="font-size:56px;font-weight:bold;color:white;margin-top:10px;">
                {market["收盤價"]:,.2f}
            </div>
            <div style="font-size:30px;font-weight:bold;color:{change_color};margin-top:10px;">
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

st.title("🔥 DeepTrend")
st.caption("AI Quant Trading Radar")

with st.container(border=True):
    st.markdown("## 📡 市場方向觀察")
    st.caption("大盤訊號僅供參考，不納入個股評分")

    markets = [
        get_market_signal("^TWII", "加權指數"),
        get_market_signal("0050.TW", "0050 ETF"),
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

df = prepare_stock_data(load_stock_result())

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
