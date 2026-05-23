import streamlit as st
import pandas as pd
import yfinance as yf
import plotly.graph_objects as go
import subprocess

def get_series(df, column):
    data = df[column]
    if isinstance(data, pd.DataFrame):
        return data.iloc[:, 0]
    return data

def get_market_signal(symbol, name):
    market_df = yf.download(
        symbol,
        period="3mo",
        interval="1d",
        progress=False,
        auto_adjust=False
    )

    if market_df.empty:
        return {
            "名稱": name,
            "代號": symbol,
            "訊號": "無資料",
            "原因": "抓不到資料"
        }

    close_series = get_series(market_df, "Close")

    ma5 = close_series.rolling(5).mean()
    ma60 = close_series.rolling(60).mean()

    latest_close = float(close_series.iloc[-1])
    latest_ma5 = float(ma5.iloc[-1])
    latest_ma60 = float(ma60.iloc[-1])

    prev_close = float(close_series.iloc[-2])
    change = latest_close - prev_close
    change_pct = (change / prev_close) * 100
    prev_ma5 = float(ma5.iloc[-2])
    prev_ma60 = float(ma60.iloc[-2])

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
        "原因": reason
    }

st.set_page_config(
    page_title="DeepTrend",
    page_icon="🔥",
    layout="wide"
)

st.title("🔥 DeepTrend")
st.caption("AI Quant Trading Radar")

with st.container(border=True):
    st.markdown("## 📡 市場方向觀察")
    st.caption("大盤訊號僅供參考，不納入個股評分")

    market_1 = get_market_signal("^TWII", "加權指數")
    market_2 = get_market_signal("0050.TW", "0050 ETF")

    col_m1, col_m2 = st.columns(2)

    with col_m1:
        st.subheader(market_1["名稱"])
        change_color_1 = "#ff4b4b" if market_1["漲跌"] > 0 else "#00c853"
        arrow_1 = "▲" if market_1["漲跌"] > 0 else "▼"

        st.markdown(f"""
        <div style="
            padding:25px;
            border:1px solid #333;
            border-radius:20px;
            background-color:#0e1117;
        ">
            <h2>{market_1["名稱"]}</h2>

        <div style="
             font-size:18px;
            color:#aaaaaa;
            margin-top:20px;
        ">
            收盤價
        </div>

        <div style="
            font-size:64px;
            font-weight:bold;
            color:white;
            margin-top:10px;
        ">
            {market_1["收盤價"]:,.2f}
        </div>

        <div style="
            font-size:32px;
            font-weight:bold;
            color:{change_color_1};
            margin-top:10px;
        ">
            {arrow_1}
            {abs(market_1["漲跌"]):,.2f}
            ({market_1["漲跌幅"]:+.2f}%)
         </div>

        <div style="
             margin-top:20px;
            color:#cccccc;
        ">
             訊號：{market_1["訊號"]}
        </div>

        <div style="
            color:#888888;
            margin-top:8px;
        ">
            {market_1["原因"]}
        </div>

        </div>
        """, unsafe_allow_html=True)

    with col_m2:
        st.subheader(market_2["名稱"])
        change_color_2 = "#ff4b4b" if market_2["漲跌"] > 0 else "#00c853"
        arrow_2 = "▲" if market_2["漲跌"] > 0 else "▼"

        st.markdown(f"""
        <div style="
            padding:25px;
            border:1px solid #333;
            border-radius:20px;
            background-color:#0e1117;
        ">
            <h2>{market_2["名稱"]}</h2>

        <div style="
             font-size:18px;
            color:#aaaaaa;
            margin-top:20px;
        ">
            收盤價
        </div>

        <div style="
            font-size:64px;
            font-weight:bold;
            color:white;
            margin-top:10px;
        ">
            {market_2["收盤價"]:,.2f}
        </div>

        <div style="
            font-size:32px;
            font-weight:bold;
            color:{change_color_1};
            margin-top:10px;
        ">
            {arrow_2}
            {abs(market_2["漲跌"]):,.2f}
            ({market_2["漲跌幅"]:+.2f}%)
         </div>

        <div style="
             margin-top:20px;
            color:#cccccc;
        ">
             訊號：{market_2["訊號"]}
        </div>

        <div style="
            color:#888888;
            margin-top:8px;
        ">
            {market_2["原因"]}
        </div>

        </div>
        """, unsafe_allow_html=True)

if st.button("🔄 更新市場資料"):
    with st.spinner("正在更新資料，請稍等..."):
        subprocess.run(["python", "main.py"])
    st.success("更新完成！")
    st.rerun()

df = pd.read_excel("output/stock_analysis_result.xlsx")

# 側邊欄篩選
st.sidebar.header("篩選條件")

status_options = ["全部"] + sorted(df["狀態"].dropna().unique().tolist())
selected_status = st.sidebar.selectbox("狀態", status_options)

min_score = st.sidebar.slider(
    "最低技術分數",
    min_value=int(df["技術分數"].min()),
    max_value=int(df["技術分數"].max()),
    value=int(df["技術分數"].min())
)

keyword = st.sidebar.text_input("搜尋股票名稱或代號")

filtered_df = df.copy()

filtered_df["漲幅%"] = (
    (filtered_df["收盤價"] - filtered_df["5日線"])
    / filtered_df["5日線"]
) * 100

top_strength = filtered_df.sort_values(
    by="漲幅%",
    ascending=False
).head(5)

if selected_status != "全部":
    filtered_df = filtered_df[filtered_df["狀態"] == selected_status]

filtered_df = filtered_df[filtered_df["技術分數"] >= min_score]

if keyword:
    filtered_df = filtered_df[
        filtered_df["股票名稱"].astype(str).str.contains(keyword, case=False, na=False) |
        filtered_df["股票代號"].astype(str).str.contains(keyword, case=False, na=False)
    ]

tab_scan, tab_rank, tab_detail = st.tabs([
    "📊 股票掃描",
    "🚀 強勢排行榜",
    "🔎 個股查詢"
])

with tab_rank:
    st.subheader("🚀 強勢股排行榜")

    for i, (_, row) in enumerate(top_strength.iterrows(), 1):

        color = "#ff4b4b" if row["漲幅%"] > 0 else "#00c853"

        st.markdown(f"""
        <div style="padding:12px;margin-bottom:10px;border-radius:12px;background-color:#111111;border:1px solid #333;">

        <span style="font-size:20px;font-weight:bold;color:white;">
        {i}️⃣ {row["股票名稱"]}
        </span>

        <span style="float:right;font-size:22px;font-weight:bold;color:{color};">
        {row["漲幅%"]:+.2f}%
        </span>

        </div>
        """, unsafe_allow_html=True)
with tab_scan:
    st.subheader("📊 股票掃描結果")
    st.write(f"目前顯示 {len(filtered_df)} 檔股票")

def color_status(val):

    if "🔥" in str(val):
        return "background-color: #14532d; color: white"

    elif "👀" in str(val):
        return "background-color: #1e3a8a; color: white"

    elif "⚠️" in str(val):
        return "background-color: #92400e; color: white"

    elif "❌" in str(val):
        return "background-color: #7f1d1d; color: white"

    return ""

styled_df = filtered_df.style.map(
    color_status,
    subset=["狀態"]
)

display_df = filtered_df.copy()

price_columns = [
    "收盤價",
    "5日線",
    "10日線",
    "20日線",
    "20日高點",
    "20日低點"
]

chip_columns = [
    "籌碼1日",
    "籌碼3日",
    "籌碼5日",
    "籌碼10日"
]

for col in chip_columns:
    if col in display_df.columns:
        display_df[col] = display_df[col].map(
            lambda x: "" if pd.isna(x) or x == "" else f"{int(x):,}"
        )

for col in price_columns:
    display_df[col] = display_df[col].map(lambda x: f"{x:,.2f}")

display_df["成交量"] = display_df["成交量"].map(lambda x: f"{int(x):,}")
display_df["5日均量"] = display_df["5日均量"].map(lambda x: f"{int(x):,}")

styled_df = display_df.style.map(
    color_status,
    subset=["狀態"]
)

st.dataframe(
    styled_df,
    use_container_width=True,
    hide_index=True
)

# =========================
# 股票選擇
# =========================

with tab_detail:

    selected_stock = st.selectbox(
        "選擇股票查看K線",
        filtered_df["股票代號"]
    )

    selected_row = filtered_df[filtered_df["股票代號"] == selected_stock].iloc[0]


    st.markdown("## 🔎 個股分析摘要")

    col1, col2, col3, col4, col5 = st.columns(5)

    col1.metric("股票名稱", selected_row["股票名稱"])
    col2.metric("技術分數", selected_row["技術分數"])
    col3.metric("狀態", selected_row["狀態"])
    col4.metric("綜合判斷", selected_row["綜合判斷"])


    st.markdown("### 📌 技術面")
    st.info(selected_row["技術面"])

    st.markdown("### 💰 籌碼面")
    st.success(selected_row["籌碼面"])



# =========================
# 抓K線資料
# =========================

    k_df = yf.download(
        selected_stock,
        period="3mo",
        interval="1d",
        progress=False,
        auto_adjust=False
    )



# =========================
# 均線
# =========================


    open_series = get_series(k_df, "Open")
    high_series = get_series(k_df, "High")
    low_series = get_series(k_df, "Low")
    close_series = get_series(k_df, "Close")
    volume_series = get_series(k_df, "Volume")

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


    

    k_df = k_df[k_df["MA20"].notna()]
    k_df = k_df[k_df["RSI"].notna()]

    open_series = get_series(k_df, "Open")
    high_series = get_series(k_df, "High")
    low_series = get_series(k_df, "Low")
    close_series = get_series(k_df, "Close")
    volume_series = get_series(k_df, "Volume")

# =========================
# 畫K線圖
# =========================

    from plotly.subplots import make_subplots

    fig = make_subplots(
        rows=3,
        cols=1,
        shared_xaxes=True,
        vertical_spacing=0.03,
        row_heights=[0.6, 0.2, 0.2]
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
            decreasing_fillcolor="#22c55e"
        ),
        row=1,
        col=1
    )

# K線
    volume_colors = [
        "#ef4444" if close_series.iloc[i] >= open_series.iloc[i]
        else "#22c55e"
        for i in range(len(k_df))
    ]



# MA5
    fig.add_trace(
        go.Scatter(
            x=k_df.index,
            y=k_df["MA5"],
            mode="lines",
            name="MA5"
        ),
        row=1,
        col=1
    )

    fig.add_trace(
        go.Scatter(
            x=k_df.index,
            y=k_df["MA10"],
            mode="lines",
            name="MA10"
        ),
        row=1,
        col=1
    )

# MA20
    fig.add_trace(
        go.Scatter(
            x=k_df.index,
            y=k_df["MA20"],
            mode="lines",
            name="MA20"
        ),
        row=1,
        col=1
    )



    # RSI 超買超賣線
    fig.add_hline(
        y=70,
        line_dash="dash",
        line_color="red",
        row=3,
        col=1
    )

    fig.add_hline(
        y=30,
        line_dash="dash",
        line_color="green",
        row=3,
        col=1
    )

# =========================
# 成交量
# =========================



    fig.add_trace(
        go.Bar(
            x=k_df.index,
            y=volume_series,
            name="成交量（紅漲綠跌）",
            marker_color=volume_colors
        ),
        row=2,
        col=1
    )

# =========================
# RSI
# =========================

    fig.add_trace(
        go.Scatter(
            x=k_df.index,
            y=k_df["RSI"],
            mode="lines",
            name="RSI",
            line=dict(color="#facc15")
        ),
        row=3,
        col=1
    )

    fig.add_hline(
        y=70,
        line_dash="dash",
        line_color="red",
        row=3,
        col=1
    )

    fig.add_hline(
        y=30,
        line_dash="dash",
        line_color="green",
        row=3,
        col=1
    )

    fig.update_layout(
        height=700,
        xaxis_rangeslider_visible=False,

        showlegend=True,

        xaxis=dict(
            rangebreaks=[
                dict(bounds=["sat", "mon"])
            ]
        )
    )

    st.caption("🔴 紅量 = 收漲　🟢 綠量 = 收跌")
    st.plotly_chart(
    fig,
    use_container_width=True
)