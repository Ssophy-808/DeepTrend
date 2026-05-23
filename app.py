import streamlit as st
import pandas as pd
import yfinance as yf
import plotly.graph_objects as go
import subprocess

st.set_page_config(
    page_title="DeepTrend",
    page_icon="🔥",
    layout="wide"
)

st.title("🔥 DeepTrend")
st.caption("AI Quant Trading Radar")

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

if selected_status != "全部":
    filtered_df = filtered_df[filtered_df["狀態"] == selected_status]

filtered_df = filtered_df[filtered_df["技術分數"] >= min_score]

if keyword:
    filtered_df = filtered_df[
        filtered_df["股票名稱"].astype(str).str.contains(keyword, case=False, na=False) |
        filtered_df["股票代號"].astype(str).str.contains(keyword, case=False, na=False)
    ]

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

st.dataframe(
    styled_df,
    use_container_width=True,
    hide_index=True
)

# =========================
# 股票選擇
# =========================

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
def get_series(df, column):
    data = df[column]
    if isinstance(data, pd.DataFrame):
        return data.iloc[:, 0]
    return data

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