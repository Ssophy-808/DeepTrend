import pandas as pd
import yfinance as yf

# =========================
# 設定
# =========================

TICKER = "0050.TW"
PERIOD = "1mo"
INTERVAL = "5m"

OUTPUT_FILE = "index_strategy_signals.xlsx"

# =========================
# 抓資料
# =========================

print(f"正在抓資料：{TICKER}")

df = yf.download(
    TICKER,
    period=PERIOD,
    interval=INTERVAL,
    progress=False,
    auto_adjust=False
)

if df.empty:
    print("抓不到資料")
    exit()

# =========================
# 修正欄位
# =========================

def get_series(df, column_name):
    data = df[column_name]

    if isinstance(data, pd.DataFrame):
        return data.iloc[:, 0]

    return data

close = pd.to_numeric(
    get_series(df, "Close"),
    errors="coerce"
).dropna()

df["Close_Price"] = close

# =========================
# 均線
# =========================

df["MA5"] = df["Close_Price"].rolling(5).mean()
df["MA60"] = df["Close_Price"].rolling(60).mean()

# 前一根K棒資料

df["Prev_Close"] = df["Close_Price"].shift(1)
df["Prev_MA5"] = df["MA5"].shift(1)
df["Prev_MA60"] = df["MA60"].shift(1)

# =========================
# 買進訊號
# =========================

df["Buy_Signal"] = (
    (df["Close_Price"] > df["MA60"]) &
    (df["Prev_Close"] <= df["Prev_MA60"]) &
    (df["MA5"] > df["MA60"]) &
    (df["Prev_MA5"] <= df["Prev_MA60"])
)

# =========================
# 賣出訊號
# =========================

df["Sell_Signal"] = (
    (df["Close_Price"] < df["MA60"]) &
    (df["Prev_Close"] >= df["Prev_MA60"]) &
    (df["MA5"] < df["MA60"]) &
    (df["Prev_MA5"] >= df["Prev_MA60"])
)

# =========================
# 只保留有訊號的
# =========================

signals = df[
    (df["Buy_Signal"] == True) |
    (df["Sell_Signal"] == True)
].copy()

# =========================
# 訊號文字
# =========================

# =========================
# 訊號文字
# =========================

signal_text = []

for i in range(len(signals)):
    buy_value = bool(signals["Buy_Signal"].iloc[i])
    sell_value = bool(signals["Sell_Signal"].iloc[i])

    if buy_value:
        signal_text.append("買進")
    elif sell_value:
        signal_text.append("賣出")
    else:
        signal_text.append("")

signals["訊號"] = signal_text

# =========================
# 輸出欄位
# =========================

signals = signals[
    [
        "Close_Price",
        "MA5",
        "MA60",
        "訊號"
    ]
]

# =========================
# 輸出 Excel
# =========================
signals.index = signals.index.tz_localize(None)

signals.to_excel(
    OUTPUT_FILE,
    index=True
)

# 額外輸出最近K線
df.index = df.index.tz_localize(None)

df.tail(100).to_excel(
    "latest_kline_check.xlsx"
)

print(f"完成：已輸出 {OUTPUT_FILE}")
print("已輸出 latest_kline_check.xlsx")