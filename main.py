import os
import pandas as pd
import yfinance as yf

# 建立 output 資料夾
os.makedirs("output", exist_ok=True)

# 讀取股票清單與籌碼資料
watchlist = pd.read_csv("watchlist.csv")
chip_data = pd.read_csv("chip.csv")

results = []


def get_series(df, column):
    data = df[column]
    if isinstance(data, pd.DataFrame):
        return data.iloc[:, 0]
    return data

def generate_ai_comment(judgement, reasons, score):
    reason_text = "、".join(reasons)

    if judgement == "強勢觀察":
        return f"短線偏強，{reason_text}，可以列入觀察，但仍要注意是否追高。"

    elif judgement == "可追蹤":
        return f"目前有部分轉強訊號，{reason_text}，可追蹤但不急著進場。"

    elif judgement == "觀望":
        return f"訊號普通，{reason_text}，目前方向不明，適合先觀察。"

    else:
        return f"目前偏弱，{reason_text}，短線不建議急著進場。"


for index, row in watchlist.iterrows():
    ticker = row["ticker"]
    stock_name = row["name"]

    print(f"正在分析 {stock_name}...")

    # =========================
    # 抓股價資料
    # =========================

    # 先抓 .TW
    df = yf.download(
        ticker,
        period="3mo",
        progress=False,
        auto_adjust=False
    )

    # 如果抓不到，自動改抓 .TWO
    if df.empty and ".TW" in ticker:

        alt_ticker = ticker.replace(".TW", ".TWO")

        print(f"{ticker} 改抓 {alt_ticker}")

        df = yf.download(
            alt_ticker,
            period="3mo",
            progress=False,
            auto_adjust=False
        )

        # 如果 .TWO 成功
        if not df.empty:
            ticker = alt_ticker

    # 如果還是抓不到
    if df.empty:
        print(f"{stock_name} 抓不到資料，跳過")
        continue

    # =========================
    # 籌碼 ticker 修正
    # =========================

    chip_ticker = ticker.replace(".TWO", ".TW")

    # =========================
    # 整理資料
    # =========================

    close_series = pd.to_numeric(
        get_series(df, "Close"),
        errors="coerce"
    ).dropna()

    high_series = pd.to_numeric(
        get_series(df, "High"),
        errors="coerce"
    ).dropna()

    low_series = pd.to_numeric(
        get_series(df, "Low"),
        errors="coerce"
    ).dropna()

    volume_series = pd.to_numeric(
        get_series(df, "Volume"),
        errors="coerce"
    ).dropna()

    # 如果資料不足
    if len(close_series) < 20:
        print(f"{stock_name} 資料不足，跳過")
        continue

    close = float(close_series.iloc[-1])
    ma5_series = close_series.rolling(5).mean()
    ma10_series = close_series.rolling(10).mean()
    ma20_series = close_series.rolling(20).mean()

    ma5 = float(ma5_series.iloc[-1])
    ma10 = float(ma10_series.iloc[-1])
    ma20 = float(ma20_series.iloc[-1])

    prev_ma5 = float(ma5_series.iloc[-2])
    prev_ma20 = float(ma20_series.iloc[-2])

    volume = int(volume_series.iloc[-1])
    avg_volume_5 = float(volume_series.rolling(5).mean().iloc[-1])

    recent_high = float(high_series.tail(20).max())
    recent_low = float(low_series.tail(20).min())

    score = 0

    technical_reasons = []
    chip_reasons = []

    # 預設籌碼欄位
    chip_1d = ""
    chip_3d = ""
    chip_5d = ""
    chip_10d = ""

    # 籌碼判斷
    chip_rows = chip_data[chip_data["ticker"] == chip_ticker]

    if not chip_rows.empty:
        chip_1d = chip_rows["buy_sell_1d"].sum()
        chip_3d = chip_rows["buy_sell_3d"].sum()
        chip_5d = chip_rows["buy_sell_5d"].sum()
        chip_10d = chip_rows["buy_sell_10d"].sum()

        if chip_1d > 0 and chip_3d > 0 and chip_5d > 0:
            score += 30
            chip_reasons.append("籌碼轉強")

        elif chip_10d > 0 and chip_5d < 0 and chip_3d < 0:
            score -= 25
            chip_reasons.append("前期買超但短線轉弱")

        elif chip_10d < 0 and chip_5d < 0 and chip_3d >= 0:
            score += 15
            chip_reasons.append("賣壓縮小")

        elif chip_1d < 0 and chip_3d < 0 and chip_5d < 0:
            score -= 25
            chip_reasons.append("短線賣壓重")

        # 主力強度

        if chip_5d > 500000:
            score += 30
            chip_reasons.append("主力瘋狂買超")

        elif chip_5d > 100000:
            score += 20
            chip_reasons.append("主力大買")

        elif chip_5d > 50000:
            score += 10
            chip_reasons.append("主力偏多")


        if chip_5d < -500000:
            score -= 30
            chip_reasons.append("主力瘋狂倒貨")

        elif chip_5d < -100000:
            score -= 20
            chip_reasons.append("主力大賣")

        elif chip_5d < -50000:
            score -= 10
            chip_reasons.append("主力偏空")

    # 技術面判斷：均線
    if close > ma5:
        score += 20
        technical_reasons.append("站上5日線")
    else:
        score -= 20
        technical_reasons.append("跌破5日線")

    if ma5 > ma10 > ma20:
        score += 25
        technical_reasons.append("多頭排列")
    elif ma5 < ma10 < ma20:
        score -= 25
        technical_reasons.append("空頭排列")

    # 黃金交叉

    if prev_ma5 <= prev_ma20 and ma5 > ma20:
        score += 25
        technical_reasons.append("黃金交叉")

    # 死亡交叉

    if prev_ma5 >= prev_ma20 and ma5 < ma20 * 0.995:
        score -= 25
        technical_reasons.append("死亡交叉")

    # 技術面判斷：成交量

    if volume > avg_volume_5 * 2:
        score += 20
        technical_reasons.append("超級爆量")

    elif volume > avg_volume_5 * 1.5:
        score += 15
        technical_reasons.append("爆量")

    elif volume > avg_volume_5:
        score += 10
        technical_reasons.append("成交量放大")

    else:
        technical_reasons.append("成交量未放大")


    # 技術面判斷：支撐壓力
    if close >= recent_high:
        score += 25
        technical_reasons.append("突破20日高點")

    elif close >= recent_high * 0.98:
        score += 15
        technical_reasons.append("接近20日高點")

    if close < recent_low:
        score -= 25
        technical_reasons.append("跌破20日低點")

    elif close <= recent_low * 1.05:
        score -= 10
        technical_reasons.append("接近20日低點")


    # 綜合判斷
    if score >= 40:
        judgement = "強勢觀察"
    elif score >= 20:
        judgement = "可追蹤"
    elif score >= 0:
        judgement = "觀望"
    else:
        judgement = "避開"

    # 狀態標示
    if score >= 70:
        status = "🔥強勢"
    elif score >= 50:
        status = "👀觀察"
    elif score >= 20:
        status = "⚠️整理"
    else:
        status = "❌避開"

    results.append({
        "股票代號": ticker,
        "股票名稱": stock_name,
        "收盤價": round(close, 2),
        "5日線": round(ma5, 2),
        "10日線": round(ma10, 2),
        "20日線": round(ma20, 2),
        "成交量": volume,
        "5日均量": round(avg_volume_5, 0),
        "20日高點": round(recent_high, 2),
        "20日低點": round(recent_low, 2),
        "籌碼1日": chip_1d,
        "籌碼3日": chip_3d,
        "籌碼5日": chip_5d,
        "籌碼10日": chip_10d,
        "技術分數": score,
        "狀態": status,
        "綜合判斷": judgement,
        "技術面": "、".join(technical_reasons),
        "籌碼面": "、".join(chip_reasons),
        "AI評語": generate_ai_comment(
            judgement,
            technical_reasons,
            score,
        ),
    })


result_df = pd.DataFrame(results)

# 分數高的排前面
result_df = result_df.sort_values(
    by="技術分數",
    ascending=False
)

result_df.to_excel(
    "output/stock_analysis_result.xlsx",
    index=False
)

print("完成：已輸出 output/stock_analysis_result.xlsx")