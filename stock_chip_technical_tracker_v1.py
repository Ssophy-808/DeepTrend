import pandas as pd
import yfinance as yf

OUTPUT_FILE = "stock_analysis_result.xlsx"

stocks = [
    ("2330.TW", "台積電"),
    ("2317.TW", "鴻海"),
    ("3481.TW", "群創"),
    ("3105.TW", "穩懋"),
]

def technical_analysis(df):
   def technical_analysis(df):
    close = float(df["Close"].iloc[-1])
    ma5 = float(df["Close"].rolling(5).mean().iloc[-1])
    ma10 = float(df["Close"].rolling(10).mean().iloc[-1])
    volume = df["Volume"].iloc[-1]
    avg_volume_5d = df["Volume"].rolling(5).mean().iloc[-1]

    score = 0
    reasons = []

    if close > ma5:
        score += 1
        reasons.append("站上5日線")
    else:
        reasons.append("跌破5日線")

    if close > ma10:
        score += 1
        reasons.append("站上10日線")

    if volume > avg_volume_5d:
        score += 1
        reasons.append("成交量放大")

    if score >= 3:
        judgement = "強勢"
    elif score == 2:
        judgement = "觀察"
    else:
        judgement = "弱勢"

    return score, judgement, "、".join(reasons)

def main():
    results = []

    for stock_id, stock_name in stocks:
        df = yf.download(stock_id, period="3mo")

        if len(df) < 20:
            continue

        def technical_analysis(df):
    close = float(df["Close"].iloc[-1])
    ma5 = float(df["Close"].rolling(5).mean().iloc[-1])
    ma10 = float(df["Close"].rolling(10).mean().iloc[-1])

        score, judgement, reasons = technical_analysis(df)

        results.append({
            "股票代號": stock_id,
            "股票名稱": stock_name,
            "收盤價": close,
            "MA5": ma5,
            "MA10": ma10,
            "技術分數": score,
            "綜合判斷": judgement,
            "原因": reasons
        })

    result_df = pd.DataFrame(results)

    result_df.to_excel(OUTPUT_FILE, index=False)

    print(f"完成：已輸出 {OUTPUT_FILE}")

if __name__ == "__main__":
    main()