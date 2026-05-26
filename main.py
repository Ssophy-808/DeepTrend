from datetime import date
from pathlib import Path

import pandas as pd
import requests
import urllib3


BASE_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = BASE_DIR / "output"
WATCHLIST_FILE = BASE_DIR / "watchlist.csv"
CHIP_FILE = BASE_DIR / "chip.csv"
RESULT_FILE = OUTPUT_DIR / "stock_analysis_result.xlsx"

OUTPUT_DIR.mkdir(exist_ok=True)
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


def to_float(value):
    try:
        text = str(value).replace(",", "").replace("--", "").strip()
        if text in ["", "-", "nan", "None"]:
            return 0
        return float(text)
    except (TypeError, ValueError):
        return 0


def parse_roc_date(value):
    parts = str(value).split("/")
    if len(parts) != 3:
        return None

    try:
        return date(int(parts[0]) + 1911, int(parts[1]), int(parts[2]))
    except ValueError:
        return None


def recent_month_starts(month_count=4):
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


def normalize_chip_ticker(ticker):
    code = str(ticker).strip().split(".")[0]
    return f"{code}.TW"


def fetch_twse_history(code):
    rows = []
    headers = {"User-Agent": "Mozilla/5.0"}

    for month_start in recent_month_starts():
        try:
            response = requests.get(
                "https://www.twse.com.tw/exchangeReport/STOCK_DAY",
                params={"response": "json", "date": month_start, "stockNo": code},
                timeout=15,
                verify=False,
                headers=headers,
            )
            payload = response.json()
        except Exception:
            continue

        if payload.get("stat") != "OK":
            continue

        for row in payload.get("data", []):
            trade_date = parse_roc_date(row[0])
            if not trade_date:
                continue

            rows.append(
                {
                    "日期": trade_date,
                    "開盤價": to_float(row[3]),
                    "最高價": to_float(row[4]),
                    "最低價": to_float(row[5]),
                    "收盤價": to_float(row[6]),
                    "成交量": to_float(row[1]),
                }
            )

    return rows


def fetch_tpex_history(code):
    rows = []
    headers = {"User-Agent": "Mozilla/5.0"}

    for month_start in recent_month_starts():
        try:
            response = requests.get(
                "https://www.tpex.org.tw/www/zh-tw/afterTrading/tradingStock",
                params={"date": f"{month_start[:4]}/{month_start[4:6]}/01", "code": code, "response": "json"},
                timeout=15,
                verify=False,
                headers=headers,
            )
            tables = response.json().get("tables", [])
        except Exception:
            continue

        data_rows = tables[0].get("data", []) if tables else []

        for row in data_rows:
            trade_date = parse_roc_date(row[0])
            if not trade_date:
                continue

            rows.append(
                {
                    "日期": trade_date,
                    "開盤價": to_float(row[3]),
                    "最高價": to_float(row[4]),
                    "最低價": to_float(row[5]),
                    "收盤價": to_float(row[6]),
                    "成交量": to_float(row[1]) * 1000,
                }
            )

    return rows


def get_official_history(ticker):
    text = str(ticker).strip()
    code = text.split(".")[0]

    if text.endswith(".TWO"):
        candidates = [("TWO", fetch_tpex_history), ("TW", fetch_twse_history)]
    else:
        candidates = [("TW", fetch_twse_history), ("TWO", fetch_tpex_history)]

    for suffix, fetcher in candidates:
        rows = fetcher(code)
        if rows:
            history = pd.DataFrame(rows)
            history = history.dropna(subset=["日期"])
            history = history.sort_values("日期").drop_duplicates(subset=["日期"], keep="last")
            history = history.tail(80).reset_index(drop=True)
            return f"{code}.{suffix}", history

    return text, pd.DataFrame()


def generate_ai_comment(judgement, technical_reasons, chip_reasons, score):
    reasons = technical_reasons + chip_reasons
    reason_text = "、".join(reasons) if reasons else "目前訊號不明顯"

    if score >= 70:
        return f"強勢偏多：{reason_text}。可列入優先觀察，但仍需留意乖離與量能是否過熱。"
    if score >= 50:
        return f"偏多觀察：{reason_text}。條件不差，適合等待拉回或突破確認。"
    if score >= 20:
        return f"整理觀察：{reason_text}。多空訊號混合，建議降低追價衝動。"

    return f"風險偏高：{reason_text}。目前不適合積極追蹤，等待結構改善。"


def score_chip(chip_1d, chip_3d, chip_5d, chip_10d):
    score = 0
    reasons = []

    if chip_1d > 0 and chip_3d > 0 and chip_5d > 0:
        score += 30
        reasons.append("法人連續偏買")
    elif chip_1d < 0 and chip_3d < 0 and chip_5d < 0:
        score -= 25
        reasons.append("法人連續偏賣")

    if chip_10d > 0 and chip_5d < 0 and chip_3d < 0:
        score -= 25
        reasons.append("中期買超轉短線賣壓")
    elif chip_10d < 0 and chip_5d < 0 and chip_3d >= 0:
        score += 15
        reasons.append("短線籌碼止穩")

    if chip_5d > 500000:
        score += 30
        reasons.append("5日法人明顯買超")
    elif chip_5d > 100000:
        score += 20
        reasons.append("5日法人買超")
    elif chip_5d > 50000:
        score += 10
        reasons.append("5日法人小幅買超")

    if chip_5d < -500000:
        score -= 30
        reasons.append("5日法人明顯賣超")
    elif chip_5d < -100000:
        score -= 20
        reasons.append("5日法人賣超")
    elif chip_5d < -50000:
        score -= 10
        reasons.append("5日法人小幅賣超")

    if not reasons:
        reasons.append("籌碼中性")

    return score, reasons


def detect_volume_price_signal(close, prev_close, volume, avg_volume_5, previous_20d_high):
    signals = []
    score = 0
    change_pct = ((close - prev_close) / prev_close * 100) if prev_close else 0
    volume_ratio = (volume / avg_volume_5) if avg_volume_5 else 0

    if volume_ratio >= 2 and change_pct <= 0.5:
        signals.append("爆量價未漲")
        score -= 20

    if volume_ratio >= 1.5 and close > previous_20d_high:
        signals.append("帶量突破20日高點")
        score += 25

    if not signals:
        signals.append("無明顯異常")

    return score, signals


def score_technical(close, ma5, ma10, ma20, prev_ma5, prev_ma20, volume, avg_volume_5, recent_high, recent_low):
    score = 0
    reasons = []

    if close > ma5:
        score += 20
        reasons.append("站上5日線")
    else:
        score -= 20
        reasons.append("跌破5日線")

    if ma5 > ma10 > ma20:
        score += 25
        reasons.append("多頭排列")
    elif ma5 < ma10 < ma20:
        score -= 25
        reasons.append("空頭排列")

    if prev_ma5 <= prev_ma20 and ma5 > ma20:
        score += 25
        reasons.append("黃金交叉")

    if prev_ma5 >= prev_ma20 and ma5 < ma20 * 0.995:
        score -= 25
        reasons.append("死亡交叉")

    if volume > avg_volume_5 * 2:
        score += 20
        reasons.append("爆量")
    elif volume > avg_volume_5 * 1.5:
        score += 15
        reasons.append("成交量放大")
    elif volume > avg_volume_5:
        score += 10
        reasons.append("量能溫和放大")
    else:
        reasons.append("量能未放大")

    if close >= recent_high:
        score += 25
        reasons.append("突破20日高點")
    elif close >= recent_high * 0.98:
        score += 15
        reasons.append("接近20日高點")

    if close < recent_low:
        score -= 25
        reasons.append("跌破20日低點")
    elif close <= recent_low * 1.05:
        score -= 10
        reasons.append("接近20日低點")

    return score, reasons


def classify_score(score):
    if score >= 70:
        return "🔥強勢", "強勢觀察"
    if score >= 50:
        return "👀觀察", "偏多觀察"
    if score >= 20:
        return "⚠️整理", "整理觀察"
    return "❌避開", "觀望"


def main():
    watchlist = pd.read_csv(WATCHLIST_FILE)
    chip_data = pd.read_csv(CHIP_FILE)
    results = []

    for _, row in watchlist.iterrows():
        original_ticker = str(row["ticker"]).strip()
        stock_name = row["name"]
        ticker, history = get_official_history(original_ticker)

        print(f"正在分析 {stock_name} ({ticker})...")

        if history.empty or len(history) < 20:
            print(f"{stock_name} 官方日K資料不足，跳過")
            continue

        close_series = pd.to_numeric(history["收盤價"], errors="coerce").dropna()
        high_series = pd.to_numeric(history["最高價"], errors="coerce").dropna()
        low_series = pd.to_numeric(history["最低價"], errors="coerce").dropna()
        volume_series = pd.to_numeric(history["成交量"], errors="coerce").dropna()

        if len(close_series) < 20:
            print(f"{stock_name} 收盤價資料不足，跳過")
            continue

        close = float(close_series.iloc[-1])
        prev_close = float(close_series.iloc[-2])
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
        previous_20d_high = float(high_series.iloc[:-1].tail(20).max())

        chip_ticker = normalize_chip_ticker(ticker)
        chip_rows = chip_data[chip_data["ticker"].astype(str) == chip_ticker]

        chip_1d = chip_3d = chip_5d = chip_10d = 0
        if not chip_rows.empty:
            chip_1d = int(chip_rows["buy_sell_1d"].sum())
            chip_3d = int(chip_rows["buy_sell_3d"].sum())
            chip_5d = int(chip_rows["buy_sell_5d"].sum())
            chip_10d = int(chip_rows["buy_sell_10d"].sum())

        technical_score, technical_reasons = score_technical(
            close,
            ma5,
            ma10,
            ma20,
            prev_ma5,
            prev_ma20,
            volume,
            avg_volume_5,
            recent_high,
            recent_low,
        )
        chip_score, chip_reasons = score_chip(chip_1d, chip_3d, chip_5d, chip_10d)
        volume_price_score, volume_price_signals = detect_volume_price_signal(
            close,
            prev_close,
            volume,
            avg_volume_5,
            previous_20d_high,
        )
        score = technical_score + chip_score + volume_price_score
        technical_reasons = technical_reasons + [
            signal for signal in volume_price_signals if signal != "無明顯異常"
        ]
        status, judgement = classify_score(score)

        results.append(
            {
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
                "量價異常": "、".join(volume_price_signals),
                "籌碼1日": chip_1d,
                "籌碼3日": chip_3d,
                "籌碼5日": chip_5d,
                "籌碼10日": chip_10d,
                "技術分數": score,
                "狀態": status,
                "綜合判斷": judgement,
                "技術面": "、".join(technical_reasons),
                "籌碼面": "、".join(chip_reasons),
                "AI評語": generate_ai_comment(judgement, technical_reasons, chip_reasons, score),
            }
        )

    result_df = pd.DataFrame(results)

    if result_df.empty:
        raise RuntimeError("沒有產生任何分析結果，請檢查官方資料來源或 watchlist.csv。")

    result_df = result_df.sort_values(by="技術分數", ascending=False)
    result_df.to_excel(RESULT_FILE, index=False)
    print(f"完成，已輸出 {RESULT_FILE}")


if __name__ == "__main__":
    main()
