from datetime import date
from pathlib import Path

import pandas as pd
import requests
import urllib3


BASE_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = BASE_DIR / "output"
WATCHLIST_FILE = BASE_DIR / "watchlist.csv"
UNIVERSE_FILE = BASE_DIR / "universe.csv"
CHIP_FILE = BASE_DIR / "chip.csv"
CHIP_DAILY_FILE = OUTPUT_DIR / "chip_daily.csv"
RESULT_FILE = OUTPUT_DIR / "stock_analysis_result.xlsx"
UNIVERSE_RESULT_FILE = OUTPUT_DIR / "universe_analysis_result.xlsx"
HISTORY_FILE = OUTPUT_DIR / "stock_analysis_history.csv"
MIN_RESULT_SUCCESS_RATIO = 0.8

OUTPUT_DIR.mkdir(exist_ok=True)
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


def stock_code_key(value):
    return str(value).strip().split(".")[0]


def to_float(value):
    try:
        text = str(value).replace(",", "").replace("--", "").strip()
        if text in ["", "-", "nan", "None"]:
            return 0
        return float(text)
    except (TypeError, ValueError):
        return 0


def latest_chip_daily_date():
    if not CHIP_DAILY_FILE.exists():
        return ""
    try:
        chip_daily_df = pd.read_csv(CHIP_DAILY_FILE, usecols=["date"])
    except Exception:
        return ""
    dates = pd.to_datetime(chip_daily_df["date"], errors="coerce").dropna()
    if dates.empty:
        return ""
    return dates.max().date().isoformat()


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


def load_previous_scores():
    if not HISTORY_FILE.exists():
        return {}

    try:
        history_df = pd.read_csv(HISTORY_FILE)
    except Exception:
        return {}

    required_columns = {"snapshot_date", "股票代號"}
    if history_df.empty or not required_columns.issubset(history_df.columns):
        return {}

    score_column = "DeepTrend分數"
    if score_column not in history_df.columns:
        return {}

    history_df["snapshot_date"] = pd.to_datetime(history_df["snapshot_date"], errors="coerce")
    history_df[score_column] = pd.to_numeric(history_df[score_column], errors="coerce")
    history_df["股票代號_key"] = history_df["股票代號"].map(stock_code_key)
    history_df = history_df.dropna(subset=["snapshot_date", score_column])

    today = pd.Timestamp(date.today())
    previous_df = history_df[history_df["snapshot_date"] < today]
    if previous_df.empty:
        return {}

    latest_date = previous_df["snapshot_date"].max()
    latest_df = previous_df[previous_df["snapshot_date"] == latest_date]
    return dict(zip(latest_df["股票代號_key"], latest_df[score_column]))


def calculate_score_change(current_score, previous_score):
    if previous_score is None or pd.isna(previous_score):
        return None, None

    change = current_score - previous_score
    if previous_score == 0:
        return round(change, 2), None

    change_rate = change / abs(previous_score) * 100
    return round(change, 2), round(change_rate, 2)


def score_chip(chip_1d, chip_3d, chip_5d, chip_10d, foreign_5d, investment_5d):
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

    if foreign_5d > 0 and investment_5d > 0:
        score += 20
        reasons.append("外資投信同步買超")
    elif foreign_5d < 0 and investment_5d < 0:
        score -= 20
        reasons.append("外資投信同步賣超")
    elif foreign_5d > 0 and investment_5d < 0:
        reasons.append("外資買投信賣")
    elif foreign_5d < 0 and investment_5d > 0:
        reasons.append("投信買外資賣")

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


def infer_asset_type(row, ticker):
    value = str(row.get("asset_type", "")).strip().lower()
    if value in {"etf", "stock"}:
        return "ETF" if value == "etf" else "個股"

    group = str(row.get("group", "")).strip().lower()
    ticker_key = stock_code_key(ticker)
    if group == "etf" or ticker_key.startswith("00"):
        return "ETF"
    return "個股"


def analyze_stock_list(input_file, output_file, label, use_previous_scores=True):
    stock_list = pd.read_csv(input_file)
    chip_data = pd.read_csv(CHIP_FILE)
    today_text = date.today().isoformat()
    chip_latest_date = latest_chip_daily_date()
    if chip_latest_date != today_text:
        print(
            f"chip_daily.csv latest date is {chip_latest_date or 'missing'}, not {today_text}. "
            "Using the latest available trading day."
        )

    previous_scores = load_previous_scores() if use_previous_scores else {}
    results = []

    for _, row in stock_list.iterrows():
        original_ticker = str(row["ticker"]).strip()
        stock_name = row["name"]
        ticker, history = get_official_history(original_ticker)
        asset_type = infer_asset_type(row, ticker)

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
        foreign_1d = foreign_3d = foreign_5d = foreign_10d = 0
        investment_1d = investment_3d = investment_5d = investment_10d = 0
        dealer_1d = dealer_3d = dealer_5d = dealer_10d = 0
        if not chip_rows.empty:
            chip_1d = int(chip_rows["buy_sell_1d"].sum())
            chip_3d = int(chip_rows["buy_sell_3d"].sum())
            chip_5d = int(chip_rows["buy_sell_5d"].sum())
            chip_10d = int(chip_rows["buy_sell_10d"].sum())
            foreign_1d = int(chip_rows["foreign_1d"].sum()) if "foreign_1d" in chip_rows else 0
            foreign_3d = int(chip_rows["foreign_3d"].sum()) if "foreign_3d" in chip_rows else 0
            foreign_5d = int(chip_rows["foreign_5d"].sum()) if "foreign_5d" in chip_rows else 0
            foreign_10d = int(chip_rows["foreign_10d"].sum()) if "foreign_10d" in chip_rows else 0
            investment_1d = int(chip_rows["investment_1d"].sum()) if "investment_1d" in chip_rows else 0
            investment_3d = int(chip_rows["investment_3d"].sum()) if "investment_3d" in chip_rows else 0
            investment_5d = int(chip_rows["investment_5d"].sum()) if "investment_5d" in chip_rows else 0
            investment_10d = int(chip_rows["investment_10d"].sum()) if "investment_10d" in chip_rows else 0
            dealer_1d = int(chip_rows["dealer_1d"].sum()) if "dealer_1d" in chip_rows else 0
            dealer_3d = int(chip_rows["dealer_3d"].sum()) if "dealer_3d" in chip_rows else 0
            dealer_5d = int(chip_rows["dealer_5d"].sum()) if "dealer_5d" in chip_rows else 0
            dealer_10d = int(chip_rows["dealer_10d"].sum()) if "dealer_10d" in chip_rows else 0

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
        chip_score, chip_reasons = score_chip(
            chip_1d,
            chip_3d,
            chip_5d,
            chip_10d,
            foreign_5d,
            investment_5d,
        )
        volume_price_score, volume_price_signals = detect_volume_price_signal(
            close,
            prev_close,
            volume,
            avg_volume_5,
            previous_20d_high,
        )
        score = technical_score * 0.4 + chip_score * 0.4 + volume_price_score * 0.2
        score = round(score, 2)
        previous_score = previous_scores.get(stock_code_key(ticker))
        score_change, score_change_rate = calculate_score_change(score, previous_score)
        technical_reasons = technical_reasons + [
            signal for signal in volume_price_signals if signal != "無明顯異常"
        ]
        status, judgement = classify_score(score)

        results.append(
            {
                "股票代號": ticker,
                "股票名稱": stock_name,
                "資產類型": asset_type,
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
                "外資1日": foreign_1d,
                "外資3日": foreign_3d,
                "外資5日": foreign_5d,
                "外資10日": foreign_10d,
                "投信1日": investment_1d,
                "投信3日": investment_3d,
                "投信5日": investment_5d,
                "投信10日": investment_10d,
                "自營商1日": dealer_1d,
                "自營商3日": dealer_3d,
                "自營商5日": dealer_5d,
                "自營商10日": dealer_10d,
                "技術分數": score,
                "DeepTrend分數": score,
                "技術面分數": technical_score,
                "籌碼分數": chip_score,
                "量價分數": volume_price_score,
                "前次分數": previous_score,
                "分數變化": score_change,
                "分數變化率": score_change_rate,
                "狀態": status,
                "綜合判斷": judgement,
                "技術面": "、".join(technical_reasons),
                "籌碼面": "、".join(chip_reasons),
                "AI評語": generate_ai_comment(judgement, technical_reasons, chip_reasons, score),
            }
        )

    result_df = pd.DataFrame(results)

    if result_df.empty:
        raise RuntimeError(f"沒有產生任何 {label} 分析結果，請檢查官方資料來源或 {input_file.name}。")

    min_required_rows = int(len(stock_list) * MIN_RESULT_SUCCESS_RATIO)
    if len(result_df) < min_required_rows:
        raise RuntimeError(
            f"{label} analysis result is incomplete: {len(result_df)}/{len(stock_list)} rows. "
            f"Need at least {min_required_rows}. Existing result file was not overwritten."
        )

    result_df = result_df.sort_values(by="技術分數", ascending=False)
    result_df.to_excel(output_file, index=False)
    print(f"完成 {label}，已輸出 {output_file}")


def main():
    analyze_stock_list(WATCHLIST_FILE, RESULT_FILE, "watchlist", use_previous_scores=True)

    if UNIVERSE_FILE.exists():
        try:
            analyze_stock_list(UNIVERSE_FILE, UNIVERSE_RESULT_FILE, "universe", use_previous_scores=False)
        except Exception as exc:
            print(f"universe 分析失敗，保留上一份市場池結果：{exc}")


if __name__ == "__main__":
    main()
