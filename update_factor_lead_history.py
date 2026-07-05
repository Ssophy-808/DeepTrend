from pathlib import Path

import pandas as pd


BASE_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = BASE_DIR / "output"
HISTORY_FILE = OUTPUT_DIR / "stock_analysis_history.csv"
FACTOR_LEAD_FILE = OUTPUT_DIR / "factor_lead_history.csv"

FACTOR_RULES = [
    {
        "lead_factor": "技術面",
        "score_column": "技術面分數",
        "drop_threshold": -20,
        "flip_negative": False,
    },
    {
        "lead_factor": "籌碼面",
        "score_column": "籌碼分數",
        "drop_threshold": -30,
        "flip_negative": True,
    },
    {
        "lead_factor": "量價面",
        "score_column": "量價分數",
        "drop_threshold": -20,
        "flip_negative": False,
    },
]

FORWARD_DAYS = [1, 3, 5, 10]
PRICE_DROP_THRESHOLD = -3


def to_number(series):
    return pd.to_numeric(series, errors="coerce")


def prepare_history():
    required_columns = {
        "snapshot_date",
        "股票代號",
        "股票名稱",
        "收盤價",
        "DeepTrend分數",
        "技術面分數",
        "籌碼分數",
        "量價分數",
    }
    if not HISTORY_FILE.exists():
        print(f"Skip: {HISTORY_FILE} does not exist.")
        return pd.DataFrame()

    history_df = pd.read_csv(HISTORY_FILE)
    missing_columns = sorted(required_columns - set(history_df.columns))
    if missing_columns:
        print(f"Skip: missing columns: {', '.join(missing_columns)}")
        return pd.DataFrame()

    history_df = history_df.copy()
    history_df["snapshot_date"] = pd.to_datetime(history_df["snapshot_date"], errors="coerce")
    for column in ["收盤價", "DeepTrend分數", "技術面分數", "籌碼分數", "量價分數"]:
        history_df[column] = to_number(history_df[column])

    history_df = history_df.dropna(subset=["snapshot_date", "股票代號", "收盤價"])
    history_df = history_df.sort_values(["股票代號", "snapshot_date"])
    history_df = history_df.drop_duplicates(subset=["股票代號", "snapshot_date"], keep="last")
    return history_df.reset_index(drop=True)


def add_forward_price_columns(stock_rows):
    stock_rows = stock_rows.sort_values("snapshot_date").reset_index(drop=True).copy()
    for days in FORWARD_DAYS:
        stock_rows[f"close_{days}d"] = stock_rows["收盤價"].shift(-days)
        stock_rows[f"return_{days}d"] = (
            (stock_rows[f"close_{days}d"] - stock_rows["收盤價"]) / stock_rows["收盤價"] * 100
        )
    return stock_rows


def first_price_drop_days(row):
    for days in FORWARD_DAYS:
        value = row.get(f"return_{days}d")
        if pd.notna(value) and value <= PRICE_DROP_THRESHOLD:
            return days
    return pd.NA


def detect_factor_events(history_df):
    rows = []
    if history_df.empty:
        return pd.DataFrame()

    for _, stock_rows in history_df.groupby("股票代號"):
        stock_rows = add_forward_price_columns(stock_rows)
        if len(stock_rows) < 2:
            continue

        stock_rows["prev_deeptrend"] = stock_rows["DeepTrend分數"].shift(1)
        for rule in FACTOR_RULES:
            column = rule["score_column"]
            prev_column = f"prev_{column}"
            change_column = f"{column}_change"

            stock_rows[prev_column] = stock_rows[column].shift(1)
            stock_rows[change_column] = stock_rows[column] - stock_rows[prev_column]

            drop_event = stock_rows[change_column].le(rule["drop_threshold"])
            if rule["flip_negative"]:
                flip_event = stock_rows[prev_column].gt(0) & stock_rows[column].lt(0)
                event_mask = drop_event | flip_event
            else:
                event_mask = drop_event

            event_rows = stock_rows[event_mask].copy()
            for _, row in event_rows.iterrows():
                lead_days = first_price_drop_days(row)
                event = {
                    "event_date": row["snapshot_date"].date().isoformat(),
                    "stock_id": row["股票代號"],
                    "stock_name": row["股票名稱"],
                    "lead_factor": rule["lead_factor"],
                    "event_type": "因子轉弱",
                    "factor_before": row[prev_column],
                    "factor_after": row[column],
                    "factor_change": row[change_column],
                    "deeptrend_before": row["prev_deeptrend"],
                    "deeptrend_after": row["DeepTrend分數"],
                    "close_at_event": row["收盤價"],
                    "price_drop_after": pd.notna(lead_days),
                    "lead_days": lead_days,
                }
                for days in FORWARD_DAYS:
                    event[f"close_{days}d"] = row.get(f"close_{days}d")
                    event[f"return_{days}d"] = row.get(f"return_{days}d")
                rows.append(event)

    if not rows:
        return pd.DataFrame()

    result_df = pd.DataFrame(rows)
    result_df = result_df.sort_values(["event_date", "stock_id", "lead_factor"])
    result_df = result_df.drop_duplicates(
        subset=["event_date", "stock_id", "lead_factor", "event_type"],
        keep="last",
    )
    return result_df.reset_index(drop=True)


def main():
    OUTPUT_DIR.mkdir(exist_ok=True)
    history_df = prepare_history()
    factor_df = detect_factor_events(history_df)

    if factor_df.empty:
        print("No factor lead events found. No file was written.")
        return

    factor_df.to_csv(FACTOR_LEAD_FILE, index=False, encoding="utf-8-sig")
    print(f"Wrote {len(factor_df)} factor lead events to {FACTOR_LEAD_FILE}")


if __name__ == "__main__":
    main()
