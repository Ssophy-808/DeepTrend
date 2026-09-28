import argparse
import io
from pathlib import Path
import subprocess

import pandas as pd


BASE_DIR = Path(__file__).resolve().parent
HISTORY_FILE = BASE_DIR / "output" / "stock_analysis_history.csv"
RESULT_FILE = BASE_DIR / "output" / "stock_analysis_result.xlsx"
UNIVERSE_RESULT_FILE = BASE_DIR / "output" / "universe_analysis_result.xlsx"
FACTOR_FILE = BASE_DIR / "output" / "factor_lead_history.csv"
CHIP_DAILY_FILE = BASE_DIR / "output" / "chip_daily.csv"


def stock_code_key(value):
    return str(value).strip().split(".")[0]


def choose_daily_snapshot(group):
    signature_columns = [
        column
        for column in ["收盤價", "DeepTrend分數", "技術面分數", "籌碼分數", "量價分數"]
        if column in group.columns
    ]
    if not signature_columns:
        return group.iloc[-1]

    signatures = group[signature_columns].fillna("").astype(str).agg("|".join, axis=1)
    counts = signatures.map(signatures.value_counts())
    best_count = counts.max()
    return group.loc[counts[counts.eq(best_count)].index].iloc[-1]


def repair_history(invalid_dates):
    history_df = pd.read_csv(HISTORY_FILE, low_memory=False)
    before_rows = len(history_df)
    parsed_dates = pd.to_datetime(history_df["snapshot_date"], errors="coerce")
    invalid_mask = parsed_dates.dt.dayofweek.ge(5) | parsed_dates.dt.strftime("%Y-%m-%d").isin(invalid_dates)
    history_df = history_df.loc[~invalid_mask & parsed_dates.notna()].copy()
    history_df["snapshot_date"] = parsed_dates.loc[history_df.index].dt.strftime("%Y-%m-%d")
    history_df["股票代號_key"] = history_df["股票代號"].map(stock_code_key)
    history_df["_row_order"] = range(len(history_df))
    history_df = history_df.sort_values(["snapshot_date", "_row_order"])
    daily_rows = [
        choose_daily_snapshot(group)
        for _, group in history_df.groupby(["snapshot_date", "股票代號_key"], sort=False)
    ]
    history_df = pd.DataFrame(daily_rows).reset_index(drop=True)
    history_df["資料日期"] = history_df["snapshot_date"]
    history_df = history_df.drop(columns=["股票代號_key", "_row_order"], errors="ignore")
    history_df.to_csv(HISTORY_FILE, index=False, encoding="utf-8-sig")
    print(f"History repaired: {before_rows} -> {len(history_df)} rows")
    return history_df


def restore_current_result(history_df):
    result_df = pd.read_excel(RESULT_FILE)
    result_df["股票代號_key"] = result_df["股票代號"].map(stock_code_key)
    latest_market_date = history_df["snapshot_date"].max()
    previous_dates = history_df.loc[
        history_df["snapshot_date"].lt(latest_market_date),
        "snapshot_date",
    ]
    previous_market_date = previous_dates.max() if not previous_dates.empty else ""
    latest_df = history_df[history_df["snapshot_date"].eq(latest_market_date)].copy()
    latest_df["股票代號_key"] = latest_df["股票代號"].map(stock_code_key)
    latest_by_code = latest_df.set_index("股票代號_key")

    restore_columns = [
        "前次分數",
        "分數變化",
        "分數變化率",
        "Entry Score",
        "進場觀察分數",
        "進場判讀",
        "進場理由",
    ]
    restored = 0
    for index, row in result_df.iterrows():
        code = row["股票代號_key"]
        if code not in latest_by_code.index:
            continue
        history_row = latest_by_code.loc[code]
        if isinstance(history_row, pd.DataFrame):
            history_row = history_row.iloc[-1]
        current_close = pd.to_numeric(pd.Series([row.get("收盤價")]), errors="coerce").iloc[0]
        history_close = pd.to_numeric(pd.Series([history_row.get("收盤價")]), errors="coerce").iloc[0]
        current_score = pd.to_numeric(pd.Series([row.get("DeepTrend分數")]), errors="coerce").iloc[0]
        history_score = pd.to_numeric(pd.Series([history_row.get("DeepTrend分數")]), errors="coerce").iloc[0]
        if pd.isna(current_close) or pd.isna(history_close) or abs(current_close - history_close) > 0.0001:
            continue
        if pd.isna(current_score) or pd.isna(history_score) or abs(current_score - history_score) > 0.0001:
            continue
        for column in restore_columns:
            if column in result_df.columns and column in history_row.index:
                result_df.at[index, column] = history_row[column]
        restored += 1

    result_df["資料日期"] = latest_market_date
    result_df["前次資料日期"] = previous_market_date
    if CHIP_DAILY_FILE.exists():
        chip_dates = pd.to_datetime(pd.read_csv(CHIP_DAILY_FILE, usecols=["date"])["date"], errors="coerce")
        if chip_dates.notna().any():
            result_df["籌碼資料日期"] = chip_dates.max().strftime("%Y-%m-%d")
    result_df = result_df.drop(columns=["股票代號_key"], errors="ignore")
    result_df.to_excel(RESULT_FILE, index=False)
    print(f"Current result restored from {latest_market_date}: {restored}/{len(result_df)} rows")


def repair_factor_history(invalid_dates):
    if not FACTOR_FILE.exists():
        return
    factor_df = pd.read_csv(FACTOR_FILE, low_memory=False)
    if "event_date" not in factor_df.columns:
        return
    before_rows = len(factor_df)
    event_dates = pd.to_datetime(factor_df["event_date"], errors="coerce")
    invalid_mask = event_dates.dt.dayofweek.ge(5) | event_dates.dt.strftime("%Y-%m-%d").isin(invalid_dates)
    if not invalid_mask.any():
        print("Factor history unchanged: no invalid events found")
        return
    factor_df = factor_df.loc[~invalid_mask].copy()
    factor_df.to_csv(FACTOR_FILE, index=False, encoding="utf-8-sig")
    print(f"Factor history repaired: {before_rows} -> {len(factor_df)} rows")


def repair_universe_result(previous_ref, market_date, previous_market_date):
    if not previous_ref or not UNIVERSE_RESULT_FILE.exists():
        return

    from main import calculate_score_change, score_entry_position

    previous_bytes = subprocess.check_output(
        ["git", "show", f"{previous_ref}:output/universe_analysis_result.xlsx"],
        cwd=BASE_DIR,
    )
    previous_df = pd.read_excel(io.BytesIO(previous_bytes))
    current_df = pd.read_excel(UNIVERSE_RESULT_FILE)
    previous_df["股票代號_key"] = previous_df["股票代號"].map(stock_code_key)
    current_df["股票代號_key"] = current_df["股票代號"].map(stock_code_key)
    previous_by_code = previous_df.set_index("股票代號_key")

    matched = 0
    for index, row in current_df.iterrows():
        code = row["股票代號_key"]
        if code not in previous_by_code.index:
            continue
        previous_row = previous_by_code.loc[code]
        if isinstance(previous_row, pd.DataFrame):
            previous_row = previous_row.iloc[-1]
        previous_score = pd.to_numeric(pd.Series([previous_row.get("DeepTrend分數")]), errors="coerce").iloc[0]
        current_score = pd.to_numeric(pd.Series([row.get("DeepTrend分數")]), errors="coerce").iloc[0]
        score_change, score_change_rate = calculate_score_change(current_score, previous_score)
        entry_score, entry_judgement, entry_reasons = score_entry_position(
            row.get("收盤價"),
            row.get("5日線"),
            row.get("10日線"),
            row.get("20日線"),
            row.get("成交量"),
            row.get("5日均量"),
            row.get("20日高點"),
            row.get("20日低點"),
            current_score,
            score_change,
        )
        current_df.at[index, "前次分數"] = previous_score
        current_df.at[index, "前次資料日期"] = previous_market_date
        current_df.at[index, "分數變化"] = score_change
        current_df.at[index, "分數變化率"] = score_change_rate
        current_df.at[index, "Entry Score"] = entry_score
        current_df.at[index, "進場觀察分數"] = entry_score
        current_df.at[index, "進場判讀"] = entry_judgement
        current_df.at[index, "進場理由"] = entry_reasons
        matched += 1

    current_df["資料日期"] = market_date
    if CHIP_DAILY_FILE.exists():
        chip_dates = pd.to_datetime(pd.read_csv(CHIP_DAILY_FILE, usecols=["date"])["date"], errors="coerce")
        if chip_dates.notna().any():
            current_df["籌碼資料日期"] = chip_dates.max().strftime("%Y-%m-%d")
    current_df = current_df.drop(columns=["股票代號_key"], errors="ignore")
    current_df.to_excel(UNIVERSE_RESULT_FILE, index=False)
    print(f"Universe result restored: {matched}/{len(current_df)} rows")


def main():
    parser = argparse.ArgumentParser(description="Repair duplicate or non-trading DeepTrend snapshots.")
    parser.add_argument("invalid_dates", nargs="*", help="Additional YYYY-MM-DD market holidays to remove.")
    parser.add_argument("--previous-universe-ref", help="Git ref containing the preceding trading-day universe result.")
    parser.add_argument("--market-date", help="Current result trading date in YYYY-MM-DD format.")
    parser.add_argument("--previous-market-date", help="Preceding trading date in YYYY-MM-DD format.")
    args = parser.parse_args()
    history_df = repair_history(set(args.invalid_dates))
    restore_current_result(history_df)
    repair_factor_history(set(args.invalid_dates))
    repair_universe_result(args.previous_universe_ref, args.market_date, args.previous_market_date)


if __name__ == "__main__":
    main()
