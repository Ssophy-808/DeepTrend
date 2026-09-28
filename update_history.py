from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

import pandas as pd


BASE_DIR = Path(__file__).resolve().parent
SOURCE_FILE = BASE_DIR / "output" / "stock_analysis_result.xlsx"
HISTORY_FILE = BASE_DIR / "output" / "stock_analysis_history.csv"
WATCHLIST_FILE = BASE_DIR / "watchlist.csv"
TIMEZONE = ZoneInfo("Asia/Taipei")
MIN_SNAPSHOT_SUCCESS_RATIO = 0.8


def comparable_snapshot(df):
    compare_columns = [
        col for col in df.columns if col not in {"snapshot_date", "snapshot_time", "source_file"}
    ]
    compare_df = df[compare_columns].copy()
    if "股票代號" in compare_df.columns:
        compare_df = compare_df.sort_values("股票代號")
    return compare_df.map(normalize_compare_value).reset_index(drop=True)


def snapshots_equal(left_df, right_df):
    left_compare = comparable_snapshot(left_df)
    right_compare = comparable_snapshot(right_df)
    if set(left_compare.columns) != set(right_compare.columns):
        return False
    right_compare = right_compare[left_compare.columns]
    if len(left_compare) != len(right_compare):
        return False
    return all(left_compare[col].equals(right_compare[col]) for col in left_compare.columns)


def normalize_compare_value(value):
    if pd.isna(value):
        return ""
    try:
        number = float(str(value).replace(",", ""))
    except (TypeError, ValueError):
        return str(value).strip()
    return f"{number:.6f}".rstrip("0").rstrip(".")


def load_current_snapshot():
    if not SOURCE_FILE.exists():
        raise FileNotFoundError(f"找不到來源檔案：{SOURCE_FILE}")

    snapshot_time = datetime.now(TIMEZONE)
    snapshot_df = pd.read_excel(SOURCE_FILE)

    if snapshot_df.empty:
        raise ValueError("stock_analysis_result.xlsx 是空的，未建立歷史快照")

    if WATCHLIST_FILE.exists():
        watchlist_df = pd.read_csv(WATCHLIST_FILE)
        min_required_rows = int(len(watchlist_df) * MIN_SNAPSHOT_SUCCESS_RATIO)
        if len(snapshot_df) > len(watchlist_df):
            raise ValueError(
                f"Snapshot has more rows than watchlist: {len(snapshot_df)}/{len(watchlist_df)} rows. "
                "History only accepts watchlist analysis, not universe analysis."
            )
        if {"ticker"}.issubset(watchlist_df.columns) and {"股票代號"}.issubset(snapshot_df.columns):
            watchlist_codes = set(watchlist_df["ticker"].astype(str).str.split(".").str[0])
            snapshot_codes = set(snapshot_df["股票代號"].astype(str).str.split(".").str[0])
            extra_codes = sorted(snapshot_codes - watchlist_codes)
            if extra_codes:
                raise ValueError(
                    "Snapshot contains stocks outside watchlist. "
                    f"History was not updated. Extra codes: {', '.join(extra_codes[:10])}"
                )
        if len(snapshot_df) < min_required_rows:
            raise ValueError(
                f"Snapshot is incomplete: {len(snapshot_df)}/{len(watchlist_df)} rows. "
                f"Need at least {min_required_rows}. History was not updated."
            )

    snapshot_date = snapshot_time.strftime("%Y-%m-%d")
    if "資料日期" in snapshot_df.columns:
        market_dates = pd.to_datetime(snapshot_df["資料日期"], errors="coerce").dropna()
        if not market_dates.empty:
            snapshot_date = market_dates.mode().iloc[0].strftime("%Y-%m-%d")

    snapshot_df.insert(0, "snapshot_date", snapshot_date)
    snapshot_df.insert(1, "snapshot_time", snapshot_time.strftime("%Y-%m-%d %H:%M:%S"))
    snapshot_df.insert(2, "source_file", "output/stock_analysis_result.xlsx")
    return snapshot_df


def append_snapshot_to_history(snapshot_df):
    HISTORY_FILE.parent.mkdir(parents=True, exist_ok=True)

    if HISTORY_FILE.exists():
        history_df = pd.read_csv(HISTORY_FILE)
        if not history_df.empty and "snapshot_date" in history_df.columns:
            snapshot_date = snapshot_df["snapshot_date"].iloc[0]
            same_date_df = history_df[history_df["snapshot_date"].astype(str).eq(snapshot_date)]
            if len(same_date_df) == len(snapshot_df) and snapshots_equal(same_date_df, snapshot_df):
                print(f"Snapshot unchanged for trading date {snapshot_date}; history was not updated.")
                return history_df
            history_df = history_df[~history_df["snapshot_date"].astype(str).eq(snapshot_date)]
            if not same_date_df.empty:
                print(f"Replaced existing snapshot for trading date {snapshot_date}.")
        combined_df = pd.concat([history_df, snapshot_df], ignore_index=True)
    else:
        combined_df = snapshot_df

    combined_df.to_csv(HISTORY_FILE, index=False, encoding="utf-8-sig")
    return combined_df


def main():
    snapshot_df = load_current_snapshot()
    combined_df = append_snapshot_to_history(snapshot_df)
    print(f"本次快照日期：{snapshot_df['snapshot_date'].iloc[0]}")
    print(f"本次寫入筆數：{len(snapshot_df)}")
    print(f"歷史資料總筆數：{len(combined_df)}")
    print(f"歷史資料檔案：{HISTORY_FILE}")


if __name__ == "__main__":
    main()
