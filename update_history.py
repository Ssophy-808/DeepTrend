from datetime import datetime
from pathlib import Path
from zoneinfo import ZoneInfo

import pandas as pd


BASE_DIR = Path(__file__).resolve().parent
SOURCE_FILE = BASE_DIR / "output" / "stock_analysis_result.xlsx"
HISTORY_FILE = BASE_DIR / "output" / "stock_analysis_history.csv"
TIMEZONE = ZoneInfo("Asia/Taipei")


def load_current_snapshot():
    if not SOURCE_FILE.exists():
        raise FileNotFoundError(f"找不到來源檔案：{SOURCE_FILE}")

    snapshot_time = datetime.now(TIMEZONE)
    snapshot_df = pd.read_excel(SOURCE_FILE)

    if snapshot_df.empty:
        raise ValueError("stock_analysis_result.xlsx 是空的，未建立歷史快照")

    snapshot_df.insert(0, "snapshot_date", snapshot_time.strftime("%Y-%m-%d"))
    snapshot_df.insert(1, "snapshot_time", snapshot_time.strftime("%Y-%m-%d %H:%M:%S"))
    snapshot_df.insert(2, "source_file", "output/stock_analysis_result.xlsx")
    return snapshot_df


def append_snapshot_to_history(snapshot_df):
    HISTORY_FILE.parent.mkdir(parents=True, exist_ok=True)
    snapshot_date = str(snapshot_df["snapshot_date"].iloc[0])

    if HISTORY_FILE.exists():
        history_df = pd.read_csv(HISTORY_FILE)
        if "snapshot_date" in history_df.columns:
            history_df = history_df[history_df["snapshot_date"].astype(str) != snapshot_date]
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
