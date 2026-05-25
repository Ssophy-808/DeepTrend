from datetime import date, timedelta
from pathlib import Path
import time

import pandas as pd
import requests
import urllib3


BASE_DIR = Path(__file__).resolve().parent
WATCHLIST_FILE = BASE_DIR / "watchlist.csv"
CHIP_FILE = BASE_DIR / "chip.csv"
LOOKBACK_CALENDAR_DAYS = 25
TARGET_TRADING_DAYS = 10

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


def to_int(value):
    try:
        text = str(value).replace(",", "").strip()
        if text in ["", "-", "nan"]:
            return 0
        return int(float(text))
    except (TypeError, ValueError):
        return 0


def to_roc_date(day):
    return f"{day.year - 1911}/{day.month:02d}/{day.day:02d}"


def normalize_chip_ticker(ticker):
    code = str(ticker).strip().split(".")[0]
    return f"{code}.TW"


def fetch_twse_chip(day):
    url = "https://www.twse.com.tw/rwd/zh/fund/T86"
    params = {
        "date": day.strftime("%Y%m%d"),
        "selectType": "ALLBUT0999",
        "response": "json",
    }
    response = requests.get(url, params=params, timeout=15, verify=False, headers={"User-Agent": "Mozilla/5.0"})
    payload = response.json()

    if payload.get("stat") != "OK":
        return {}

    fields = payload.get("fields", [])
    code_index = fields.index("證券代號")
    total_index = fields.index("三大法人買賣超股數")

    return {
        normalize_chip_ticker(row[code_index]): to_int(row[total_index])
        for row in payload.get("data", [])
    }


def fetch_tpex_chip(day):
    url = "https://www.tpex.org.tw/web/stock/3insti/daily_trade/3itrade_hedge_result.php"
    params = {
        "l": "zh-tw",
        "o": "json",
        "se": "EW",
        "t": "D",
        "d": to_roc_date(day),
        "s": "0,asc",
    }
    response = requests.get(url, params=params, timeout=15, verify=False, headers={"User-Agent": "Mozilla/5.0"})
    payload = response.json()
    tables = payload.get("tables", [])

    if not tables or payload.get("stat") == "很抱歉，沒有符合條件的資料!":
        return {}

    return {
        normalize_chip_ticker(row[0]): to_int(row[-1])
        for row in tables[0].get("data", [])
    }


def fetch_recent_chip_days():
    chip_days = []
    today = date.today()

    for offset in range(LOOKBACK_CALENDAR_DAYS):
        day = today - timedelta(days=offset)

        if day.weekday() >= 5:
            continue

        day_chip = {}
        day_chip.update(fetch_twse_chip(day))
        day_chip.update(fetch_tpex_chip(day))

        if day_chip:
            chip_days.append(day_chip)

        if len(chip_days) >= TARGET_TRADING_DAYS:
            break

        time.sleep(0.2)

    return chip_days


def main():
    watchlist = pd.read_csv(WATCHLIST_FILE)
    tickers = [normalize_chip_ticker(ticker) for ticker in watchlist["ticker"].astype(str)]
    chip_days = fetch_recent_chip_days()

    if not chip_days:
        raise RuntimeError("抓不到三大法人籌碼資料，保留原本 chip.csv。")

    rows = []
    for ticker in tickers:
        values = [day_chip.get(ticker, 0) for day_chip in chip_days]
        rows.append(
            {
                "ticker": ticker,
                "buy_sell_1d": sum(values[:1]),
                "buy_sell_3d": sum(values[:3]),
                "buy_sell_5d": sum(values[:5]),
                "buy_sell_10d": sum(values[:10]),
            }
        )

    chip_df = pd.DataFrame(rows).drop_duplicates(subset=["ticker"], keep="last")
    chip_df.to_csv(CHIP_FILE, index=False, encoding="utf-8")
    print(f"已更新 {len(chip_df)} 檔籌碼資料，使用最近 {len(chip_days)} 個交易日。")


if __name__ == "__main__":
    main()
