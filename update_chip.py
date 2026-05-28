from datetime import date, timedelta
from pathlib import Path
import time

import pandas as pd
import requests
import urllib3
from requests.exceptions import RequestException


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


def find_field(fields, keyword, exclude=None):
    exclude = exclude or []
    for index, field in enumerate(fields):
        if keyword in field and all(word not in field for word in exclude):
            return index
    return None


def empty_chip():
    return {"total": 0, "foreign": 0, "investment": 0, "dealer": 0}


def get_json(url, *, params=None):
    try:
        response = requests.get(
            url,
            params=params,
            timeout=15,
            verify=False,
            headers={"User-Agent": "Mozilla/5.0"},
        )
        response.raise_for_status()
        return response.json()
    except (RequestException, ValueError) as exc:
        print(f"略過籌碼來源：{url}，原因：{exc}")
        return {}


def fetch_twse_chip(day):
    url = "https://www.twse.com.tw/rwd/zh/fund/T86"
    params = {
        "date": day.strftime("%Y%m%d"),
        "selectType": "ALLBUT0999",
        "response": "json",
    }
    payload = get_json(url, params=params)

    if payload.get("stat") != "OK":
        return {}

    fields = payload.get("fields", [])
    code_index = fields.index("證券代號")
    total_index = fields.index("三大法人買賣超股數")
    foreign_index = find_field(fields, "外陸資買賣超股數", exclude=["不含"])
    foreign_ex_dealer_index = find_field(fields, "外陸資買賣超股數(不含外資自營商)")
    foreign_dealer_index = find_field(fields, "外資自營商買賣超股數")
    investment_index = find_field(fields, "投信買賣超股數")
    dealer_index = find_field(fields, "自營商買賣超股數")

    results = {}
    for row in payload.get("data", []):
        if len(row) <= max(code_index, total_index):
            continue

        total = to_int(row[total_index])
        investment = to_int(row[investment_index]) if investment_index is not None and len(row) > investment_index else 0
        if foreign_index is not None and len(row) > foreign_index:
            foreign = to_int(row[foreign_index])
        else:
            foreign = 0
            if foreign_ex_dealer_index is not None and len(row) > foreign_ex_dealer_index:
                foreign += to_int(row[foreign_ex_dealer_index])
            if foreign_dealer_index is not None and len(row) > foreign_dealer_index:
                foreign += to_int(row[foreign_dealer_index])

        dealer = total - foreign - investment

        results[normalize_chip_ticker(row[code_index])] = {
            "total": total,
            "foreign": foreign,
            "investment": investment,
            "dealer": dealer,
        }

    return results


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
    payload = get_json(url, params=params)
    tables = payload.get("tables", [])

    if not tables or payload.get("stat") == "很抱歉，沒有符合條件的資料!":
        return {}

    results = {}
    for row in tables[0].get("data", []):
        if len(row) < 2:
            continue
        total = to_int(row[-1])
        foreign = to_int(row[10]) if len(row) > 10 else 0
        investment = to_int(row[13]) if len(row) > 13 else 0
        dealer = total - foreign - investment

        results[normalize_chip_ticker(row[0])] = {
            "total": total,
            "foreign": foreign,
            "investment": investment,
            "dealer": dealer,
        }

    return results


def fetch_recent_chip_days():
    chip_days = []
    today = date.today()

    for offset in range(LOOKBACK_CALENDAR_DAYS):
        day = today - timedelta(days=offset)

        if day.weekday() >= 5:
            continue

        day_chip = {}
        for fetcher in (fetch_twse_chip, fetch_tpex_chip):
            source_data = fetcher(day)
            for ticker, chip in source_data.items():
                merged = day_chip.setdefault(ticker, empty_chip())
                for key in merged:
                    merged[key] += chip.get(key, 0)

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
        values = [day_chip.get(ticker, empty_chip()) for day_chip in chip_days]
        total_values = [value["total"] for value in values]
        foreign_values = [value["foreign"] for value in values]
        investment_values = [value["investment"] for value in values]
        dealer_values = [value["dealer"] for value in values]
        rows.append(
            {
                "ticker": ticker,
                "buy_sell_1d": sum(total_values[:1]),
                "buy_sell_3d": sum(total_values[:3]),
                "buy_sell_5d": sum(total_values[:5]),
                "buy_sell_10d": sum(total_values[:10]),
                "foreign_1d": sum(foreign_values[:1]),
                "foreign_3d": sum(foreign_values[:3]),
                "foreign_5d": sum(foreign_values[:5]),
                "foreign_10d": sum(foreign_values[:10]),
                "investment_1d": sum(investment_values[:1]),
                "investment_3d": sum(investment_values[:3]),
                "investment_5d": sum(investment_values[:5]),
                "investment_10d": sum(investment_values[:10]),
                "dealer_1d": sum(dealer_values[:1]),
                "dealer_3d": sum(dealer_values[:3]),
                "dealer_5d": sum(dealer_values[:5]),
                "dealer_10d": sum(dealer_values[:10]),
            }
        )

    chip_df = pd.DataFrame(rows).drop_duplicates(subset=["ticker"], keep="last")
    chip_df.to_csv(CHIP_FILE, index=False, encoding="utf-8")
    print(f"已更新 {len(chip_df)} 檔籌碼資料，使用最近 {len(chip_days)} 個交易日。")


if __name__ == "__main__":
    main()
