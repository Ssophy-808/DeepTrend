from datetime import date, timedelta
from pathlib import Path
import time

import pandas as pd
import requests
import urllib3
from requests.exceptions import RequestException


BASE_DIR = Path(__file__).resolve().parent
WATCHLIST_FILE = BASE_DIR / "watchlist.csv"
UNIVERSE_FILE = BASE_DIR / "universe.csv"
CHIP_FILE = BASE_DIR / "chip.csv"
OUTPUT_DIR = BASE_DIR / "output"
CHIP_DAILY_FILE = OUTPUT_DIR / "chip_daily.csv"
LOOKBACK_CALENDAR_DAYS = 25
TARGET_TRADING_DAYS = 10

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


DAILY_COLUMNS = [
    "date",
    "ticker",
    "stock_name",
    "foreign_buy",
    "foreign_sell",
    "foreign_net",
    "foreign_dealer_buy",
    "foreign_dealer_sell",
    "foreign_dealer_net",
    "investment_buy",
    "investment_sell",
    "investment_net",
    "dealer_buy",
    "dealer_sell",
    "dealer_net",
    "dealer_self_buy",
    "dealer_self_sell",
    "dealer_self_net",
    "dealer_hedge_buy",
    "dealer_hedge_sell",
    "dealer_hedge_net",
    "total_net",
    "source",
]


def to_int(value):
    try:
        text = str(value).replace(",", "").strip()
        if text in ["", "-", "nan", "None"]:
            return 0
        return int(float(text))
    except (TypeError, ValueError):
        return 0


def to_roc_date(day):
    return f"{day.year - 1911}/{day.month:02d}/{day.day:02d}"


def normalize_chip_ticker(ticker):
    code = str(ticker).strip().split(".")[0]
    return f"{code}.TW"


def empty_chip():
    return {"total": 0, "foreign": 0, "investment": 0, "dealer": 0}


def empty_daily_row(day, ticker, stock_name="", source=""):
    row = {column: 0 for column in DAILY_COLUMNS}
    row.update(
        {
            "date": day.isoformat(),
            "ticker": normalize_chip_ticker(ticker),
            "stock_name": stock_name,
            "source": source,
        }
    )
    return row


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
        print(f"資料抓取失敗：{url}，原因：{exc}")
        return {}


def get_index(fields, *names):
    for name in names:
        if name in fields:
            return fields.index(name)
    return None


def get_value(fields, row, *names):
    index = get_index(fields, *names)
    if index is None or len(row) <= index:
        return 0
    return to_int(row[index])


def get_text(fields, row, *names):
    index = get_index(fields, *names)
    if index is None or len(row) <= index:
        return ""
    return str(row[index]).strip()


def fetch_twse_chip(day):
    url = "https://www.twse.com.tw/rwd/zh/fund/T86"
    params = {
        "date": day.strftime("%Y%m%d"),
        "selectType": "ALLBUT0999",
        "response": "json",
    }
    payload = get_json(url, params=params)

    if payload.get("stat") != "OK":
        return {}, []

    fields = payload.get("fields", [])
    code_index = get_index(fields, "證券代號")
    if code_index is None:
        return {}, []

    results = {}
    daily_rows = []
    for row in payload.get("data", []):
        if len(row) <= code_index:
            continue

        ticker = normalize_chip_ticker(row[code_index])
        daily = empty_daily_row(
            day,
            ticker,
            get_text(fields, row, "證券名稱"),
            "TWSE",
        )
        daily.update(
            {
                "foreign_buy": get_value(fields, row, "外陸資買進股數(不含外資自營商)"),
                "foreign_sell": get_value(fields, row, "外陸資賣出股數(不含外資自營商)"),
                "foreign_net": get_value(fields, row, "外陸資買賣超股數(不含外資自營商)"),
                "foreign_dealer_buy": get_value(fields, row, "外資自營商買進股數"),
                "foreign_dealer_sell": get_value(fields, row, "外資自營商賣出股數"),
                "foreign_dealer_net": get_value(fields, row, "外資自營商買賣超股數"),
                "investment_buy": get_value(fields, row, "投信買進股數"),
                "investment_sell": get_value(fields, row, "投信賣出股數"),
                "investment_net": get_value(fields, row, "投信買賣超股數"),
                "dealer_net": get_value(fields, row, "自營商買賣超股數"),
                "dealer_self_buy": get_value(fields, row, "自營商買進股數(自行買賣)"),
                "dealer_self_sell": get_value(fields, row, "自營商賣出股數(自行買賣)"),
                "dealer_self_net": get_value(fields, row, "自營商買賣超股數(自行買賣)"),
                "dealer_hedge_buy": get_value(fields, row, "自營商買進股數(避險)"),
                "dealer_hedge_sell": get_value(fields, row, "自營商賣出股數(避險)"),
                "dealer_hedge_net": get_value(fields, row, "自營商買賣超股數(避險)"),
                "total_net": get_value(fields, row, "三大法人買賣超股數"),
            }
        )
        daily["dealer_buy"] = daily["dealer_self_buy"] + daily["dealer_hedge_buy"]
        daily["dealer_sell"] = daily["dealer_self_sell"] + daily["dealer_hedge_sell"]

        if daily["total_net"] == 0:
            daily["total_net"] = (
                daily["foreign_net"]
                + daily["foreign_dealer_net"]
                + daily["investment_net"]
                + daily["dealer_net"]
            )

        results[ticker] = {
            "total": daily["total_net"],
            "foreign": daily["foreign_net"] + daily["foreign_dealer_net"],
            "investment": daily["investment_net"],
            "dealer": daily["dealer_net"],
        }
        daily_rows.append(daily)

    return results, daily_rows


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

    if not tables:
        return {}, []

    table = tables[0]
    fields = table.get("fields", [])
    results = {}
    daily_rows = []
    for row in table.get("data", []):
        if len(row) < 2:
            continue

        ticker = normalize_chip_ticker(row[0])
        daily = empty_daily_row(day, ticker, str(row[1]).strip(), "TPEX")

        if fields:
            daily.update(
                {
                    "foreign_buy": get_value(fields, row, "外資及陸資買進股數", "外陸資買進股數"),
                    "foreign_sell": get_value(fields, row, "外資及陸資賣出股數", "外陸資賣出股數"),
                    "foreign_net": get_value(fields, row, "外資及陸資買賣超股數", "外陸資買賣超股數"),
                    "investment_buy": get_value(fields, row, "投信買進股數"),
                    "investment_sell": get_value(fields, row, "投信賣出股數"),
                    "investment_net": get_value(fields, row, "投信買賣超股數"),
                    "dealer_self_buy": get_value(fields, row, "自營商自行買賣買進股數"),
                    "dealer_self_sell": get_value(fields, row, "自營商自行買賣賣出股數"),
                    "dealer_self_net": get_value(fields, row, "自營商自行買賣買賣超股數"),
                    "dealer_hedge_buy": get_value(fields, row, "自營商避險買進股數"),
                    "dealer_hedge_sell": get_value(fields, row, "自營商避險賣出股數"),
                    "dealer_hedge_net": get_value(fields, row, "自營商避險買賣超股數"),
                    "total_net": get_value(fields, row, "三大法人買賣超股數"),
                }
            )
        else:
            daily["foreign_net"] = to_int(row[10]) if len(row) > 10 else 0
            daily["investment_net"] = to_int(row[13]) if len(row) > 13 else 0
            daily["total_net"] = to_int(row[-1])

        daily["dealer_net"] = daily["dealer_self_net"] + daily["dealer_hedge_net"]
        if daily["dealer_net"] == 0:
            daily["dealer_net"] = daily["total_net"] - daily["foreign_net"] - daily["investment_net"]
        daily["dealer_buy"] = daily["dealer_self_buy"] + daily["dealer_hedge_buy"]
        daily["dealer_sell"] = daily["dealer_self_sell"] + daily["dealer_hedge_sell"]

        if daily["total_net"] == 0:
            daily["total_net"] = daily["foreign_net"] + daily["investment_net"] + daily["dealer_net"]

        results[ticker] = {
            "total": daily["total_net"],
            "foreign": daily["foreign_net"],
            "investment": daily["investment_net"],
            "dealer": daily["dealer_net"],
        }
        daily_rows.append(daily)

    return results, daily_rows


def fetch_recent_chip_days():
    chip_days = []
    today = date.today()

    for offset in range(LOOKBACK_CALENDAR_DAYS):
        day = today - timedelta(days=offset)

        if day.weekday() >= 5:
            continue

        day_chip = {}
        day_rows = []
        for fetcher in (fetch_twse_chip, fetch_tpex_chip):
            source_data, source_rows = fetcher(day)
            day_rows.extend(source_rows)
            for ticker, chip in source_data.items():
                merged = day_chip.setdefault(ticker, empty_chip())
                for key in merged:
                    merged[key] += chip.get(key, 0)

        if day_chip:
            chip_days.append({"date": day, "summary": day_chip, "rows": day_rows})

        if len(chip_days) >= TARGET_TRADING_DAYS:
            break

        time.sleep(0.2)

    return chip_days


def save_chip_daily(chip_days):
    OUTPUT_DIR.mkdir(exist_ok=True)
    new_rows = []
    for chip_day in chip_days:
        new_rows.extend(chip_day["rows"])

    if not new_rows:
        return 0

    new_df = pd.DataFrame(new_rows)
    new_df = new_df.reindex(columns=DAILY_COLUMNS)

    if CHIP_DAILY_FILE.exists():
        old_df = pd.read_csv(CHIP_DAILY_FILE)
        combined = pd.concat([old_df, new_df], ignore_index=True)
    else:
        combined = new_df

    combined = combined.reindex(columns=DAILY_COLUMNS)
    combined = combined.drop_duplicates(subset=["date", "ticker", "source"], keep="last")
    combined = combined.sort_values(["date", "ticker", "source"])
    combined.to_csv(CHIP_DAILY_FILE, index=False, encoding="utf-8-sig")
    return len(new_df)


def load_ticker_universe():
    ticker_frames = []
    for source_file in [WATCHLIST_FILE, UNIVERSE_FILE]:
        if not source_file.exists():
            continue
        source_df = pd.read_csv(source_file)
        if "ticker" not in source_df.columns:
            continue
        ticker_frames.append(source_df[["ticker"]])

    if not ticker_frames:
        raise FileNotFoundError("找不到 watchlist.csv 或 universe.csv 可更新籌碼")

    combined = pd.concat(ticker_frames, ignore_index=True)
    combined["ticker"] = combined["ticker"].astype(str).map(normalize_chip_ticker)
    return combined["ticker"].drop_duplicates().tolist()


def main():
    tickers = load_ticker_universe()
    chip_days = fetch_recent_chip_days()

    if not chip_days:
        raise RuntimeError("抓不到法人籌碼資料，未更新 chip.csv")

    rows = []
    for ticker in tickers:
        values = [chip_day["summary"].get(ticker, empty_chip()) for chip_day in chip_days]
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
    daily_count = save_chip_daily(chip_days)
    print(
        f"已更新 {len(chip_df)} 檔籌碼彙總，"
        f"最近 {len(chip_days)} 個交易日；"
        f"每日原始明細新增/覆蓋 {daily_count} 筆。"
    )


if __name__ == "__main__":
    main()
