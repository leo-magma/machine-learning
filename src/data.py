from __future__ import annotations

from functools import lru_cache
from io import StringIO

import pandas as pd
import yfinance as yf


PRICE_COLUMNS = ["Open", "High", "Low", "Close", "Adj Close", "Volume"]


class DataDownloadError(RuntimeError):
    """Raised when market data cannot be downloaded or normalized."""


@lru_cache(maxsize=32)
def _download_cached(ticker: str, start_date: str, end_date: str, interval: str) -> str:
    data = yf.download(
        ticker,
        start=start_date,
        end=end_date,
        interval=interval,
        progress=False,
        auto_adjust=False,
        group_by="column",
    )
    normalized = normalize_price_data(data, ticker)
    return normalized.to_json(date_format="iso")


def download_price_data(
    ticker: str,
    start_date: str,
    end_date: str,
    interval: str = "1d",
) -> pd.DataFrame:
    """Download and normalize OHLCV data from yfinance."""
    ticker = ticker.strip().upper()
    if not ticker:
        raise DataDownloadError("Please enter a ticker.")

    try:
        json_data = _download_cached(ticker, start_date, end_date, interval)
    except Exception as exc:  # noqa: BLE001 - surface readable errors in the UI
        if isinstance(exc, DataDownloadError):
            raise
        raise DataDownloadError(f"Failed to download data: {exc}") from exc

    data = pd.read_json(StringIO(json_data))
    data.index = pd.to_datetime(data.index)
    data.index.name = "Date"
    return data


def normalize_price_data(data: pd.DataFrame, ticker: str) -> pd.DataFrame:
    if data.empty:
        raise DataDownloadError(f"No price data was found for {ticker}.")

    if isinstance(data.columns, pd.MultiIndex):
        if ticker in data.columns.get_level_values(-1):
            data = data.xs(ticker, axis=1, level=-1)
        else:
            data.columns = data.columns.get_level_values(0)

    data = data.copy()
    data.index = pd.to_datetime(data.index)
    data.index.name = "Date"
    data = data.sort_index()

    for column in PRICE_COLUMNS:
        if column not in data.columns:
            if column == "Adj Close" and "Close" in data.columns:
                data[column] = data["Close"]
            else:
                raise DataDownloadError(f"Required column `{column}` is missing from the downloaded data.")

    data = data[PRICE_COLUMNS].apply(pd.to_numeric, errors="coerce")
    data = data.dropna(subset=["Open", "High", "Low", "Close", "Volume"])
    if data.empty:
        raise DataDownloadError("No valid price data remains. Please use a wider date range.")
    return data


def dataframe_to_records(data: pd.DataFrame) -> list[dict]:
    records = data.reset_index().copy()
    records["Date"] = records["Date"].dt.strftime("%Y-%m-%d")
    return records.to_dict("records")


def records_to_dataframe(records: list[dict]) -> pd.DataFrame:
    data = pd.DataFrame(records)
    if data.empty:
        return data
    data["Date"] = pd.to_datetime(data["Date"])
    data = data.set_index("Date").sort_index()
    return data

