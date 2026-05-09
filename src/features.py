from __future__ import annotations

import numpy as np
import pandas as pd


FEATURE_SETS = {
    "basic": "Basic Features",
    "technical": "Technical Indicators",
    "all": "All Features",
}

TARGETS = {
    "next_return": "Next Trading Day Return",
    "next_close": "Next Trading Day Close",
    "direction": "Next Trading Day Direction",
}


def build_feature_frame(
    prices: pd.DataFrame,
    feature_set: str = "all",
    target_type: str = "next_return",
) -> tuple[pd.DataFrame, list[str]]:
    """Create model-ready features and target without future leakage."""
    if prices.empty:
        raise ValueError("Price data is empty.")

    data = prices.copy()
    data["return_1d"] = data["Adj Close"].pct_change()
    data["log_return_1d"] = np.log(data["Adj Close"]).diff()
    data["volume_change"] = data["Volume"].pct_change()
    data["intraday_range"] = (data["High"] - data["Low"]) / data["Close"]
    data["close_open_return"] = (data["Close"] - data["Open"]) / data["Open"]

    for window in [3, 5, 10, 20]:
        data[f"return_lag_{window}d"] = data["return_1d"].shift(window)
        data[f"volume_lag_{window}d"] = data["volume_change"].shift(window)

    basic_features = [
        "return_1d",
        "log_return_1d",
        "volume_change",
        "intraday_range",
        "close_open_return",
        "return_lag_3d",
        "return_lag_5d",
        "return_lag_10d",
        "return_lag_20d",
        "volume_lag_3d",
        "volume_lag_5d",
        "volume_lag_10d",
        "volume_lag_20d",
    ]

    add_technical_features(data)
    technical_features = [
        "sma_5_gap",
        "sma_20_gap",
        "ema_12_gap",
        "ema_26_gap",
        "rsi_14",
        "macd",
        "macd_signal",
        "bb_position",
        "volatility_10",
        "volatility_20",
        "momentum_10",
        "momentum_20",
    ]

    if feature_set == "basic":
        feature_columns = basic_features
    elif feature_set == "technical":
        feature_columns = technical_features
    else:
        feature_columns = basic_features + technical_features

    data["target_next_return"] = data["Adj Close"].pct_change().shift(-1)
    data["target_next_close"] = data["Adj Close"].shift(-1)
    data["forward_return"] = data["target_next_return"]
    data["target_direction"] = np.where(
        data["target_next_return"].notna(),
        (data["target_next_return"] > 0).astype(int),
        np.nan,
    )

    target_column = {
        "next_return": "target_next_return",
        "next_close": "target_next_close",
        "direction": "target_direction",
    }.get(target_type)
    if target_column is None:
        raise ValueError(f"Unsupported prediction target: {target_type}")

    output_columns = PRICE_COLUMNS_FOR_CONTEXT + ["forward_return"] + feature_columns + [target_column]
    model_data = data[output_columns].replace([np.inf, -np.inf], np.nan).dropna()
    model_data = model_data.rename(columns={target_column: "target"})
    return model_data, feature_columns


def add_technical_features(data: pd.DataFrame) -> None:
    close = data["Adj Close"]

    sma_5 = close.rolling(5).mean()
    sma_20 = close.rolling(20).mean()
    data["sma_5_gap"] = close / sma_5 - 1
    data["sma_20_gap"] = close / sma_20 - 1

    ema_12 = close.ewm(span=12, adjust=False).mean()
    ema_26 = close.ewm(span=26, adjust=False).mean()
    data["ema_12_gap"] = close / ema_12 - 1
    data["ema_26_gap"] = close / ema_26 - 1

    delta = close.diff()
    gain = delta.clip(lower=0).rolling(14).mean()
    loss = (-delta.clip(upper=0)).rolling(14).mean()
    rs = gain / loss.replace(0, np.nan)
    data["rsi_14"] = 100 - (100 / (1 + rs))

    data["macd"] = ema_12 - ema_26
    data["macd_signal"] = data["macd"].ewm(span=9, adjust=False).mean()

    rolling_mean = close.rolling(20).mean()
    rolling_std = close.rolling(20).std()
    lower = rolling_mean - 2 * rolling_std
    upper = rolling_mean + 2 * rolling_std
    data["bb_position"] = (close - lower) / (upper - lower)

    data["volatility_10"] = data["return_1d"].rolling(10).std()
    data["volatility_20"] = data["return_1d"].rolling(20).std()
    data["momentum_10"] = close / close.shift(10) - 1
    data["momentum_20"] = close / close.shift(20) - 1


PRICE_COLUMNS_FOR_CONTEXT = ["Open", "High", "Low", "Close", "Adj Close", "Volume"]
