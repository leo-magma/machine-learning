from __future__ import annotations

from dataclasses import dataclass

import numpy as np
import pandas as pd

from src.models import train_and_evaluate


@dataclass(frozen=True)
class WalkForwardWindow:
    train_start: pd.Timestamp
    train_end: pd.Timestamp
    test_start: pd.Timestamp
    test_end: pd.Timestamp
    train_rows: int
    test_rows: int


def build_walk_forward_windows(
    model_data: pd.DataFrame,
    train_size: int = 252,
    test_size: int = 63,
    step_size: int = 63,
) -> list[WalkForwardWindow]:
    if model_data.empty:
        return []
    windows = []
    start = 0
    while start + train_size + test_size <= len(model_data):
        train = model_data.iloc[start : start + train_size]
        test = model_data.iloc[start + train_size : start + train_size + test_size]
        windows.append(
            WalkForwardWindow(
                train_start=train.index.min(),
                train_end=train.index.max(),
                test_start=test.index.min(),
                test_end=test.index.max(),
                train_rows=len(train),
                test_rows=len(test),
            )
        )
        start += step_size
    return windows


def slice_window_data(model_data: pd.DataFrame, window: WalkForwardWindow) -> pd.DataFrame:
    return model_data.loc[window.train_start : window.test_end]


def window_train_ratio(window: WalkForwardWindow) -> float:
    total = window.train_rows + window.test_rows
    return window.train_rows / total if total else 0.75


def run_walk_forward_validation(
    model_data: pd.DataFrame,
    feature_columns: list[str],
    model_key: str,
    target_type: str,
    train_size: int = 252,
    test_size: int = 63,
    step_size: int = 63,
) -> tuple[pd.DataFrame, pd.DataFrame]:
    windows = build_walk_forward_windows(model_data, train_size, test_size, step_size)
    metrics_rows = []
    prediction_frames = []
    for idx, window in enumerate(windows, start=1):
        window_data = slice_window_data(model_data, window)
        result = train_and_evaluate(
            model_data=window_data,
            feature_columns=feature_columns,
            model_key=model_key,
            target_type=target_type,
            train_ratio=window_train_ratio(window),
        )
        row = {
            "window": idx,
            "train_start": window.train_start,
            "train_end": window.train_end,
            "test_start": window.test_start,
            "test_end": window.test_end,
            "train_rows": window.train_rows,
            "test_rows": window.test_rows,
            **result.metrics,
        }
        metrics_rows.append(row)
        predictions = result.predictions.copy()
        predictions["window"] = idx
        prediction_frames.append(predictions)
    metrics = pd.DataFrame(metrics_rows)
    predictions = pd.concat(prediction_frames) if prediction_frames else pd.DataFrame()
    return metrics, predictions


def summarize_walk_forward(metrics: pd.DataFrame) -> pd.DataFrame:
    if metrics.empty:
        return pd.DataFrame()
    numeric = metrics.select_dtypes(include=[np.number])
    rows = []
    for column in numeric.columns:
        if column in {"window", "train_rows", "test_rows"}:
            continue
        series = numeric[column].dropna()
        rows.append(
            {
                "metric": column,
                "mean": series.mean(),
                "median": series.median(),
                "std": series.std(),
                "min": series.min(),
                "max": series.max(),
                "positive_windows": (series > 0).sum(),
                "window_count": len(series),
            }
        )
    return pd.DataFrame(rows)


def walk_forward_stability_score(metrics: pd.DataFrame, primary_metric: str) -> float:
    if metrics.empty or primary_metric not in metrics:
        return 0.0
    series = metrics[primary_metric].dropna()
    if series.empty:
        return 0.0
    mean = series.mean()
    volatility = series.std()
    positive_rate = (series > 0).mean()
    stability = mean / (abs(volatility) + 1e-9)
    score = 100 * (0.55 * np.clip(stability / 2, 0, 1) + 0.45 * positive_rate)
    return float(score)


def detect_performance_decay(metrics: pd.DataFrame, primary_metric: str) -> dict[str, float | str]:
    if metrics.empty or primary_metric not in metrics or len(metrics) < 3:
        return {"status": "Insufficient Data", "slope": 0.0, "r2": 0.0}
    y = metrics[primary_metric].to_numpy()
    x = np.arange(len(y))
    x_design = np.column_stack([np.ones(len(x)), x])
    beta = np.linalg.pinv(x_design) @ y
    fitted = x_design @ beta
    ss_res = ((y - fitted) ** 2).sum()
    ss_tot = ((y - y.mean()) ** 2).sum()
    r2 = 1 - ss_res / ss_tot if ss_tot else 0.0
    slope = beta[1]
    if slope < 0 and r2 > 0.25:
        status = "Decay Risk"
    elif slope > 0 and r2 > 0.25:
        status = "Improving"
    else:
        status = "Stable / Noisy"
    return {"status": status, "slope": float(slope), "r2": float(r2)}


def embargo_split_indices(
    n_rows: int,
    train_size: int,
    test_size: int,
    embargo: int,
) -> list[tuple[np.ndarray, np.ndarray]]:
    splits = []
    start = 0
    while start + train_size + embargo + test_size <= n_rows:
        train_idx = np.arange(start, start + train_size)
        test_idx = np.arange(start + train_size + embargo, start + train_size + embargo + test_size)
        splits.append((train_idx, test_idx))
        start += test_size
    return splits


def purged_cross_validation_summary(
    model_data: pd.DataFrame,
    train_size: int = 252,
    test_size: int = 63,
    embargo: int = 5,
) -> pd.DataFrame:
    splits = embargo_split_indices(len(model_data), train_size, test_size, embargo)
    return pd.DataFrame(
        [
            {
                "split": idx,
                "train_start": model_data.index[train_idx[0]] if len(train_idx) else pd.NaT,
                "train_end": model_data.index[train_idx[-1]] if len(train_idx) else pd.NaT,
                "test_start": model_data.index[test_idx[0]] if len(test_idx) else pd.NaT,
                "test_end": model_data.index[test_idx[-1]] if len(test_idx) else pd.NaT,
                "embargo_rows": embargo,
                "train_rows": len(train_idx),
                "test_rows": len(test_idx),
            }
            for idx, (train_idx, test_idx) in enumerate(splits, start=1)
        ]
    )
