from __future__ import annotations

from dataclasses import dataclass

import numpy as np
import pandas as pd


@dataclass(frozen=True)
class FactorResult:
    factor: str
    information_coefficient: float
    rank_ic: float
    t_stat: float
    hit_rate: float
    turnover: float


def build_factor_panel(model_data: pd.DataFrame, feature_columns: list[str]) -> pd.DataFrame:
    if model_data.empty:
        return pd.DataFrame()
    columns = [column for column in feature_columns if column in model_data.columns]
    panel = model_data[columns + ["target"]].copy()
    panel = panel.replace([np.inf, -np.inf], np.nan).dropna()
    return panel


def information_coefficient(feature: pd.Series, target: pd.Series) -> float:
    aligned = pd.concat([feature.rename("feature"), target.rename("target")], axis=1).dropna()
    if aligned.empty or aligned["feature"].std() == 0 or aligned["target"].std() == 0:
        return 0.0
    return float(aligned["feature"].corr(aligned["target"]))


def rank_information_coefficient(feature: pd.Series, target: pd.Series) -> float:
    aligned = pd.concat([feature.rename("feature"), target.rename("target")], axis=1).dropna()
    if aligned.empty:
        return 0.0
    return float(aligned["feature"].rank().corr(aligned["target"].rank()))


def factor_hit_rate(feature: pd.Series, target: pd.Series) -> float:
    aligned = pd.concat([feature.rename("feature"), target.rename("target")], axis=1).dropna()
    if aligned.empty:
        return 0.0
    signal = np.sign(aligned["feature"])
    outcome = np.sign(aligned["target"])
    active = signal != 0
    if active.sum() == 0:
        return 0.0
    return float((signal[active] == outcome[active]).mean())


def factor_turnover(feature: pd.Series) -> float:
    signal = np.sign(feature.fillna(0))
    return float(signal.diff().abs().fillna(0).mean())


def factor_t_stat(ic_series: pd.Series) -> float:
    series = ic_series.dropna()
    if len(series) < 2 or series.std() == 0:
        return 0.0
    return float(series.mean() / (series.std() / np.sqrt(len(series))))


def rolling_information_coefficient(
    panel: pd.DataFrame,
    factor: str,
    window: int = 63,
) -> pd.Series:
    if panel.empty or factor not in panel:
        return pd.Series(dtype=float)
    return panel[factor].rolling(window).corr(panel["target"])


def evaluate_factor(
    panel: pd.DataFrame,
    factor: str,
    rolling_window: int = 63,
) -> FactorResult:
    if panel.empty or factor not in panel:
        return FactorResult(factor, 0.0, 0.0, 0.0, 0.0, 0.0)
    rolling_ic = rolling_information_coefficient(panel, factor, rolling_window)
    return FactorResult(
        factor=factor,
        information_coefficient=information_coefficient(panel[factor], panel["target"]),
        rank_ic=rank_information_coefficient(panel[factor], panel["target"]),
        t_stat=factor_t_stat(rolling_ic),
        hit_rate=factor_hit_rate(panel[factor], panel["target"]),
        turnover=factor_turnover(panel[factor]),
    )


def evaluate_factor_library(
    model_data: pd.DataFrame,
    feature_columns: list[str],
    rolling_window: int = 63,
) -> pd.DataFrame:
    panel = build_factor_panel(model_data, feature_columns)
    rows = [
        evaluate_factor(panel, factor, rolling_window).__dict__
        for factor in feature_columns
        if factor in panel.columns
    ]
    if not rows:
        return pd.DataFrame()
    return pd.DataFrame(rows).sort_values("rank_ic", key=lambda series: series.abs(), ascending=False)


def factor_quantile_returns(
    panel: pd.DataFrame,
    factor: str,
    quantiles: int = 5,
) -> pd.DataFrame:
    if panel.empty or factor not in panel:
        return pd.DataFrame()
    frame = panel[[factor, "target"]].dropna().copy()
    unique_values = frame[factor].nunique()
    if unique_values < quantiles:
        return pd.DataFrame()
    frame["quantile"] = pd.qcut(frame[factor], quantiles, labels=False, duplicates="drop") + 1
    result = frame.groupby("quantile")["target"].agg(["mean", "std", "count"]).reset_index()
    result["sharpe_like"] = result["mean"] / result["std"].replace(0, np.nan)
    return result.fillna(0)


def factor_decay_profile(
    model_data: pd.DataFrame,
    factor: str,
    horizons: tuple[int, ...] = (1, 2, 3, 5, 10, 20),
) -> pd.DataFrame:
    if model_data.empty or factor not in model_data:
        return pd.DataFrame()
    rows = []
    close = model_data["Adj Close"] if "Adj Close" in model_data else None
    if close is None:
        return pd.DataFrame()
    for horizon in horizons:
        forward_return = close.pct_change(horizon).shift(-horizon)
        ic = information_coefficient(model_data[factor], forward_return)
        rank_ic = rank_information_coefficient(model_data[factor], forward_return)
        rows.append({"horizon": horizon, "ic": ic, "rank_ic": rank_ic})
    return pd.DataFrame(rows)


def orthogonalize_factor(
    target_factor: pd.Series,
    control_factors: pd.DataFrame,
) -> pd.Series:
    aligned = pd.concat([target_factor.rename("target_factor"), control_factors], axis=1).dropna()
    if aligned.empty or control_factors.empty:
        return target_factor
    y = aligned["target_factor"].to_numpy()
    x = aligned.drop(columns=["target_factor"]).to_numpy()
    x = np.column_stack([np.ones(len(x)), x])
    beta = np.linalg.pinv(x) @ y
    residual = y - x @ beta
    result = pd.Series(index=aligned.index, data=residual, name=f"{target_factor.name}_orthogonal")
    return result.reindex(target_factor.index)


def factor_correlation_clusters(
    panel: pd.DataFrame,
    feature_columns: list[str],
    threshold: float = 0.75,
) -> dict[str, list[str]]:
    available = [feature for feature in feature_columns if feature in panel]
    corr = panel[available].corr().fillna(0)
    remaining = set(available)
    clusters: dict[str, list[str]] = {}
    cluster_id = 1
    while remaining:
        feature = sorted(remaining)[0]
        peers = set(corr.index[corr[feature].abs() >= threshold])
        group = sorted(peers & remaining)
        clusters[f"Factor Cluster {cluster_id}"] = group
        remaining -= set(group)
        cluster_id += 1
    return clusters


def factor_capacity_proxy(
    model_data: pd.DataFrame,
    factor: str,
    volume_column: str = "Volume",
) -> pd.DataFrame:
    if model_data.empty or factor not in model_data or volume_column not in model_data:
        return pd.DataFrame()
    frame = model_data[[factor, volume_column, "target"]].dropna().copy()
    frame["abs_signal"] = frame[factor].abs()
    frame["liquidity_bucket"] = pd.qcut(frame[volume_column], 5, labels=False, duplicates="drop") + 1
    return (
        frame.groupby("liquidity_bucket")
        .agg(
            signal_strength=("abs_signal", "mean"),
            average_forward_return=("target", "mean"),
            observations=("target", "count"),
        )
        .reset_index()
    )


def factor_report(
    model_data: pd.DataFrame,
    feature_columns: list[str],
    top_n: int = 10,
) -> dict[str, pd.DataFrame]:
    library = evaluate_factor_library(model_data, feature_columns)
    if library.empty:
        return {"library": library}
    top_factors = library.head(top_n)["factor"].tolist()
    quantiles = []
    decays = []
    panel = build_factor_panel(model_data, feature_columns)
    for factor in top_factors:
        q = factor_quantile_returns(panel, factor)
        if not q.empty:
            q["factor"] = factor
            quantiles.append(q)
        d = factor_decay_profile(model_data, factor)
        if not d.empty:
            d["factor"] = factor
            decays.append(d)
    return {
        "library": library,
        "quantiles": pd.concat(quantiles, ignore_index=True) if quantiles else pd.DataFrame(),
        "decay": pd.concat(decays, ignore_index=True) if decays else pd.DataFrame(),
    }
