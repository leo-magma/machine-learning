from __future__ import annotations

from dataclasses import dataclass

import numpy as np
import pandas as pd


TRADING_DAYS = 252


@dataclass(frozen=True)
class PortfolioResult:
    weights: pd.Series
    expected_return: float
    volatility: float
    sharpe: float
    concentration: float


def returns_from_price_panel(price_panel: pd.DataFrame) -> pd.DataFrame:
    if price_panel.empty:
        return pd.DataFrame()
    return price_panel.sort_index().pct_change().replace([np.inf, -np.inf], np.nan).dropna(how="all")


def annualized_return(returns: pd.Series | pd.DataFrame) -> pd.Series | float:
    compounded = (1 + returns).prod()
    years = max(len(returns) / TRADING_DAYS, 1 / TRADING_DAYS)
    result = compounded ** (1 / years) - 1
    if isinstance(result, pd.Series):
        return result
    return float(result)


def annualized_covariance(returns: pd.DataFrame) -> pd.DataFrame:
    return returns.cov() * TRADING_DAYS


def portfolio_return(weights: pd.Series, expected_returns: pd.Series) -> float:
    aligned = expected_returns.reindex(weights.index).fillna(0)
    return float(np.dot(weights, aligned))


def portfolio_volatility(weights: pd.Series, covariance: pd.DataFrame) -> float:
    cov = covariance.reindex(index=weights.index, columns=weights.index).fillna(0)
    variance = float(weights.to_numpy().T @ cov.to_numpy() @ weights.to_numpy())
    return float(np.sqrt(max(variance, 0)))


def portfolio_sharpe(
    weights: pd.Series,
    expected_returns: pd.Series,
    covariance: pd.DataFrame,
    risk_free_rate: float = 0.0,
) -> float:
    ret = portfolio_return(weights, expected_returns)
    vol = portfolio_volatility(weights, covariance)
    return (ret - risk_free_rate) / vol if vol else 0.0


def concentration_score(weights: pd.Series) -> float:
    return float((weights**2).sum())


def equal_weight_portfolio(assets: list[str]) -> pd.Series:
    if not assets:
        return pd.Series(dtype=float)
    return pd.Series(1 / len(assets), index=assets)


def inverse_volatility_portfolio(returns: pd.DataFrame) -> pd.Series:
    vol = returns.std().replace(0, np.nan)
    inv = 1 / vol
    weights = inv / inv.sum()
    return weights.fillna(0)


def minimum_variance_portfolio(
    covariance: pd.DataFrame,
    long_only: bool = True,
) -> pd.Series:
    assets = covariance.columns
    cov = covariance.fillna(0).to_numpy()
    inv_cov = np.linalg.pinv(cov)
    ones = np.ones(len(assets))
    raw = inv_cov @ ones
    denominator = ones.T @ inv_cov @ ones
    if denominator == 0:
        return equal_weight_portfolio(list(assets))
    weights = raw / denominator
    result = pd.Series(weights, index=assets)
    if long_only:
        result = result.clip(lower=0)
        if result.sum() == 0:
            return equal_weight_portfolio(list(assets))
        result = result / result.sum()
    return result


def max_sharpe_grid_search(
    expected_returns: pd.Series,
    covariance: pd.DataFrame,
    risk_free_rate: float = 0.0,
    steps: int = 5000,
    seed: int = 42,
) -> PortfolioResult:
    rng = np.random.default_rng(seed)
    assets = list(expected_returns.index)
    if not assets:
        empty = pd.Series(dtype=float)
        return PortfolioResult(empty, 0.0, 0.0, 0.0, 0.0)

    best_weights = equal_weight_portfolio(assets)
    best_sharpe = portfolio_sharpe(best_weights, expected_returns, covariance, risk_free_rate)
    for _ in range(steps):
        raw = rng.random(len(assets))
        weights = pd.Series(raw / raw.sum(), index=assets)
        sharpe = portfolio_sharpe(weights, expected_returns, covariance, risk_free_rate)
        if sharpe > best_sharpe:
            best_sharpe = sharpe
            best_weights = weights

    return summarize_portfolio(best_weights, expected_returns, covariance, risk_free_rate)


def efficient_frontier(
    expected_returns: pd.Series,
    covariance: pd.DataFrame,
    portfolios: int = 2000,
    seed: int = 42,
) -> pd.DataFrame:
    rng = np.random.default_rng(seed)
    assets = list(expected_returns.index)
    rows = []
    for _ in range(portfolios):
        raw = rng.random(len(assets))
        weights = pd.Series(raw / raw.sum(), index=assets)
        rows.append(
            {
                "expected_return": portfolio_return(weights, expected_returns),
                "volatility": portfolio_volatility(weights, covariance),
                "sharpe": portfolio_sharpe(weights, expected_returns, covariance),
                "concentration": concentration_score(weights),
                **{f"weight_{asset}": weight for asset, weight in weights.items()},
            }
        )
    return pd.DataFrame(rows)


def summarize_portfolio(
    weights: pd.Series,
    expected_returns: pd.Series,
    covariance: pd.DataFrame,
    risk_free_rate: float = 0.0,
) -> PortfolioResult:
    return PortfolioResult(
        weights=weights,
        expected_return=portfolio_return(weights, expected_returns),
        volatility=portfolio_volatility(weights, covariance),
        sharpe=portfolio_sharpe(weights, expected_returns, covariance, risk_free_rate),
        concentration=concentration_score(weights),
    )


def risk_contribution(weights: pd.Series, covariance: pd.DataFrame) -> pd.Series:
    cov = covariance.reindex(index=weights.index, columns=weights.index).fillna(0)
    marginal = cov @ weights
    total_variance = float(weights.T @ cov @ weights)
    if total_variance == 0:
        return pd.Series(0.0, index=weights.index)
    contribution = weights * marginal / total_variance
    return contribution.sort_values(ascending=False)


def correlation_clusters(returns: pd.DataFrame, threshold: float = 0.65) -> dict[str, list[str]]:
    corr = returns.corr().fillna(0)
    remaining = set(corr.columns)
    clusters: dict[str, list[str]] = {}
    cluster_id = 1
    while remaining:
        asset = sorted(remaining)[0]
        peers = set(corr.index[corr[asset].abs() >= threshold])
        group = sorted(peers & remaining)
        clusters[f"Cluster {cluster_id}"] = group
        remaining -= set(group)
        cluster_id += 1
    return clusters


def portfolio_report_to_frame(result: PortfolioResult) -> pd.DataFrame:
    summary = pd.DataFrame(
        [
            ("Expected Return", result.expected_return),
            ("Volatility", result.volatility),
            ("Sharpe", result.sharpe),
            ("Concentration", result.concentration),
        ],
        columns=["metric", "value"],
    )
    weights = result.weights.reset_index()
    weights.columns = ["metric", "value"]
    weights["metric"] = "Weight: " + weights["metric"].astype(str)
    return pd.concat([summary, weights], ignore_index=True)


def rebalance_drift(
    target_weights: pd.Series,
    current_weights: pd.Series,
    tolerance: float = 0.03,
) -> pd.DataFrame:
    frame = pd.concat(
        [
            target_weights.rename("target"),
            current_weights.rename("current"),
        ],
        axis=1,
    ).fillna(0)
    frame["drift"] = frame["current"] - frame["target"]
    frame["action"] = np.where(
        frame["drift"].abs() > tolerance,
        np.where(frame["drift"] > 0, "Trim", "Add"),
        "Hold",
    )
    return frame.sort_values("drift", key=lambda series: series.abs(), ascending=False)


def turnover_between_weights(previous: pd.Series, target: pd.Series) -> float:
    aligned = pd.concat([previous.rename("previous"), target.rename("target")], axis=1).fillna(0)
    return float((aligned["target"] - aligned["previous"]).abs().sum() / 2)
