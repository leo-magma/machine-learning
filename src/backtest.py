from __future__ import annotations

import numpy as np
import pandas as pd


TRADING_DAYS = 252


def run_backtest(
    predictions: pd.DataFrame,
    initial_capital: float,
    threshold: float,
    transaction_cost_bps: float,
    slippage_bps: float,
    strategy_mode: str = "long_cash",
) -> tuple[pd.DataFrame, dict[str, float]]:
    """Convert predicted returns into a configurable signal-based backtest."""
    if predictions.empty:
        raise ValueError("There are no predictions to backtest.")

    data = predictions.copy()
    total_cost = (transaction_cost_bps + slippage_bps) / 10_000

    data["signal"] = build_signal(data["predicted_return"], threshold, strategy_mode)
    data["position_change"] = data["signal"].diff().abs().fillna(data["signal"].abs())
    data["strategy_return"] = data["signal"] * data["actual_return"]
    data["strategy_return_net"] = data["strategy_return"] - data["position_change"] * total_cost
    data["buy_hold_return"] = data["actual_return"]

    data["strategy_equity"] = initial_capital * (1 + data["strategy_return_net"]).cumprod()
    data["buy_hold_equity"] = initial_capital * (1 + data["buy_hold_return"]).cumprod()
    data["strategy_drawdown"] = data["strategy_equity"] / data["strategy_equity"].cummax() - 1
    data["buy_hold_drawdown"] = data["buy_hold_equity"] / data["buy_hold_equity"].cummax() - 1

    metrics = calculate_backtest_metrics(data)
    return data, metrics


def build_signal(predicted_return: pd.Series, threshold: float, strategy_mode: str) -> pd.Series:
    threshold = abs(threshold)
    if strategy_mode == "long_short":
        return pd.Series(
            np.select(
                [predicted_return > threshold, predicted_return < -threshold],
                [1, -1],
                default=0,
            ),
            index=predicted_return.index,
        )
    if strategy_mode == "short_cash":
        return pd.Series(
            np.where(predicted_return < -threshold, -1, 0),
            index=predicted_return.index,
        )
    return pd.Series(
        np.where(predicted_return > threshold, 1, 0),
        index=predicted_return.index,
    )


def calculate_backtest_metrics(data: pd.DataFrame) -> dict[str, float]:
    strategy_returns = data["strategy_return_net"].dropna()
    buy_hold_returns = data["buy_hold_return"].dropna()
    years = max(len(data) / TRADING_DAYS, 1 / TRADING_DAYS)

    final_strategy = data["strategy_equity"].iloc[-1] / data["strategy_equity"].iloc[0] - 1
    final_buy_hold = data["buy_hold_equity"].iloc[-1] / data["buy_hold_equity"].iloc[0] - 1

    annual_return = (1 + final_strategy) ** (1 / years) - 1
    annual_volatility = strategy_returns.std() * np.sqrt(TRADING_DAYS)
    sharpe = annual_return / annual_volatility if annual_volatility else 0.0
    downside_returns = strategy_returns[strategy_returns < 0]
    downside_volatility = downside_returns.std() * np.sqrt(TRADING_DAYS)
    sortino = annual_return / downside_volatility if downside_volatility else 0.0
    max_drawdown = float(data["strategy_drawdown"].min())
    calmar = annual_return / abs(max_drawdown) if max_drawdown else 0.0
    var_95 = strategy_returns.quantile(0.05)
    cvar_95 = strategy_returns[strategy_returns <= var_95].mean()

    active_trades = int(data["position_change"].sum())
    wins = strategy_returns[strategy_returns > 0]
    losses = strategy_returns[strategy_returns < 0]
    active_days = strategy_returns[data["signal"] != 0]
    win_rate = len(wins) / len(active_days) if len(active_days) else 0.0
    exposure = len(active_days) / len(data) if len(data) else 0.0
    gross_profit = wins.sum()
    gross_loss = abs(losses.sum())
    profit_factor = gross_profit / gross_loss if gross_loss else 0.0
    payoff_ratio = wins.mean() / abs(losses.mean()) if len(wins) and len(losses) else 0.0

    return {
        "Strategy Total Return": float(final_strategy),
        "Buy & Hold Total Return": float(final_buy_hold),
        "Annual Return": float(annual_return),
        "Annual Volatility": float(annual_volatility),
        "Sharpe": float(sharpe),
        "Sortino": float(sortino),
        "Calmar": float(calmar),
        "Max Drawdown": max_drawdown,
        "VaR 95": float(var_95),
        "CVaR 95": float(cvar_95),
        "Win Rate": float(win_rate),
        "Profit Factor": float(profit_factor),
        "Payoff Ratio": float(payoff_ratio),
        "Exposure": float(exposure),
        "Long Days": float((data["signal"] > 0).sum()),
        "Short Days": float((data["signal"] < 0).sum()),
        "Trades": float(active_trades),
        "Best Period": float(strategy_returns.max()),
        "Worst Period": float(strategy_returns.min()),
    }


def backtest_metrics_to_frame(metrics: dict[str, float]) -> pd.DataFrame:
    return pd.DataFrame(
        [{"metric": metric, "value": value} for metric, value in metrics.items()]
    )
