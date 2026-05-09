from __future__ import annotations

from dataclasses import dataclass

import numpy as np
import pandas as pd


@dataclass(frozen=True)
class ExecutionAssumption:
    spread_bps: float
    impact_coefficient: float
    participation_rate: float
    minimum_ticket_cost: float


DEFAULT_EXECUTION = ExecutionAssumption(
    spread_bps=2.0,
    impact_coefficient=8.0,
    participation_rate=0.10,
    minimum_ticket_cost=1.0,
)


def estimate_spread_cost(notional: pd.Series, spread_bps: float) -> pd.Series:
    return notional.abs() * spread_bps / 10_000


def estimate_market_impact(
    notional: pd.Series,
    average_dollar_volume: pd.Series,
    impact_coefficient: float,
) -> pd.Series:
    adv = average_dollar_volume.replace(0, np.nan)
    participation = (notional.abs() / adv).clip(lower=0).fillna(0)
    return notional.abs() * impact_coefficient * np.sqrt(participation) / 10_000


def estimate_slippage_cost(notional: pd.Series, slippage_bps: float) -> pd.Series:
    return notional.abs() * slippage_bps / 10_000


def estimate_ticket_cost(notional: pd.Series, minimum_ticket_cost: float) -> pd.Series:
    return pd.Series(
        np.where(notional.abs() > 0, minimum_ticket_cost, 0.0),
        index=notional.index,
    )


def build_order_schedule(
    target_position: pd.Series,
    current_position: pd.Series | None = None,
    max_participation: float = 0.10,
    average_volume: pd.Series | None = None,
) -> pd.DataFrame:
    target = target_position.fillna(0)
    current = current_position.reindex(target.index).fillna(0) if current_position is not None else pd.Series(0, index=target.index)
    order = target - current
    frame = pd.DataFrame({"current": current, "target": target, "order": order})
    if average_volume is not None:
        volume = average_volume.reindex(target.index).replace(0, np.nan)
        frame["participation"] = (frame["order"].abs() / volume).fillna(0)
        frame["clipped_order"] = np.sign(frame["order"]) * np.minimum(
            frame["order"].abs(),
            volume * max_participation,
        )
    else:
        frame["participation"] = 0.0
        frame["clipped_order"] = frame["order"]
    return frame


def execution_cost_report(
    orders: pd.DataFrame,
    price: pd.Series,
    volume: pd.Series,
    slippage_bps: float,
    assumptions: ExecutionAssumption = DEFAULT_EXECUTION,
) -> pd.DataFrame:
    aligned = orders.join(price.rename("price"), how="left").join(volume.rename("volume"), how="left")
    aligned["notional"] = aligned["clipped_order"] * aligned["price"]
    aligned["average_dollar_volume"] = aligned["price"] * aligned["volume"]
    aligned["spread_cost"] = estimate_spread_cost(aligned["notional"], assumptions.spread_bps)
    aligned["impact_cost"] = estimate_market_impact(
        aligned["notional"],
        aligned["average_dollar_volume"],
        assumptions.impact_coefficient,
    )
    aligned["slippage_cost"] = estimate_slippage_cost(aligned["notional"], slippage_bps)
    aligned["ticket_cost"] = estimate_ticket_cost(aligned["notional"], assumptions.minimum_ticket_cost)
    aligned["total_cost"] = (
        aligned["spread_cost"]
        + aligned["impact_cost"]
        + aligned["slippage_cost"]
        + aligned["ticket_cost"]
    )
    aligned["total_cost_bps"] = np.where(
        aligned["notional"].abs() > 0,
        aligned["total_cost"] / aligned["notional"].abs() * 10_000,
        0.0,
    )
    return aligned


def summarize_execution_costs(cost_report: pd.DataFrame) -> pd.DataFrame:
    if cost_report.empty:
        return pd.DataFrame()
    rows = [
        ("Total Notional", cost_report["notional"].abs().sum()),
        ("Total Cost", cost_report["total_cost"].sum()),
        ("Spread Cost", cost_report["spread_cost"].sum()),
        ("Impact Cost", cost_report["impact_cost"].sum()),
        ("Slippage Cost", cost_report["slippage_cost"].sum()),
        ("Ticket Cost", cost_report["ticket_cost"].sum()),
        ("Average Cost Bps", cost_report.loc[cost_report["notional"].abs() > 0, "total_cost_bps"].mean()),
        ("Max Participation", cost_report["participation"].max()),
    ]
    return pd.DataFrame(rows, columns=["metric", "value"])


def turnover_from_positions(positions: pd.DataFrame | pd.Series) -> pd.Series:
    if isinstance(positions, pd.Series):
        return positions.diff().abs().fillna(positions.abs())
    return positions.diff().abs().sum(axis=1).fillna(positions.abs().sum(axis=1))


def capacity_curve(
    signal_returns: pd.Series,
    average_dollar_volume: pd.Series,
    capital_levels: tuple[float, ...] = (100_000, 500_000, 1_000_000, 5_000_000, 10_000_000, 50_000_000),
    participation_cap: float = 0.10,
) -> pd.DataFrame:
    rows = []
    avg_adv = average_dollar_volume.mean()
    for capital in capital_levels:
        feasible = capital <= avg_adv * participation_cap if avg_adv else False
        impact_drag = np.sqrt(capital / avg_adv) * 0.001 if avg_adv else 0.0
        gross_return = signal_returns.mean() * 252
        net_return = gross_return - impact_drag * 252
        rows.append(
            {
                "capital": capital,
                "feasible": feasible,
                "gross_return": gross_return,
                "impact_drag": impact_drag,
                "net_return": net_return,
            }
        )
    return pd.DataFrame(rows)


def liquidity_bucket_report(
    returns: pd.Series,
    volume: pd.Series,
    price: pd.Series,
    buckets: int = 5,
) -> pd.DataFrame:
    frame = pd.concat(
        [
            returns.rename("return"),
            volume.rename("volume"),
            price.rename("price"),
        ],
        axis=1,
    ).dropna()
    if frame.empty:
        return pd.DataFrame()
    frame["dollar_volume"] = frame["volume"] * frame["price"]
    frame["liquidity_bucket"] = pd.qcut(frame["dollar_volume"], buckets, labels=False, duplicates="drop") + 1
    return (
        frame.groupby("liquidity_bucket")
        .agg(
            average_return=("return", "mean"),
            volatility=("return", "std"),
            average_dollar_volume=("dollar_volume", "mean"),
            observations=("return", "count"),
        )
        .reset_index()
    )


def implementation_shortfall(
    decision_price: pd.Series,
    execution_price: pd.Series,
    side: pd.Series,
    quantity: pd.Series,
) -> pd.DataFrame:
    frame = pd.concat(
        [
            decision_price.rename("decision_price"),
            execution_price.rename("execution_price"),
            side.rename("side"),
            quantity.rename("quantity"),
        ],
        axis=1,
    ).dropna()
    frame["shortfall_per_share"] = (frame["execution_price"] - frame["decision_price"]) * frame["side"]
    frame["shortfall"] = frame["shortfall_per_share"] * frame["quantity"].abs()
    frame["shortfall_bps"] = frame["shortfall_per_share"] / frame["decision_price"] * 10_000
    return frame


def execution_quality_score(cost_report: pd.DataFrame) -> float:
    if cost_report.empty:
        return 0.0
    avg_cost = cost_report.loc[cost_report["notional"].abs() > 0, "total_cost_bps"].mean()
    max_participation = cost_report["participation"].max()
    cost_score = 1 - min(max(avg_cost / 50, 0), 1) if pd.notna(avg_cost) else 0.0
    participation_score = 1 - min(max(max_participation / 0.25, 0), 1) if pd.notna(max_participation) else 0.0
    return float((0.7 * cost_score + 0.3 * participation_score) * 100)
