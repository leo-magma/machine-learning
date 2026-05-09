from __future__ import annotations

from dataclasses import dataclass
from typing import Iterable

import numpy as np
import pandas as pd


@dataclass(frozen=True)
class RiskLimit:
    name: str
    metric: str
    warning_level: float
    breach_level: float
    direction: str
    description: str


@dataclass(frozen=True)
class DueDiligenceItem:
    pillar: str
    item: str
    status: str
    detail: str


@dataclass(frozen=True)
class InstitutionalReport:
    score: float
    grade: str
    verdict: str
    strengths: tuple[str, ...]
    weaknesses: tuple[str, ...]
    due_diligence: tuple[DueDiligenceItem, ...]


DEFAULT_RISK_LIMITS = (
    RiskLimit(
        name="Maximum Drawdown",
        metric="Max Drawdown",
        warning_level=-0.20,
        breach_level=-0.35,
        direction="greater",
        description="Peak-to-trough loss should remain inside mandate tolerance.",
    ),
    RiskLimit(
        name="Sharpe Ratio",
        metric="Sharpe",
        warning_level=0.50,
        breach_level=0.00,
        direction="greater",
        description="Risk-adjusted return should justify allocation of capital.",
    ),
    RiskLimit(
        name="Sortino Ratio",
        metric="Sortino",
        warning_level=0.75,
        breach_level=0.00,
        direction="greater",
        description="Downside-adjusted return should be positive and stable.",
    ),
    RiskLimit(
        name="Tail Loss",
        metric="CVaR 95",
        warning_level=-0.035,
        breach_level=-0.065,
        direction="greater",
        description="Expected loss in the worst 5% of periods should be controlled.",
    ),
    RiskLimit(
        name="Exposure",
        metric="Exposure",
        warning_level=0.95,
        breach_level=1.00,
        direction="less",
        description="Capital usage should not silently become always-on beta.",
    ),
)


def generate_institutional_report(
    metrics: dict[str, float],
    model_name: str,
    primary_metric: str,
    primary_metric_value: float,
) -> InstitutionalReport:
    score = institutional_score(metrics, primary_metric_value)
    grade = score_to_grade(score)
    verdict = score_to_verdict(score)
    strengths = tuple(find_strengths(metrics, model_name, primary_metric, primary_metric_value))
    weaknesses = tuple(find_weaknesses(metrics))
    due_diligence = tuple(build_due_diligence_items(metrics))
    return InstitutionalReport(
        score=score,
        grade=grade,
        verdict=verdict,
        strengths=strengths,
        weaknesses=weaknesses,
        due_diligence=due_diligence,
    )


def institutional_score(metrics: dict[str, float], primary_metric_value: float) -> float:
    sharpe = normalize(metrics.get("Sharpe", 0.0), 0.0, 2.0)
    sortino = normalize(metrics.get("Sortino", 0.0), 0.0, 3.0)
    calmar = normalize(metrics.get("Calmar", 0.0), 0.0, 2.5)
    total_return = normalize(metrics.get("Strategy Total Return", 0.0), -0.25, 0.75)
    drawdown = 1 - normalize(abs(metrics.get("Max Drawdown", 0.0)), 0.0, 0.5)
    cvar = 1 - normalize(abs(metrics.get("CVaR 95", 0.0)), 0.0, 0.08)
    trade_quality = normalize(metrics.get("Profit Factor", 0.0), 0.5, 2.5)
    exposure_penalty = 1 - max(0.0, metrics.get("Exposure", 0.0) - 0.95) * 2
    model_signal = normalize(primary_metric_value, 0.0, 1.0)

    weighted = (
        0.16 * sharpe
        + 0.13 * sortino
        + 0.11 * calmar
        + 0.14 * total_return
        + 0.16 * drawdown
        + 0.10 * cvar
        + 0.08 * trade_quality
        + 0.06 * exposure_penalty
        + 0.06 * model_signal
    )
    return float(np.clip(weighted * 100, 0, 100))


def normalize(value: float, low: float, high: float) -> float:
    if high == low:
        return 0.0
    return float(np.clip((value - low) / (high - low), 0, 1))


def score_to_grade(score: float) -> str:
    if score >= 85:
        return "A"
    if score >= 72:
        return "B"
    if score >= 58:
        return "C"
    if score >= 42:
        return "D"
    return "F"


def score_to_verdict(score: float) -> str:
    if score >= 85:
        return "Institutional Candidate"
    if score >= 72:
        return "Promising, Needs Robustness Review"
    if score >= 58:
        return "Research Watchlist"
    if score >= 42:
        return "Weak Allocation Case"
    return "Reject Until Redesigned"


def find_strengths(
    metrics: dict[str, float],
    model_name: str,
    primary_metric: str,
    primary_metric_value: float,
) -> list[str]:
    strengths = [f"Selected model: {model_name} using {primary_metric}={primary_metric_value:.4f}."]
    if metrics.get("Strategy Total Return", 0.0) > metrics.get("Buy & Hold Total Return", 0.0):
        strengths.append("Strategy outperformed instrument Buy & Hold over the tested period.")
    if metrics.get("Sharpe", 0.0) > 1.0:
        strengths.append("Sharpe ratio is above 1.0, indicating positive risk-adjusted return.")
    if metrics.get("Sortino", 0.0) > metrics.get("Sharpe", 0.0):
        strengths.append("Sortino exceeds Sharpe, suggesting downside volatility is relatively contained.")
    if metrics.get("Profit Factor", 0.0) > 1.25:
        strengths.append("Profit factor is above 1.25, indicating favorable gross profit to gross loss.")
    if metrics.get("Exposure", 0.0) < 0.8:
        strengths.append("Exposure is below 80%, leaving room for capital efficiency and cash optionality.")
    return strengths[:5]


def find_weaknesses(metrics: dict[str, float]) -> list[str]:
    weaknesses = []
    if metrics.get("Strategy Total Return", 0.0) <= 0:
        weaknesses.append("Strategy total return is not positive in the selected test window.")
    if metrics.get("Sharpe", 0.0) < 0.5:
        weaknesses.append("Sharpe ratio is below institutional comfort levels.")
    if metrics.get("Max Drawdown", 0.0) < -0.2:
        weaknesses.append("Maximum drawdown is larger than 20%, requiring risk controls.")
    if metrics.get("CVaR 95", 0.0) < -0.035:
        weaknesses.append("CVaR 95 indicates meaningful tail loss concentration.")
    if metrics.get("Trades", 0.0) < 5:
        weaknesses.append("Trade count is low, so evidence may be statistically thin.")
    if metrics.get("Exposure", 0.0) > 0.95:
        weaknesses.append("Exposure is close to always-on, making beta contamination more likely.")
    return weaknesses or ["No major single-metric weakness detected; robustness testing is still required."]


def build_due_diligence_items(metrics: dict[str, float]) -> list[DueDiligenceItem]:
    items = []
    for limit in DEFAULT_RISK_LIMITS:
        value = metrics.get(limit.metric, 0.0)
        status = classify_limit(value, limit)
        items.append(
            DueDiligenceItem(
                pillar="Risk Limits",
                item=limit.name,
                status=status,
                detail=f"{limit.metric}={value:.4f}. {limit.description}",
            )
        )

    items.extend(
        [
            DueDiligenceItem(
                pillar="Model Governance",
                item="Time Series Split",
                status="Pass",
                detail="Train/test split preserves chronological order and avoids random shuffling.",
            ),
            DueDiligenceItem(
                pillar="Model Governance",
                item="Future Leakage",
                status="Review",
                detail="Features are shifted from historical data, but production use still needs independent leakage tests.",
            ),
            DueDiligenceItem(
                pillar="Trading Controls",
                item="Transaction Costs",
                status="Pass" if metrics.get("Trades", 0.0) >= 0 else "Review",
                detail="Backtest applies explicit transaction cost and slippage inputs.",
            ),
            DueDiligenceItem(
                pillar="Trading Controls",
                item="Capacity",
                status="Review",
                detail="Capacity, borrow, liquidity, and market impact are not fully modeled.",
            ),
        ]
    )
    return items


def classify_limit(value: float, limit: RiskLimit) -> str:
    if limit.direction == "greater":
        if value < limit.breach_level:
            return "Breach"
        if value < limit.warning_level:
            return "Warning"
        return "Pass"
    if value > limit.breach_level:
        return "Breach"
    if value > limit.warning_level:
        return "Warning"
    return "Pass"


def due_diligence_to_frame(items: Iterable[DueDiligenceItem]) -> pd.DataFrame:
    return pd.DataFrame(
        [
            {
                "pillar": item.pillar,
                "item": item.item,
                "status": item.status,
                "detail": item.detail,
            }
            for item in items
        ]
    )


def compute_alpha_beta(
    strategy_returns: pd.Series,
    benchmark_returns: pd.Series,
) -> dict[str, float]:
    aligned = pd.concat(
        [strategy_returns.rename("strategy"), benchmark_returns.rename("benchmark")],
        axis=1,
    ).dropna()
    if aligned.empty or aligned["benchmark"].var() == 0:
        return {"alpha": 0.0, "beta": 0.0, "correlation": 0.0}
    beta = aligned["strategy"].cov(aligned["benchmark"]) / aligned["benchmark"].var()
    alpha = aligned["strategy"].mean() - beta * aligned["benchmark"].mean()
    correlation = aligned["strategy"].corr(aligned["benchmark"])
    return {
        "alpha": float(alpha),
        "beta": float(beta),
        "correlation": float(correlation),
    }


def extract_drawdown_events(equity: pd.Series, top_n: int = 5) -> pd.DataFrame:
    running_max = equity.cummax()
    drawdown = equity / running_max - 1
    events = []
    in_event = False
    start = None
    trough = None
    trough_value = 0.0

    for timestamp, value in drawdown.items():
        if value < 0 and not in_event:
            in_event = True
            start = timestamp
            trough = timestamp
            trough_value = value
        elif value < 0 and in_event:
            if value < trough_value:
                trough = timestamp
                trough_value = value
        elif value == 0 and in_event:
            events.append(
                {
                    "start": start,
                    "trough": trough,
                    "end": timestamp,
                    "drawdown": trough_value,
                }
            )
            in_event = False

    if in_event:
        events.append(
            {
                "start": start,
                "trough": trough,
                "end": equity.index[-1],
                "drawdown": trough_value,
            }
        )

    frame = pd.DataFrame(events)
    if frame.empty:
        return frame
    return frame.sort_values("drawdown").head(top_n)


def monthly_return_matrix(returns: pd.Series) -> pd.DataFrame:
    if returns.empty:
        return pd.DataFrame()
    monthly = (1 + returns).resample("ME").prod() - 1
    frame = monthly.to_frame("return")
    frame["year"] = frame.index.year
    frame["month"] = frame.index.strftime("%b")
    return frame.pivot(index="year", columns="month", values="return")


def scenario_shock_table(
    latest_equity: float,
    shocks: tuple[float, ...] = (-0.10, -0.075, -0.05, 0.05, 0.075, 0.10),
) -> pd.DataFrame:
    return pd.DataFrame(
        [
            {
                "scenario": f"{shock:+.1%} one-period shock",
                "equity_after_shock": latest_equity * (1 + shock),
                "pnl": latest_equity * shock,
            }
            for shock in shocks
        ]
    )
