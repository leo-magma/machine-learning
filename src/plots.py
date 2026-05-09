from __future__ import annotations

import numpy as np
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots


EMPTY_TEMPLATE = "plotly_dark"
PAPER_BG = "#020805"
PLOT_BG = "#04120a"
TEXT = "#d9ffe5"
GREEN = "#00ff88"
SOFT_GREEN = "#5dffb1"
RED = "#ff4d6d"


def apply_dark_layout(fig: go.Figure, title: str | None = None, height: int | None = None) -> go.Figure:
    fig.update_layout(
        template=EMPTY_TEMPLATE,
        title=title,
        height=height,
        paper_bgcolor=PAPER_BG,
        plot_bgcolor=PLOT_BG,
        font={"color": TEXT},
        title_font={"color": GREEN},
        legend={"font": {"color": TEXT}},
        margin={"l": 50, "r": 30, "t": 60, "b": 50},
    )
    fig.update_xaxes(gridcolor="#12351f", zerolinecolor="#1f6f43")
    fig.update_yaxes(gridcolor="#12351f", zerolinecolor="#1f6f43")
    return fig


def empty_figure(message: str = "No data available") -> go.Figure:
    fig = go.Figure()
    fig.add_annotation(text=message, x=0.5, y=0.5, showarrow=False, font={"color": TEXT})
    return apply_dark_layout(fig, height=360)


def price_chart(prices: pd.DataFrame) -> go.Figure:
    if prices.empty:
        return empty_figure()
    fig = go.Figure(
        data=[
            go.Candlestick(
                x=prices.index,
                open=prices["Open"],
                high=prices["High"],
                low=prices["Low"],
                close=prices["Close"],
                name="Price",
                increasing_line_color=GREEN,
                decreasing_line_color=RED,
            )
        ]
    )
    apply_dark_layout(fig, "Price Chart")
    fig.update_layout(xaxis_rangeslider_visible=False)
    return fig


def feature_chart(model_data: pd.DataFrame, feature_columns: list[str]) -> go.Figure:
    if model_data.empty or not feature_columns:
        return empty_figure()
    selected = feature_columns[:5]
    normalized = model_data[selected].copy()
    normalized = (normalized - normalized.mean()) / normalized.std(ddof=0)
    fig = px.line(normalized, x=normalized.index, y=selected, title="Feature Trends (Standardized)")
    apply_dark_layout(fig)
    fig.update_layout(legend_title_text="Feature")
    return fig


def feature_correlation_heatmap(model_data: pd.DataFrame, feature_columns: list[str]) -> go.Figure:
    if model_data.empty or not feature_columns:
        return empty_figure()
    selected = feature_columns[:20] + ["target"]
    corr = model_data[selected].corr(numeric_only=True)
    fig = px.imshow(
        corr,
        color_continuous_scale="Greens",
        zmin=-1,
        zmax=1,
        title="Feature and Target Correlation Heatmap",
    )
    apply_dark_layout(fig, height=650)
    return fig


def target_distribution_chart(model_data: pd.DataFrame, target_type: str) -> go.Figure:
    if model_data.empty:
        return empty_figure()
    if target_type == "direction":
        counts = model_data["target"].map({0: "Down", 1: "Up"}).value_counts().reset_index()
        counts.columns = ["direction", "count"]
        fig = px.bar(counts, x="direction", y="count", title="Target Distribution")
    else:
        fig = px.histogram(model_data, x="target", nbins=60, title="Target Distribution")
    fig.update_traces(marker_color=GREEN)
    apply_dark_layout(fig)
    return fig


def prediction_chart(predictions: pd.DataFrame, target_type: str) -> go.Figure:
    if predictions.empty:
        return empty_figure()

    if "actual_return" in predictions.columns and "predicted_return" in predictions.columns:
        aligned = predictions[["actual_return", "predicted_return"]].astype(np.float64)
        aligned = aligned.replace([np.inf, -np.inf], np.nan).dropna(how="any")
        if aligned.empty:
            return empty_figure("No aligned return series for cumulative plot.")
        cum_actual = (1 + aligned["actual_return"]).cumprod()
        cum_pred = (1 + aligned["predicted_return"]).cumprod()
        cum_actual = cum_actual / cum_actual.iloc[0]
        cum_pred = cum_pred / cum_pred.iloc[0]
        idx = aligned.index
        fig = go.Figure()
        fig.add_trace(
            go.Scatter(
                x=idx,
                y=cum_actual,
                mode="lines",
                name="Actual (cumulative)",
                line={"color": SOFT_GREEN},
            )
        )
        fig.add_trace(
            go.Scatter(
                x=idx,
                y=cum_pred,
                mode="lines",
                name="Prediction (cumulative)",
                line={"color": "#f8ff7a"},
            )
        )
        title = "Cumulative growth: actual vs predicted returns (normalized to 1 at test start)"
        if target_type == "direction":
            title += " — direction model uses ±1% proxy returns per period"
        apply_dark_layout(fig, title)
        fig.update_layout(yaxis_title="Normalized wealth (compound)")
        return fig

    fig = go.Figure()
    fig.add_trace(
        go.Scatter(
            x=predictions.index,
            y=predictions["target"],
            mode="lines",
            name="Actual",
        )
    )
    fig.add_trace(
        go.Scatter(
            x=predictions.index,
            y=predictions["prediction"],
            mode="lines",
            name="Prediction",
        )
    )
    title = "Prediction vs Actual" if target_type != "direction" else "Direction Prediction vs Actual"
    fig.data[0].line.color = SOFT_GREEN
    fig.data[1].line.color = "#f8ff7a"
    apply_dark_layout(fig, title)
    return fig


def residual_chart(predictions: pd.DataFrame, target_type: str) -> go.Figure:
    if predictions.empty or target_type == "direction":
        return empty_figure("Residual plots are not shown for classification models")

    if "predicted_return" in predictions.columns and "actual_return" in predictions.columns:
        aligned = predictions[["predicted_return", "actual_return"]].astype(np.float64)
        aligned = aligned.replace([np.inf, -np.inf], np.nan).dropna(how="any")
        if aligned.empty:
            return empty_figure()
        diff = aligned["predicted_return"] - aligned["actual_return"]
        cum_diff = diff.cumsum()
        idx = aligned.index
        fig = make_subplots(specs=[[{"secondary_y": True}]])
        fig.add_trace(
            go.Scatter(
                x=idx,
                y=diff,
                mode="markers",
                name="Period error (pred − actual return)",
                marker={"color": GREEN, "size": 6},
            ),
            secondary_y=False,
        )
        fig.add_trace(
            go.Scatter(
                x=idx,
                y=cum_diff,
                mode="lines",
                name="Cumulative error",
                line={"color": "#42a5ff"},
            ),
            secondary_y=True,
        )
        fig.add_hline(y=0, secondary_y=False)
        apply_dark_layout(fig, "Return residuals and cumulative error")
        fig.update_yaxes(title_text="Per-period pred − actual", secondary_y=False)
        fig.update_yaxes(title_text="Cumulative sum", secondary_y=True)
        return fig

    residual = predictions["target"] - predictions["prediction"]
    fig = px.scatter(
        x=predictions.index,
        y=residual,
        labels={"x": "Date", "y": "Residual"},
        title="Residual Plot",
    )
    fig.add_hline(y=0)
    fig.update_traces(marker_color=GREEN)
    apply_dark_layout(fig)
    return fig


def feature_importance_chart(importance: pd.DataFrame) -> go.Figure:
    if importance.empty:
        return empty_figure("Feature importance is not available")
    plot_data = importance.copy()
    if "method" not in plot_data:
        plot_data["method"] = "importance"
    fig = px.bar(
        plot_data.sort_values("importance").tail(25),
        x="importance",
        y="feature",
        color="method",
        orientation="h",
        title="Feature Contribution: Native, Coefficient, and Permutation",
        color_discrete_sequence=[GREEN, "#f8ff7a", "#42a5ff", RED],
    )
    apply_dark_layout(fig)
    fig.update_layout(yaxis_title="")
    return fig


def model_comparison_chart(comparison: pd.DataFrame, primary_metric: str) -> go.Figure:
    if comparison.empty or primary_metric not in comparison.columns:
        return empty_figure("Model comparison data is not available")
    ascending = primary_metric in {"RMSE", "MAE"}
    ranked = comparison.sort_values(primary_metric, ascending=ascending)
    fig = px.bar(
        ranked,
        x=primary_metric,
        y="model",
        orientation="h",
        title=f"Model Comparison: {primary_metric}",
    )
    fig.update_traces(marker_color=GREEN)
    apply_dark_layout(fig)
    fig.update_layout(yaxis_title="")
    return fig


def backtest_equity_chart(backtest: pd.DataFrame) -> go.Figure:
    if backtest.empty:
        return empty_figure()
    fig = go.Figure()
    fig.add_trace(
        go.Scatter(
            x=backtest.index,
            y=backtest["strategy_equity"],
            mode="lines",
            name="Strategy",
        )
    )
    fig.add_trace(
        go.Scatter(
            x=backtest.index,
            y=backtest["buy_hold_equity"],
            mode="lines",
            name="All Buy (100% long)",
        )
    )
    fig.data[0].line.color = GREEN
    fig.data[1].line.color = "#f8ff7a"
    apply_dark_layout(fig, "Equity: Strategy vs All Buy (100% long)")
    return fig


def rolling_performance_chart(backtest: pd.DataFrame, window: int = 63) -> go.Figure:
    if backtest.empty:
        return empty_figure()
    rolling_return = (1 + backtest["strategy_return_net"]).rolling(window).apply(lambda x: x.prod() - 1)
    rolling_vol = backtest["strategy_return_net"].rolling(window).std()
    rolling_sharpe = rolling_return / rolling_vol.replace(0, np.nan)
    fig = go.Figure()
    fig.add_trace(
        go.Scatter(
            x=backtest.index,
            y=rolling_return,
            mode="lines",
            name=f"{window} Period Return",
            line={"color": GREEN},
        )
    )
    fig.add_trace(
        go.Scatter(
            x=backtest.index,
            y=rolling_sharpe,
            mode="lines",
            name=f"{window} Period Return / Vol",
            line={"color": "#f8ff7a"},
            yaxis="y2",
        )
    )
    apply_dark_layout(fig, "Rolling Strategy Diagnostics")
    fig.update_layout(
        yaxis={"title": "Rolling Return"},
        yaxis2={
            "title": "Rolling Return / Vol",
            "overlaying": "y",
            "side": "right",
            "gridcolor": "#12351f",
        },
    )
    return fig


def return_distribution_chart(backtest: pd.DataFrame) -> go.Figure:
    if backtest.empty:
        return empty_figure()
    returns = backtest[["strategy_return_net", "buy_hold_return"]].rename(
        columns={"strategy_return_net": "Strategy", "buy_hold_return": "Buy & Hold"}
    )
    fig = px.histogram(
        returns,
        x=["Strategy", "Buy & Hold"],
        nbins=50,
        barmode="overlay",
        title="Return Distribution",
    )
    fig.update_traces(opacity=0.65)
    apply_dark_layout(fig)
    return fig


def benchmark_comparison_chart(
    backtest: pd.DataFrame,
    benchmark_prices: pd.DataFrame,
    initial_capital: float,
    benchmark_label: str,
    instrument_ticker: str = "",
) -> go.Figure:
    if backtest.empty or benchmark_prices.empty:
        return empty_figure("Benchmark comparison is not available")
    benchmark = benchmark_prices[["Adj Close"]].copy()
    benchmark["benchmark_return"] = benchmark["Adj Close"].pct_change()
    aligned = backtest[["strategy_equity", "buy_hold_equity"]].join(
        benchmark["benchmark_return"],
        how="inner",
    )
    if aligned.empty:
        return empty_figure("Benchmark and strategy dates do not overlap")
    aligned["benchmark_equity"] = initial_capital * (1 + aligned["benchmark_return"].fillna(0)).cumprod()
    ins = (instrument_ticker or "Instrument").strip()
    all_buy_label = f"All Buy (100% long): {ins}"
    fig = go.Figure()
    fig.add_trace(
        go.Scatter(
            x=aligned.index,
            y=aligned["strategy_equity"],
            mode="lines",
            name="Strategy",
            line={"color": GREEN, "width": 2.4},
        )
    )
    fig.add_trace(
        go.Scatter(
            x=aligned.index,
            y=aligned["buy_hold_equity"],
            mode="lines",
            name=all_buy_label,
            line={"color": "#f8ff7a", "width": 2},
        )
    )
    fig.add_trace(
        go.Scatter(
            x=aligned.index,
            y=aligned["benchmark_equity"],
            mode="lines",
            name=f"Benchmark ({benchmark_label})",
            line={"color": "#42a5ff", "width": 1.8},
        )
    )
    apply_dark_layout(fig, "Strategy vs All Buy vs benchmark index")
    return fig


def risk_regime_chart(backtest: pd.DataFrame, window: int = 63) -> go.Figure:
    if backtest.empty:
        return empty_figure()
    rolling_vol = backtest["strategy_return_net"].rolling(window).std() * np.sqrt(252)
    rolling_dd = backtest["strategy_drawdown"]
    fig = go.Figure()
    fig.add_trace(
        go.Scatter(
            x=backtest.index,
            y=rolling_vol,
            mode="lines",
            name="Rolling Volatility",
            line={"color": "#42a5ff"},
        )
    )
    fig.add_trace(
        go.Scatter(
            x=backtest.index,
            y=rolling_dd,
            mode="lines",
            name="Drawdown",
            line={"color": RED},
            yaxis="y2",
        )
    )
    apply_dark_layout(fig, "Risk Regime Monitor")
    fig.update_layout(
        yaxis={"title": "Rolling Volatility"},
        yaxis2={
            "title": "Drawdown",
            "overlaying": "y",
            "side": "right",
            "gridcolor": "#12351f",
        },
    )
    return fig


def stress_scenario_chart(metrics: dict[str, float]) -> go.Figure:
    rows = [
        ("Worst Period", metrics.get("Worst Period", 0.0)),
        ("VaR 95", metrics.get("VaR 95", 0.0)),
        ("CVaR 95", metrics.get("CVaR 95", 0.0)),
        ("Max Drawdown", metrics.get("Max Drawdown", 0.0)),
    ]
    frame = pd.DataFrame(rows, columns=["scenario", "impact"])
    fig = px.bar(
        frame,
        x="impact",
        y="scenario",
        orientation="h",
        title="Stress and Tail-Risk Snapshot",
    )
    fig.update_traces(marker_color=RED)
    apply_dark_layout(fig)
    fig.update_layout(yaxis_title="")
    return fig


def drawdown_chart(backtest: pd.DataFrame) -> go.Figure:
    if backtest.empty:
        return empty_figure()
    fig = go.Figure()
    fig.add_trace(
        go.Scatter(
            x=backtest.index,
            y=backtest["strategy_drawdown"],
            mode="lines",
            name="Strategy Drawdown",
        )
    )
    fig.add_trace(
        go.Scatter(
            x=backtest.index,
            y=backtest["buy_hold_drawdown"],
            mode="lines",
            name="Buy & Hold Drawdown",
        )
    )
    fig.data[0].line.color = RED
    fig.data[1].line.color = "#f8ff7a"
    apply_dark_layout(fig, "Drawdown")
    return fig


def signal_chart(backtest: pd.DataFrame) -> go.Figure:
    if backtest.empty:
        return empty_figure()
    fig = go.Figure()
    fig.add_trace(
        go.Scatter(x=backtest.index, y=backtest["Close"], mode="lines", name="Close", line={"color": SOFT_GREEN})
    )
    buy_points = backtest[backtest["position_change"] > 0]
    fig.add_trace(
        go.Scatter(
            x=buy_points.index,
            y=buy_points["Close"],
            mode="markers",
            name="Signal Change",
            marker={"size": 8, "color": GREEN},
        )
    )
    apply_dark_layout(fig, "Trading Signals")
    return fig
