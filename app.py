from __future__ import annotations

import base64
import io
import re
import zipfile
from datetime import date, timedelta

import dash_bootstrap_components as dbc
import pandas as pd
from dash import Dash, Input, Output, State, callback_context, dash_table, dcc, html
from dash.exceptions import PreventUpdate

from src.backtest import backtest_metrics_to_frame, run_backtest
from src.data import DataDownloadError, download_price_data
from src.features import FEATURE_SETS, TARGETS, build_feature_frame
from src.factor_research import evaluate_factor_library
from src.institutional import due_diligence_to_frame, generate_institutional_report
from src.model_governance import checks_to_frame, governance_checklist
from src.models import available_models, metrics_to_frame, train_and_evaluate
from src.plots import (
    backtest_equity_chart,
    benchmark_comparison_chart,
    drawdown_chart,
    empty_figure,
    feature_chart,
    feature_correlation_heatmap,
    feature_importance_chart,
    model_comparison_chart,
    prediction_chart,
    price_chart,
    return_distribution_chart,
    residual_chart,
    risk_regime_chart,
    rolling_performance_chart,
    signal_chart,
    stress_scenario_chart,
    target_distribution_chart,
)


PRESET_TICKERS = [
    "AAPL",
    "MSFT",
    "GOOGL",
    "NVDA",
    "TSLA",
    "JPM",
    "GS",
    "SPY",
    "QQQ",
    "^GSPC",
    "BTC-USD",
    "ETH-USD",
    "7203.T",
    "6758.T",
    "9984.T",
]
INTERVALS = {
    "1d": "Daily",
    "1wk": "Weekly",
    "1mo": "Monthly",
}
STRATEGY_MODES = {
    "long_cash": "Long / Cash",
    "long_short": "Long / Short",
    "short_cash": "Short / Cash",
}
DEFAULT_END = date.today()
DEFAULT_START = DEFAULT_END - timedelta(days=365 * 10)
MIN_MODEL_ROWS = 60

# English copy for the Overview tab (Dash Markdown).
OVERVIEW_GUIDE_MARKDOWN = r"""
# What Poseidon Quant Lab is

**Poseidon Quant Lab** pulls market-style prices (via `yfinance`), engineers features, trains and compares machine‑learning models, turns forecasts into a simple signal‑based **paper portfolio**, and surfaces diagnostics across risk and governance.
It is **not investment advice**. Outputs are historical simulations only.

---

## Layout (sidebar vs main)

| Area | Role |
| ---- | ---- |
| **Left sidebar** | **Ticker** (preset dropdown), optional **symbol override**, benchmark, dates, target, features, model, costs. Press **Run Full Analysis**. |
| **Main** | Status banner, KPI summary cards, then charts and tables in **Dashboard** tabs. |
| **Overview · Guide** | This page—workflow and glossary. |
| **Dashboard** | Executive Review, Data Overview, Feature Analysis, Model Evaluation, Predictions, Backtest. |

---

## Quick start

1. Choose a **Ticker** from the preset list, or type another listing (e.g. `7203.T`) in **Symbol override**—override wins if it is not empty.
2. Set **Benchmark** (e.g. `SPY`)—shown compounded from **Initial Capital** next to strategy vs **All Buy**.
3. Set **Date Range** and **Interval** (`Daily` and multi‑year windows work best).
4. Adjust target, features, model, threshold and modes—or leave defaults.
5. Click **Run Full Analysis** and wait for data + training.
6. Inspect tabs left‑to‑right; download **ZIP** for CSV exports.

---

## Reading Strategy vs All Buy vs Benchmark

| Trace | Meaning |
| ----- | ------- |
| **Strategy** | Signal rule from forecasts + threshold + strategy mode. |
| **All Buy (100% long): TICKER** | Passive full long in the analyzed symbol. |
| **Benchmark (your symbol)** | Compounded path of the benchmark returns from **Initial Capital**. |

---

## Date range & interval

Coarser intervals reduce rows → errors below ~60 usable rows after warmup. Widen dates, use **Daily**, or pick **Basic** features.

---

## Prediction targets

| Key | Label | Task |
| --- | ----- | ---- |
| `next_return` | Next Trading Day Return | Regression on next-period return. |
| `next_close` | Next Trading Day Close | Regression on next close. |
| `direction` | Next Trading Day Direction | Binary classification. |

---

## Feature sets

| Value | Contents |
| ----- | -------- |
| **Basic** | Returns, volume deltas, lags. |
| **Technical** | RSI/MACD/Bollinger-style columns (no Basic overlap). |
| **All** | Basic + Technical. |

---

## Model & training ratio

Several algorithms may be compared internally; the headline pick follows the primary metric for each target type. Inspect **Model Evaluation** for the full roster.

---

## Threshold, strategy mode, costs

- **Trading Threshold**: gate on forecasts.  
- **Strategy Mode**: Long/Cash, Long/Short, Short/Cash—textbook simplifications.  
- **Costs / Slippage (bps)**: deducted per turnover.  
- **Initial Capital**: starting equity for curves.

---

## Dashboard tabs

1. **Executive Review** — Brief, benchmark comparison chart, risk snippets.  
2. **Data Overview** — Prices, QA table, target distribution.  
3. **Feature Analysis** — Trends, correlation heatmap.  
4. **Model Evaluation** — Metrics, comparison, importance, governance.  
5. **Predictions** — Forecast vs realised.  
6. **Backtest** — Equity, drawdowns, signals.

---

## ZIP export

One folder per run containing comparison CSVs, predictions, features, backtest series and metrics.

---

## Troubleshooting

Network required for downloads; bad tickers or thin calendars need wider windows or different symbols.

---

## Data caveat

Vendor adjustments vary; cite primary sources when publishing elsewhere.

---

**Bottom line**: Configure one symbol on the left, run analysis, read **Strategy** vs **All Buy** vs **Benchmark** as historical sanity checks—not forecasts of future returns.
"""



def overview_panel() -> html.Div:
    """Static onboarding content for the Overview tab."""
    return html.Div(
        [
            html.H4("Overview · Guide", className="overview-hero mb-2"),
            dbc.Alert(
                [
                    html.Strong("Disclaimer · intended use: "),
                    "education, research, and personal experimentation only—not a solicitation. "
                    "You are responsible for any decisions outside this sandbox.",
                ],
                color="secondary",
                className="overview-disclaimer mb-3",
            ),
            dcc.Markdown(OVERVIEW_GUIDE_MARKDOWN, className="overview-markdown"),
        ],
        className="overview-panel",
    )


def resolve_symbol(preset: str | None, custom: str | None) -> str:
    """Single Yahoo-style symbol: override field wins if non-empty (first token only)."""
    for part in re.split(r"[,;\s]+", (custom or "").strip()):
        sym = part.strip().upper()
        if sym:
            return sym
    if preset is None or (isinstance(preset, str) and not str(preset).strip()):
        return "AAPL"
    return str(preset).strip().upper()


def analyze_single_symbol(
    *,
    selected_ticker: str,
    selected_benchmark: str,
    start_date,
    end_date,
    interval,
    target_type,
    feature_set,
    model_key,
    train_ratio,
    threshold,
    strategy_mode,
    transaction_cost,
    slippage,
    initial_capital,
) -> dict:
    prices = download_price_data(selected_ticker, start_date, end_date, interval=interval)
    benchmark_prices = download_price_data(selected_benchmark, start_date, end_date, interval=interval)
    model_data, feature_columns = build_feature_frame(prices, feature_set, target_type)
    if len(model_data) < MIN_MODEL_ROWS:
        raise DataDownloadError(
            "At least about 60 rows are required after building features "
            f"(currently {len(model_data):,}) for {selected_ticker}. Widen the date range, "
            "set Interval to Daily, or choose Feature Set → Basic to shorten indicator warmup."
        )
    model_keys = selected_model_keys(model_key, target_type)
    results = [
        train_and_evaluate(
            model_data=model_data,
            feature_columns=feature_columns,
            model_key=key,
            target_type=target_type,
            train_ratio=float(train_ratio),
        )
        for key in model_keys
    ]
    backtests = {
        result.model_key: run_backtest(
            predictions=result.predictions,
            initial_capital=float(initial_capital or 1_000_000),
            threshold=float(threshold or 0.0),
            transaction_cost_bps=float(transaction_cost or 0.0),
            slippage_bps=float(slippage or 0.0),
            strategy_mode=strategy_mode,
        )
        for result in results
    }
    comparison = build_model_comparison(results, backtests)
    primary_metric = primary_metric_for_target(target_type)
    result = choose_best_result(results, comparison, primary_metric)
    backtest, backtest_metrics = backtests[result.model_key]
    return {
        "selected_ticker": selected_ticker,
        "prices": prices,
        "benchmark_prices": benchmark_prices,
        "model_data": model_data,
        "feature_columns": feature_columns,
        "comparison": comparison,
        "primary_metric": primary_metric,
        "result": result,
        "backtest": backtest,
        "backtest_metrics": backtest_metrics,
    }


def build_export_zip_bytes(bundles: list[dict]) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for b in bundles:
            sym = b["selected_ticker"].replace("/", "-").replace("^", "")
            zf.writestr(f"{sym}/model_comparison.csv", b["comparison"].to_csv(index=False))
            zf.writestr(f"{sym}/predictions.csv", b["result"].predictions.to_csv())
            zf.writestr(f"{sym}/feature_importance.csv", b["result"].feature_importance.to_csv(index=False))
            zf.writestr(f"{sym}/backtest_series.csv", b["backtest"].to_csv())
            zf.writestr(
                f"{sym}/backtest_metrics.csv",
                backtest_metrics_to_frame(b["backtest_metrics"]).to_csv(index=False),
            )
        zf.writestr(
            "README.txt",
            "Poseidon export: one folder per ticker with CSV tables and backtest series.\n",
        )
    return buf.getvalue()


def render_dashboard_slice(
    bundle: dict,
    *,
    selected_benchmark: str,
    target_type: str,
    strategy_mode: str,
    interval: str,
    initial_capital: float,
) -> tuple:
    sym = bundle["selected_ticker"]
    prices = bundle["prices"]
    benchmark_prices = bundle["benchmark_prices"]
    model_data = bundle["model_data"]
    feature_columns = bundle["feature_columns"]
    comparison = bundle["comparison"]
    result = bundle["result"]
    backtest = bundle["backtest"]
    backtest_metrics = bundle["backtest_metrics"]
    primary_metric = bundle["primary_metric"]

    ic = float(initial_capital or 1_000_000)
    status = dbc.Alert(
        f"{sym} · {INTERVALS.get(interval, interval)} · {STRATEGY_MODES.get(strategy_mode, strategy_mode)} · "
        f"{len(prices):,} rows downloaded · {len(model_data):,} rows used · Model: {result.model_name}",
        color="success",
    )
    return (
        status,
        build_summary_cards(
            selected_ticker=sym,
            result=result,
            model_data=model_data,
            backtest_metrics=backtest_metrics,
            primary_metric=primary_metric,
            strategy_mode=strategy_mode,
        ),
        build_executive_brief(
            selected_ticker=sym,
            selected_benchmark=selected_benchmark,
            result=result,
            backtest_metrics=backtest_metrics,
            comparison=comparison,
            primary_metric=primary_metric,
        ),
        benchmark_comparison_chart(
            backtest=backtest,
            benchmark_prices=benchmark_prices,
            initial_capital=ic,
            benchmark_label=selected_benchmark,
            instrument_ticker=sym,
        ),
        risk_regime_chart(backtest),
        stress_scenario_chart(backtest_metrics),
        table_from_frame(build_institutional_risk_frame(backtest_metrics), "Institutional Risk Stack"),
        price_chart(prices),
        table_from_frame(build_data_summary(prices, model_data), "Market Data Summary"),
        target_distribution_chart(model_data, target_type),
        feature_chart(model_data, feature_columns),
        feature_correlation_heatmap(model_data, feature_columns),
        table_from_frame(metrics_to_frame(result), "Selected Model Metrics"),
        model_comparison_chart(comparison, primary_metric),
        table_from_frame(comparison, "Model Evaluation and Backtest Comparison"),
        table_from_frame(
            evaluate_factor_library(model_data, feature_columns).head(15),
            "Factor Research Leaderboard",
        ),
        table_from_frame(
            checks_to_frame(governance_checklist(model_data, feature_columns, result.predictions)),
            "Model Governance Checklist",
        ),
        table_from_frame(result.feature_importance, "Feature Contribution Matrix"),
        feature_importance_chart(result.feature_importance),
        prediction_chart(result.predictions, target_type),
        residual_chart(result.predictions, target_type),
        table_from_frame(backtest_metrics_to_frame(backtest_metrics), "Backtest Metrics"),
        backtest_equity_chart(backtest),
        drawdown_chart(backtest),
        rolling_performance_chart(backtest),
        return_distribution_chart(backtest),
        signal_chart(backtest),
    )


app = Dash(__name__, external_stylesheets=[dbc.themes.CYBORG])
server = app.server


def dropdown_options(items: dict[str, str]) -> list[dict[str, str]]:
    return [{"label": label, "value": key} for key, label in items.items()]


app.layout = dbc.Container(
    [
        dbc.Row(
            [
                dbc.Col(
                    [
                        html.H1("Poseidon Quant Lab", className="mt-4 app-title"),
                        html.P(
                            "A dark trading research dashboard for market data, feature engineering, model comparison, prediction diagnostics, and strategy backtesting.",
                            className="hero-subtitle",
                        ),
                        html.Div(
                            [
                                html.Span("Market Data", className="terminal-badge"),
                                html.Span("ML Model Arena", className="terminal-badge"),
                                html.Span("Backtest Lab", className="terminal-badge"),
                                html.Span("Risk Diagnostics", className="terminal-badge"),
                            ],
                            className="badge-row",
                        ),
                    ]
                )
            ]
        ),
        dbc.Alert(
            "This app is for analysis and learning only. It is not investment advice.",
            color="success",
            className="mt-2 disclaimer",
        ),
        dcc.Store(id="export-store"),
        dcc.Download(id="download-results"),
        html.Div(
            dcc.Interval(id="bootstrap-interval", interval=800, max_intervals=1, n_intervals=0),
            style={"display": "none"},
        ),
        dbc.Row(
            [
                dbc.Col(
                    [
                        dbc.Card(
                            dbc.CardBody(
                                [
                                    dbc.Label("Ticker"),
                                    dcc.Dropdown(
                                        id="ticker",
                                        options=[{"label": t, "value": t} for t in PRESET_TICKERS],
                                        value="AAPL",
                                        multi=False,
                                        clearable=False,
                                        searchable=True,
                                        className="sidebar-dropdown",
                                    ),
                                    dbc.Label("Symbol override (optional)", className="mt-2"),
                                    dbc.Input(
                                        id="custom-ticker",
                                        type="text",
                                        placeholder="e.g. MSFT or 7203.T — overrides preset if set",
                                    ),
                                    html.Small(
                                        "Leave blank to use the preset ticker. One symbol only.",
                                        className="text-muted d-block mt-1 workspace-hint",
                                    ),
                                    dbc.Label("Benchmark", className="mt-2"),
                                    dbc.Input(
                                        id="benchmark-ticker",
                                        type="text",
                                        value="SPY",
                                        placeholder="SPY, ^GSPC",
                                    ),
                                    dbc.Label("Date Range", className="mt-2"),
                                    dcc.DatePickerRange(
                                        id="date-range",
                                        start_date=DEFAULT_START,
                                        end_date=DEFAULT_END,
                                        display_format="YYYY-MM-DD",
                                        className="sidebar-date-range",
                                        with_portal=False,
                                        number_of_months_shown=1,
                                        calendar_orientation="horizontal",
                                        minimum_nights=0,
                                    ),
                                    dbc.Label("Interval", className="mt-2"),
                                    dcc.Dropdown(
                                        id="interval",
                                        options=dropdown_options(INTERVALS),
                                        value="1d",
                                        clearable=False,
                                        className="sidebar-dropdown",
                                    ),
                                    dbc.Label("Prediction Target", className="mt-2"),
                                    dcc.Dropdown(
                                        id="target-type",
                                        options=dropdown_options(TARGETS),
                                        value="next_return",
                                        clearable=False,
                                        className="sidebar-dropdown",
                                    ),
                                    dbc.Label("Feature Set", className="mt-2"),
                                    dcc.Dropdown(
                                        id="feature-set",
                                        options=dropdown_options(FEATURE_SETS),
                                        value="all",
                                        clearable=False,
                                        className="sidebar-dropdown",
                                    ),
                                    dbc.Label("Model", className="mt-2"),
                                    dcc.Dropdown(
                                        id="model-key",
                                        clearable=False,
                                        className="sidebar-dropdown",
                                    ),
                                    dbc.Label("Training Data Ratio", className="mt-2"),
                                    dcc.Slider(
                                        id="train-ratio",
                                        min=0.5,
                                        max=0.9,
                                        step=0.05,
                                        value=0.75,
                                        marks={0.5: "50%", 0.75: "75%", 0.9: "90%"},
                                    ),
                                    dbc.Label("Trading Threshold", className="mt-2"),
                                    dbc.Input(id="threshold", type="number", value=0.0, step=0.001),
                                    dbc.Label("Strategy Mode", className="mt-2"),
                                    dcc.Dropdown(
                                        id="strategy-mode",
                                        options=dropdown_options(STRATEGY_MODES),
                                        value="long_cash",
                                        clearable=False,
                                        className="sidebar-dropdown",
                                    ),
                                    dbc.Label("Transaction Cost (bps)", className="mt-2"),
                                    dbc.Input(id="transaction-cost", type="number", value=5, min=0),
                                    dbc.Label("Slippage (bps)", className="mt-2"),
                                    dbc.Input(id="slippage", type="number", value=5, min=0),
                                    dbc.Label("Initial Capital", className="mt-2"),
                                    dbc.Input(id="initial-capital", type="number", value=1_000_000, min=1),
                                    dbc.Button(
                                        "Run Full Analysis",
                                        id="run-button",
                                        color="success",
                                        className="mt-3 run-button w-100",
                                    ),
                                    dbc.Button(
                                        "Download results (ZIP)",
                                        id="download-btn",
                                        color="success",
                                        outline=True,
                                        className="mt-2 w-100",
                                    ),
                                ],
                                className="settings-panel",
                            ),
                            className="mt-3 control-card",
                        ),
                    ],
                    xs=12,
                    lg=3,
                    className="settings-column",
                ),
                dbc.Col(
                    [
                        html.Div(id="status-message", className="mt-2"),
                        html.Div(id="summary-cards"),
                        dbc.Tabs(
                            [
                                dbc.Tab(
                                    overview_panel(),
                                    label="Overview · Guide",
                                    tab_id="tab-overview",
                                ),
                                dbc.Tab(
                                    dcc.Loading(
                                        [
                                            dbc.Tabs(
                                                [
                                                    dbc.Tab(
                                                        [
                                                            dbc.Row(
                                                                [
                                                                    dbc.Col(html.Div(id="executive-brief"), lg=5),
                                                                    dbc.Col(
                                                                        dcc.Graph(id="benchmark-comparison-chart"),
                                                                        lg=7,
                                                                    ),
                                                                ],
                                                                className="g-3 mt-2",
                                                            ),
                                                            dbc.Row(
                                                                [
                                                                    dbc.Col(dcc.Graph(id="risk-regime-chart"), lg=7),
                                                                    dbc.Col(
                                                                        dcc.Graph(id="stress-scenario-chart"),
                                                                        lg=5,
                                                                    ),
                                                                ],
                                                                className="g-3",
                                                            ),
                                                            html.Div(id="institutional-risk-table", className="mt-3"),
                                                        ],
                                                        label="Executive Review",
                                                    ),
                                                    dbc.Tab(
                                                        [
                                                            dbc.Row(
                                                                [
                                                                    dbc.Col(dcc.Graph(id="price-chart"), lg=8),
                                                                    dbc.Col(html.Div(id="data-summary-table"), lg=4),
                                                                ],
                                                                className="g-3 mt-2",
                                                            ),
                                                            dcc.Graph(id="target-distribution-chart"),
                                                        ],
                                                        label="Data Overview",
                                                    ),
                                                    dbc.Tab(
                                                        [
                                                            dcc.Graph(id="feature-chart"),
                                                            dcc.Graph(id="feature-correlation-chart"),
                                                        ],
                                                        label="Feature Analysis",
                                                    ),
                                                    dbc.Tab(
                                                        [
                                                            dbc.Row(
                                                                [
                                                                    dbc.Col(html.Div(id="model-metrics-table"), md=4),
                                                                    dbc.Col(
                                                                        dcc.Graph(id="model-comparison-chart"),
                                                                        md=4,
                                                                    ),
                                                                    dbc.Col(
                                                                        dcc.Graph(id="feature-importance-chart"),
                                                                        md=4,
                                                                    ),
                                                                ],
                                                                className="g-3 mt-2",
                                                            ),
                                                            html.Div(id="model-comparison-table", className="mt-3"),
                                                            dbc.Row(
                                                                [
                                                                    dbc.Col(
                                                                        html.Div(id="factor-research-table"),
                                                                        lg=6,
                                                                    ),
                                                                    dbc.Col(
                                                                        html.Div(id="governance-check-table"),
                                                                        lg=6,
                                                                    ),
                                                                ],
                                                                className="g-3 mt-3",
                                                            ),
                                                            html.Div(id="feature-contribution-table", className="mt-3"),
                                                        ],
                                                        label="Model Evaluation",
                                                    ),
                                                    dbc.Tab(
                                                        [
                                                            dcc.Graph(id="prediction-chart"),
                                                            dcc.Graph(id="residual-chart"),
                                                        ],
                                                        label="Predictions",
                                                    ),
                                                    dbc.Tab(
                                                        [
                                                            dbc.Row(
                                                                [
                                                                    dbc.Col(html.Div(id="backtest-metrics-table"), md=4),
                                                                    dbc.Col(dcc.Graph(id="equity-chart"), md=8),
                                                                ],
                                                                className="g-3 mt-2",
                                                            ),
                                                            dcc.Graph(id="drawdown-chart"),
                                                            dbc.Row(
                                                                [
                                                                    dbc.Col(
                                                                        dcc.Graph(id="rolling-performance-chart"),
                                                                        lg=7,
                                                                    ),
                                                                    dbc.Col(
                                                                        dcc.Graph(id="return-distribution-chart"),
                                                                        lg=5,
                                                                    ),
                                                                ],
                                                                className="g-3",
                                                            ),
                                                            dcc.Graph(id="signal-chart"),
                                                        ],
                                                        label="Backtest",
                                                    ),
                                                ],
                                                className="mt-2",
                                            )
                                        ],
                                        type="circle",
                                    ),
                                    label="Dashboard",
                                    tab_id="tab-analysis",
                                ),
                            ],
                            active_tab="tab-analysis",
                            id="main-workspace-tabs",
                            className="main-workspace-tabs mt-2",
                        ),
                    ],
                    xs=12,
                    lg=9,
                    className="results-scroll pe-lg-2",
                ),
            ],
            className="workspace-row g-3 mt-1",
        ),
    ],
    fluid=True,
)


@app.callback(
    Output("model-key", "options"),
    Output("model-key", "value"),
    Input("target-type", "value"),
)
def update_model_options(target_type: str):
    models = available_models(target_type)
    options = dropdown_options(models)
    options = [{"label": "Compare All Models (Recommended)", "value": "all_models"}] + options
    fast_default = "ridge" if target_type != "direction" else "logistic"
    if fast_default not in models:
        fast_default = next(iter(models.keys()))
    return options, fast_default


def build_analysis_outputs(
    ticker,
    custom_ticker,
    benchmark_ticker,
    start_date,
    end_date,
    interval,
    target_type,
    feature_set,
    model_key,
    train_ratio,
    threshold,
    strategy_mode,
    transaction_cost,
    slippage,
    initial_capital,
):
    """Full analysis pipeline; used by Dash callback and offline smoke tests."""
    model_key_effective = model_key if model_key is not None else "all_models"
    try:
        sym = resolve_symbol(ticker, custom_ticker)
        selected_benchmark = normalize_ticker(benchmark_ticker, "SPY")
        bundle = analyze_single_symbol(
            selected_ticker=sym,
            selected_benchmark=selected_benchmark,
            start_date=start_date,
            end_date=end_date,
            interval=interval,
            target_type=target_type,
            feature_set=feature_set,
            model_key=model_key_effective,
            train_ratio=train_ratio,
            threshold=threshold,
            strategy_mode=strategy_mode,
            transaction_cost=transaction_cost,
            slippage=slippage,
            initial_capital=initial_capital,
        )
        zip_bytes = build_export_zip_bytes([bundle])
        export_data = {"zip_b64": base64.b64encode(zip_bytes).decode("ascii")}
    except Exception as exc:  # noqa: BLE001 - Dash should show recoverable user errors
        message = dbc.Alert(str(exc), color="danger")
        return default_outputs(message)

    ic = float(initial_capital or 1_000_000)
    dash_tuple = render_dashboard_slice(
        bundle,
        selected_benchmark=selected_benchmark,
        target_type=target_type,
        strategy_mode=strategy_mode,
        interval=interval,
        initial_capital=ic,
    )
    return (*dash_tuple, export_data)


@app.callback(
    Output("status-message", "children"),
    Output("summary-cards", "children"),
    Output("executive-brief", "children"),
    Output("benchmark-comparison-chart", "figure"),
    Output("risk-regime-chart", "figure"),
    Output("stress-scenario-chart", "figure"),
    Output("institutional-risk-table", "children"),
    Output("price-chart", "figure"),
    Output("data-summary-table", "children"),
    Output("target-distribution-chart", "figure"),
    Output("feature-chart", "figure"),
    Output("feature-correlation-chart", "figure"),
    Output("model-metrics-table", "children"),
    Output("model-comparison-chart", "figure"),
    Output("model-comparison-table", "children"),
    Output("factor-research-table", "children"),
    Output("governance-check-table", "children"),
    Output("feature-contribution-table", "children"),
    Output("feature-importance-chart", "figure"),
    Output("prediction-chart", "figure"),
    Output("residual-chart", "figure"),
    Output("backtest-metrics-table", "children"),
    Output("equity-chart", "figure"),
    Output("drawdown-chart", "figure"),
    Output("rolling-performance-chart", "figure"),
    Output("return-distribution-chart", "figure"),
    Output("signal-chart", "figure"),
    Output("export-store", "data"),
    Input("run-button", "n_clicks"),
    Input("bootstrap-interval", "n_intervals"),
    State("ticker", "value"),
    State("custom-ticker", "value"),
    State("benchmark-ticker", "value"),
    State("date-range", "start_date"),
    State("date-range", "end_date"),
    State("interval", "value"),
    State("target-type", "value"),
    State("feature-set", "value"),
    State("model-key", "value"),
    State("train-ratio", "value"),
    State("threshold", "value"),
    State("strategy-mode", "value"),
    State("transaction-cost", "value"),
    State("slippage", "value"),
    State("initial-capital", "value"),
    prevent_initial_call=True,
)
def run_analysis(
    n_clicks,
    bootstrap_ticks,
    ticker,
    custom_ticker,
    benchmark_ticker,
    start_date,
    end_date,
    interval,
    target_type,
    feature_set,
    model_key,
    train_ratio,
    threshold,
    strategy_mode,
    transaction_cost,
    slippage,
    initial_capital,
):
    trig = callback_context.triggered_id
    auto_run = trig == "bootstrap-interval" and bootstrap_ticks is not None and bootstrap_ticks >= 1
    manual_run = trig == "run-button" and n_clicks
    if not auto_run and not manual_run:
        raise PreventUpdate

    return build_analysis_outputs(
        ticker,
        custom_ticker,
        benchmark_ticker,
        start_date,
        end_date,
        interval,
        target_type,
        feature_set,
        model_key,
        train_ratio,
        threshold,
        strategy_mode,
        transaction_cost,
        slippage,
        initial_capital,
    )


def default_outputs(message):
    """Return tuple aligned with run_analysis outputs: figures vs HTML children."""
    return (
        message,
        html.Div(),
        html.Div(),
        empty_figure(),
        empty_figure(),
        empty_figure(),
        html.Div(),
        empty_figure(),
        html.Div(),
        empty_figure(),
        empty_figure(),
        empty_figure(),
        html.Div(),
        empty_figure(),
        html.Div(),
        html.Div(),
        html.Div(),
        html.Div(),
        empty_figure(),
        empty_figure(),
        empty_figure(),
        html.Div(),
        empty_figure(),
        empty_figure(),
        empty_figure(),
        empty_figure(),
        empty_figure(),
        None,
    )


@app.callback(
    Output("download-results", "data"),
    Input("download-btn", "n_clicks"),
    State("export-store", "data"),
    prevent_initial_call=True,
)
def download_analysis_bundle(_n_clicks, store_data):
    if not store_data or not store_data.get("zip_b64"):
        raise PreventUpdate
    raw = base64.b64decode(store_data["zip_b64"])
    return dcc.send_bytes(raw, filename="poseidon_analysis_export.zip")


def normalize_ticker(custom_ticker: str | None, preset_ticker: str) -> str:
    ticker = (custom_ticker or "").strip() or preset_ticker
    return ticker.upper()


def selected_model_keys(model_key: str, target_type: str) -> list[str]:
    models = available_models(target_type)
    if model_key == "all_models":
        return list(models.keys())
    return [model_key]


def primary_metric_for_target(target_type: str) -> str:
    if target_type == "direction":
        return "F1"
    if target_type == "next_close":
        return "RMSE"
    return "Direction Accuracy"


def choose_best_result(results, comparison: pd.DataFrame, primary_metric: str):
    ascending = primary_metric in {"RMSE", "MAE"}
    best_model = comparison.sort_values(primary_metric, ascending=ascending).iloc[0]["model_key"]
    return next(result for result in results if result.model_key == best_model)


def build_model_comparison(results, backtests: dict) -> pd.DataFrame:
    rows = []
    for result in results:
        _, backtest_metrics = backtests[result.model_key]
        rows.append(
            {
                "model_key": result.model_key,
                "model": result.model_name,
                **result.metrics,
                "Strategy Total Return": backtest_metrics["Strategy Total Return"],
                "Annual Return": backtest_metrics["Annual Return"],
                "Sharpe": backtest_metrics["Sharpe"],
                "Sortino": backtest_metrics["Sortino"],
                "Calmar": backtest_metrics["Calmar"],
                "Max Drawdown": backtest_metrics["Max Drawdown"],
                "VaR 95": backtest_metrics["VaR 95"],
                "CVaR 95": backtest_metrics["CVaR 95"],
                "Exposure": backtest_metrics["Exposure"],
                "Long Days": backtest_metrics["Long Days"],
                "Short Days": backtest_metrics["Short Days"],
                "Trades": backtest_metrics["Trades"],
            }
        )
    return pd.DataFrame(rows)


def build_institutional_risk_frame(metrics: dict[str, float]) -> pd.DataFrame:
    rows = [
        ("Total Return", format_percent(metrics["Strategy Total Return"]), "Outcome"),
        ("Annual Return", format_percent(metrics["Annual Return"]), "Outcome"),
        ("Annual Volatility", format_percent(metrics["Annual Volatility"]), "Risk"),
        ("Sharpe", format_metric_value(metrics["Sharpe"]), "Risk-adjusted"),
        ("Sortino", format_metric_value(metrics["Sortino"]), "Downside-adjusted"),
        ("Calmar", format_metric_value(metrics["Calmar"]), "Drawdown-adjusted"),
        ("Max Drawdown", format_percent(metrics["Max Drawdown"]), "Tail Risk"),
        ("VaR 95", format_percent(metrics["VaR 95"]), "Tail Risk"),
        ("CVaR 95", format_percent(metrics["CVaR 95"]), "Tail Risk"),
        ("Profit Factor", format_metric_value(metrics["Profit Factor"]), "Trade Quality"),
        ("Payoff Ratio", format_metric_value(metrics["Payoff Ratio"]), "Trade Quality"),
        ("Win Rate", format_percent(metrics["Win Rate"]), "Trade Quality"),
        ("Exposure", format_percent(metrics["Exposure"]), "Capital Usage"),
        ("Trades", f"{metrics['Trades']:,.0f}", "Turnover"),
    ]
    return pd.DataFrame(rows, columns=["metric", "value", "pillar"])


def build_executive_brief(
    selected_ticker: str,
    selected_benchmark: str,
    result,
    backtest_metrics: dict[str, float],
    comparison: pd.DataFrame,
    primary_metric: str,
):
    primary_metric_value = result.metrics.get(primary_metric, 0.0)
    report = generate_institutional_report(
        metrics=backtest_metrics,
        model_name=result.model_name,
        primary_metric=primary_metric,
        primary_metric_value=primary_metric_value,
    )
    _, verdict_class = investment_committee_verdict(backtest_metrics)
    ascending = primary_metric in {"RMSE", "MAE"}
    ranked = comparison.sort_values(primary_metric, ascending=ascending).reset_index(drop=True)
    rank = ranked.index[ranked["model_key"].eq(result.model_key)].tolist()
    rank_text = f"#{rank[0] + 1}" if rank else "N/A"
    bullets = [
        f"Selected {result.model_name} for {selected_ticker}; primary metric is {primary_metric}.",
        f"Strategy return is {format_percent(backtest_metrics['Strategy Total Return'])} with Sharpe {format_metric_value(backtest_metrics['Sharpe'])}.",
        f"Tail profile: max drawdown {format_percent(backtest_metrics['Max Drawdown'])}, CVaR 95 {format_percent(backtest_metrics['CVaR 95'])}.",
        f"Benchmark lens uses {selected_benchmark}; model leaderboard rank is {rank_text}.",
    ]
    diligence_frame = due_diligence_to_frame(report.due_diligence).head(6)
    return dbc.Card(
        dbc.CardBody(
            [
                html.Div("Executive Brief", className="brief-eyebrow"),
                html.Div(report.verdict, className=f"brief-verdict {verdict_class}"),
                html.Div(
                    f"Institutional Grade {report.grade} / Score {report.score:.1f}",
                    className="brief-title",
                ),
                html.Ul([html.Li(item) for item in bullets], className="brief-list"),
                html.Div("Top Due Diligence Checks", className="brief-eyebrow"),
                dash_table.DataTable(
                    data=diligence_frame.to_dict("records"),
                    columns=[{"name": col, "id": col} for col in diligence_frame.columns],
                    style_cell={
                        "backgroundColor": "#04120a",
                        "border": "1px solid #14532d",
                        "color": "#d9ffe5",
                        "fontSize": "12px",
                        "padding": "6px",
                        "textAlign": "left",
                    },
                    style_header={
                        "backgroundColor": "#062415",
                        "border": "1px solid #00ff88",
                        "color": "#00ff88",
                        "fontWeight": "bold",
                    },
                    page_size=6,
                ),
                html.Div(
                    "Use this as a research dashboard only. Production deployment still requires independent validation, controls, and compliance review.",
                    className="brief-footnote",
                ),
            ]
        ),
        className="executive-card",
    )


def investment_committee_verdict(metrics: dict[str, float]) -> tuple[str, str]:
    sharpe = metrics.get("Sharpe", 0.0)
    max_drawdown = metrics.get("Max Drawdown", 0.0)
    total_return = metrics.get("Strategy Total Return", 0.0)
    if sharpe >= 1.0 and total_return > 0 and max_drawdown > -0.2:
        return "Advance To Deeper Due Diligence", "verdict-positive"
    if sharpe >= 0.3 and max_drawdown > -0.35:
        return "Watchlist: Needs Robustness Work", "verdict-watch"
    return "Do Not Allocate Without Redesign", "verdict-negative"


def build_data_summary(prices: pd.DataFrame, model_data: pd.DataFrame) -> pd.DataFrame:
    close = prices["Adj Close"]
    total_return = close.iloc[-1] / close.iloc[0] - 1
    daily_returns = close.pct_change().dropna()
    annualized_vol = daily_returns.std() * (252**0.5)
    rows = [
        ("Rows Downloaded", f"{len(prices):,}"),
        ("Model-Ready Rows", f"{len(model_data):,}"),
        ("Start Date", prices.index.min().strftime("%Y-%m-%d")),
        ("End Date", prices.index.max().strftime("%Y-%m-%d")),
        ("Price Return", format_percent(total_return)),
        ("Annualized Volatility", format_percent(annualized_vol)),
        ("Average Volume", f"{prices['Volume'].mean():,.0f}"),
        ("Latest Close", f"{close.iloc[-1]:,.2f}"),
    ]
    return pd.DataFrame(rows, columns=["metric", "value"])


def build_summary_cards(
    selected_ticker: str,
    result,
    model_data: pd.DataFrame,
    backtest_metrics: dict[str, float],
    primary_metric: str,
    strategy_mode: str,
):
    primary_value = result.metrics.get(primary_metric, 0.0)
    cards = [
        ("Instrument", selected_ticker, "Universe"),
        ("Selected Model", result.model_name, "Best by primary metric"),
        (primary_metric, format_metric_value(primary_value), "Model score"),
        ("Strategy Return", format_percent(backtest_metrics["Strategy Total Return"]), "Net of cost"),
        ("Sharpe", format_metric_value(backtest_metrics["Sharpe"]), "Annualized"),
        ("Max Drawdown", format_percent(backtest_metrics["Max Drawdown"]), "Peak-to-trough"),
        ("Exposure", format_percent(backtest_metrics["Exposure"]), STRATEGY_MODES.get(strategy_mode, strategy_mode)),
        ("Sample Size", f"{len(model_data):,}", "Model-ready rows"),
    ]
    return dbc.Row(
        [
            dbc.Col(
                dbc.Card(
                    dbc.CardBody(
                        [
                            html.Div(label, className="kpi-label"),
                            html.Div(value, className="kpi-value"),
                            html.Div(caption, className="kpi-caption"),
                        ]
                    ),
                    className="kpi-card",
                ),
                xl=3,
                lg=3,
                md=6,
                sm=12,
            )
            for label, value, caption in cards
        ],
        className="g-3",
    )


def table_from_frame(frame: pd.DataFrame, title: str):
    formatted = frame.copy()
    for column in formatted.columns:
        if pd.api.types.is_numeric_dtype(formatted[column]):
            formatted[column] = formatted[column].map(format_metric_value)
    return dbc.Card(
        dbc.CardBody(
            [
                html.H5(title),
                dash_table.DataTable(
                    data=formatted.to_dict("records"),
                    columns=[{"name": col, "id": col} for col in formatted.columns],
                    export_format="csv",
                    filter_action="native",
                    page_size=12,
                    sort_action="native",
                    style_cell={
                        "backgroundColor": "#04120a",
                        "border": "1px solid #14532d",
                        "color": "#d9ffe5",
                        "fontFamily": "Consolas, monospace",
                        "padding": "8px",
                        "textAlign": "left",
                    },
                    style_header={
                        "backgroundColor": "#062415",
                        "border": "1px solid #00ff88",
                        "color": "#00ff88",
                        "fontWeight": "bold",
                    },
                    style_data_conditional=[
                        {
                            "if": {"row_index": "odd"},
                            "backgroundColor": "#061a0f",
                        }
                    ],
                    style_table={"overflowX": "auto"},
                ),
            ]
        )
    )


def format_metric_value(value: float) -> str:
    if abs(value) >= 10:
        return f"{value:,.2f}"
    return f"{value:.4f}"


def format_percent(value: float) -> str:
    return f"{value * 100:,.2f}%"


if __name__ == "__main__":
    app.run(debug=True)
