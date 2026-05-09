#!/usr/bin/env python3
"""Ten automated quality gates for Poseidon (exit 0 = all pass)."""

from __future__ import annotations

import compileall
import sys
from pathlib import Path

import numpy as np
import pandas as pd

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

EXPECTED_CALLBACK_OUTPUTS = 28


def _synthetic_prices(rows: int = 420) -> pd.DataFrame:
    rng = pd.date_range("2019-01-01", periods=rows, freq="B")
    rs = np.random.RandomState(42)
    close = np.linspace(80.0, 130.0, len(rng)) + rs.normal(0, 0.4, len(rng))
    close = np.maximum(close, 1.0)
    return pd.DataFrame(
        {
            "Open": close * (1 + rs.normal(0, 0.001, len(rng))),
            "High": close * 1.008,
            "Low": close * 0.992,
            "Close": close,
            "Adj Close": close,
            "Volume": np.full(len(rng), 1_000_000.0),
        },
        index=rng,
    )


def check_compileall() -> None:
    """1. Byte-compile all package modules and app entry."""
    src_ok = compileall.compile_dir(str(ROOT / "src"), quiet=1)
    app_ok = compileall.compile_file(str(ROOT / "app.py"), quiet=1)
    if not (src_ok and app_ok):
        raise RuntimeError("compileall reported failures.")


def check_app_default_outputs() -> None:
    """2. Import app and verify default_outputs arity matches dashboard."""
    import dash_bootstrap_components as dbc

    import app as app_module

    msg = dbc.Alert("ok", color="info")
    tpl = app_module.default_outputs(msg)
    if len(tpl) != EXPECTED_CALLBACK_OUTPUTS:
        raise AssertionError(
            f"default_outputs length {len(tpl)} != {EXPECTED_CALLBACK_OUTPUTS}"
        )


def check_regression_train() -> None:
    """3. Synthetic OHLCV → features → ridge regression."""
    from src.features import build_feature_frame
    from src.models import train_and_evaluate

    prices = _synthetic_prices()
    model_data, cols = build_feature_frame(prices, "basic", "next_return")
    result = train_and_evaluate(model_data, cols, "ridge", "next_return", 0.75)
    if result.predictions.empty or result.feature_importance.empty:
        raise AssertionError("Regression pipeline produced empty artifacts.")


def check_classification_train() -> None:
    """4. Classification direction path with logistic regression."""
    from src.features import build_feature_frame
    from src.models import train_and_evaluate

    prices = _synthetic_prices()
    model_data, cols = build_feature_frame(prices, "basic", "direction")
    result = train_and_evaluate(model_data, cols, "logistic", "direction", 0.75)
    if result.predictions.empty:
        raise AssertionError("Classification predictions empty.")


def check_baseline_path() -> None:
    """5. Baseline momentum regression (no sklearn fit)."""
    from src.features import build_feature_frame
    from src.models import train_and_evaluate

    prices = _synthetic_prices()
    model_data, cols = build_feature_frame(prices, "basic", "next_return")
    result = train_and_evaluate(model_data, cols, "baseline_momentum", "next_return", 0.75)
    if "predicted_return" not in result.predictions.columns:
        raise AssertionError("Baseline predictions missing predicted_return.")


def check_walk_forward() -> None:
    """6. Walk-forward windows + one ridge fit."""
    from src.features import build_feature_frame
    from src.walk_forward import build_walk_forward_windows, run_walk_forward_validation

    prices = _synthetic_prices(400)
    model_data, cols = build_feature_frame(prices, "basic", "next_return")
    wins = build_walk_forward_windows(model_data, train_size=120, test_size=40, step_size=40)
    if len(wins) < 2:
        raise AssertionError("Expected multiple walk-forward windows.")
    metrics, _preds = run_walk_forward_validation(
        model_data, cols, "ridge", "next_return", train_size=120, test_size=40, step_size=40
    )
    if metrics.empty:
        raise AssertionError("Walk-forward metrics empty.")


def check_institutional_report() -> None:
    """7. Institutional scoring memo."""
    from src.institutional import generate_institutional_report

    metrics = {
        "Sharpe": 0.9,
        "Sortino": 1.1,
        "Calmar": 0.4,
        "Max Drawdown": -0.15,
        "Strategy Total Return": 0.12,
        "Annual Return": 0.08,
    }
    report = generate_institutional_report(metrics, "Test Model", "Sharpe", 0.9)
    if report.score <= 0 or not report.grade:
        raise AssertionError("Institutional report incomplete.")


def check_reporting_markdown() -> None:
    """8. Markdown table export (requires tabulate)."""
    from src.reporting import format_table_markdown

    md = format_table_markdown(pd.DataFrame({"a": [1, 2], "b": [3.0, 4.0]}))
    if not md.strip():
        raise AssertionError("format_table_markdown returned empty string.")


def check_backtest() -> None:
    """9. Minimal backtest path."""
    from src.backtest import run_backtest

    idx = pd.date_range("2020-01-01", periods=50, freq="B")
    preds = pd.DataFrame(
        {
            "predicted_return": np.linspace(-0.01, 0.02, len(idx)),
            "actual_return": np.full(len(idx), 0.0005),
        },
        index=idx,
    )
    curve, metrics = run_backtest(
        preds,
        initial_capital=1_000_000.0,
        threshold=0.0,
        transaction_cost_bps=0.0,
        slippage_bps=0.0,
        strategy_mode="long_cash",
    )
    if curve.empty or "Sharpe" not in metrics:
        raise AssertionError("Backtest output incomplete.")


def check_live_run_analysis() -> None:
    """10. End-to-end pipeline with real yfinance data (single fast model)."""
    import app as app_module

    outs = app_module.build_analysis_outputs(
        "SPY",
        None,
        "SPY",
        "2023-01-01",
        "2024-06-01",
        "1d",
        "next_return",
        "basic",
        "ridge",
        0.75,
        0,
        "long_cash",
        0,
        0,
        1_000_000,
    )
    if len(outs) != EXPECTED_CALLBACK_OUTPUTS:
        raise AssertionError(f"build_analysis_outputs arity {len(outs)} != {EXPECTED_CALLBACK_OUTPUTS}")
    status = outs[0]
    props = getattr(status, "props", {}) or {}
    color = props.get("color")
    if color == "danger":
        raise RuntimeError(f"build_analysis_outputs failed: {props.get('children', status)}")


CHECKS = (
    ("Byte-compile src + app.py", check_compileall),
    ("Import app + default_outputs arity", check_app_default_outputs),
    ("Synthetic regression (ridge)", check_regression_train),
    ("Synthetic classification (logistic)", check_classification_train),
    ("Baseline momentum regression", check_baseline_path),
    ("Walk-forward validation", check_walk_forward),
    ("Institutional report", check_institutional_report),
    ("Reporting markdown", check_reporting_markdown),
    ("Backtest engine", check_backtest),
    ("Live run_analysis (SPY, ridge)", check_live_run_analysis),
)


def main() -> int:
    passed = 0
    for name, fn in CHECKS:
        try:
            fn()
        except Exception as exc:  # noqa: BLE001 — aggregate failures for summary
            print(f"[FAIL] {name}: {exc}")
            return 1
        passed += 1
        print(f"[PASS] {passed}/10 {name}")
    print(f"\nAll {passed} checks passed.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
