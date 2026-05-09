from __future__ import annotations

from dataclasses import dataclass

import numpy as np
import pandas as pd


@dataclass(frozen=True)
class GovernanceCheck:
    category: str
    check: str
    status: str
    severity: str
    detail: str


def feature_stability_report(
    train: pd.DataFrame,
    test: pd.DataFrame,
    feature_columns: list[str],
) -> pd.DataFrame:
    rows = []
    for feature in feature_columns:
        train_series = train[feature].dropna()
        test_series = test[feature].dropna()
        if train_series.empty or test_series.empty:
            rows.append(
                {
                    "feature": feature,
                    "mean_shift": np.nan,
                    "vol_shift": np.nan,
                    "psi": np.nan,
                    "status": "Review",
                }
            )
            continue
        train_std = train_series.std()
        mean_shift = (test_series.mean() - train_series.mean()) / train_std if train_std else 0.0
        vol_shift = test_series.std() / train_std - 1 if train_std else 0.0
        psi = population_stability_index(train_series, test_series)
        rows.append(
            {
                "feature": feature,
                "mean_shift": float(mean_shift),
                "vol_shift": float(vol_shift),
                "psi": float(psi),
                "status": stability_status(psi),
            }
        )
    return pd.DataFrame(rows).sort_values("psi", ascending=False)


def population_stability_index(
    expected: pd.Series,
    actual: pd.Series,
    buckets: int = 10,
) -> float:
    quantiles = np.linspace(0, 1, buckets + 1)
    breakpoints = expected.quantile(quantiles).drop_duplicates().to_numpy()
    if len(breakpoints) < 3:
        return 0.0
    expected_counts = pd.cut(expected, bins=breakpoints, include_lowest=True).value_counts(normalize=True)
    actual_counts = pd.cut(actual, bins=breakpoints, include_lowest=True).value_counts(normalize=True)
    aligned = pd.concat(
        [expected_counts.rename("expected"), actual_counts.rename("actual")],
        axis=1,
    ).fillna(0.0001)
    aligned = aligned.clip(lower=0.0001)
    psi = ((aligned["actual"] - aligned["expected"]) * np.log(aligned["actual"] / aligned["expected"])).sum()
    return float(psi)


def stability_status(psi: float) -> str:
    if psi >= 0.25:
        return "Breach"
    if psi >= 0.10:
        return "Warning"
    return "Pass"


def prediction_drift_report(predictions: pd.DataFrame) -> pd.DataFrame:
    if predictions.empty:
        return pd.DataFrame()
    frame = predictions.copy()
    frame["prediction_error"] = frame["target"] - frame["prediction"]
    midpoint = len(frame) // 2
    first = frame.iloc[:midpoint]
    second = frame.iloc[midpoint:]
    rows = [
        ("Prediction Mean First Half", first["prediction"].mean()),
        ("Prediction Mean Second Half", second["prediction"].mean()),
        ("Prediction Std First Half", first["prediction"].std()),
        ("Prediction Std Second Half", second["prediction"].std()),
        ("Error Mean First Half", first["prediction_error"].mean()),
        ("Error Mean Second Half", second["prediction_error"].mean()),
        ("Error Std First Half", first["prediction_error"].std()),
        ("Error Std Second Half", second["prediction_error"].std()),
    ]
    return pd.DataFrame(rows, columns=["metric", "value"])


def governance_checklist(
    model_data: pd.DataFrame,
    feature_columns: list[str],
    predictions: pd.DataFrame,
    min_rows: int = 250,
) -> list[GovernanceCheck]:
    checks = []
    checks.append(
        GovernanceCheck(
            category="Data",
            check="Minimum Sample Size",
            status="Pass" if len(model_data) >= min_rows else "Warning",
            severity="Medium" if len(model_data) < min_rows else "Low",
            detail=f"Model-ready rows: {len(model_data):,}. Preferred minimum: {min_rows:,}.",
        )
    )
    checks.append(
        GovernanceCheck(
            category="Data",
            check="Feature Count",
            status="Pass" if len(feature_columns) >= 8 else "Warning",
            severity="Low",
            detail=f"Feature count: {len(feature_columns):,}.",
        )
    )
    checks.append(
        GovernanceCheck(
            category="Prediction",
            check="Prediction Availability",
            status="Pass" if not predictions.empty else "Breach",
            severity="High" if predictions.empty else "Low",
            detail=f"Prediction rows: {len(predictions):,}.",
        )
    )
    if not predictions.empty:
        directional_hit = (predictions["actual_direction"] == predictions["predicted_direction"]).mean()
        checks.append(
            GovernanceCheck(
                category="Prediction",
                check="Directional Hit Rate",
                status="Pass" if directional_hit >= 0.52 else "Review",
                severity="Medium",
                detail=f"Directional hit rate: {directional_hit:.2%}.",
            )
        )
    checks.extend(leakage_proxy_checks(model_data, feature_columns))
    return checks


def leakage_proxy_checks(model_data: pd.DataFrame, feature_columns: list[str]) -> list[GovernanceCheck]:
    checks = []
    if "target" not in model_data:
        return checks
    correlations = model_data[feature_columns + ["target"]].corr(numeric_only=True)["target"].drop("target")
    suspicious = correlations[correlations.abs() > 0.95]
    checks.append(
        GovernanceCheck(
            category="Model Risk",
            check="Extreme Feature/Target Correlation",
            status="Review" if not suspicious.empty else "Pass",
            severity="High" if not suspicious.empty else "Low",
            detail=(
                "Suspicious features: " + ", ".join(suspicious.index[:5])
                if not suspicious.empty
                else "No feature has absolute correlation above 0.95 with target."
            ),
        )
    )
    return checks


def checks_to_frame(checks: list[GovernanceCheck]) -> pd.DataFrame:
    return pd.DataFrame(
        [
            {
                "category": check.category,
                "check": check.check,
                "status": check.status,
                "severity": check.severity,
                "detail": check.detail,
            }
            for check in checks
        ]
    )


def model_card(
    model_name: str,
    target_type: str,
    feature_set: str,
    train_ratio: float,
    metrics: dict[str, float],
) -> pd.DataFrame:
    rows = [
        ("Model", model_name),
        ("Target", target_type),
        ("Feature Set", feature_set),
        ("Train Ratio", f"{train_ratio:.0%}"),
    ]
    rows.extend((metric, value) for metric, value in metrics.items())
    return pd.DataFrame(rows, columns=["field", "value"])
