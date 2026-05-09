from __future__ import annotations

from dataclasses import dataclass

import numpy as np
import pandas as pd
from sklearn.ensemble import (
    AdaBoostClassifier,
    AdaBoostRegressor,
    BaggingClassifier,
    BaggingRegressor,
    ExtraTreesClassifier,
    ExtraTreesRegressor,
    GradientBoostingClassifier,
    GradientBoostingRegressor,
    HistGradientBoostingClassifier,
    HistGradientBoostingRegressor,
    RandomForestClassifier,
    RandomForestRegressor,
)
from sklearn.inspection import permutation_importance
from sklearn.linear_model import (
    BayesianRidge,
    ElasticNet,
    HuberRegressor,
    Lasso,
    LinearRegression,
    LogisticRegression,
    Ridge,
    RidgeClassifier,
    SGDClassifier,
    SGDRegressor,
)
from sklearn.metrics import (
    accuracy_score,
    f1_score,
    mean_absolute_error,
    mean_squared_error,
    precision_score,
    r2_score,
    recall_score,
)
from sklearn.naive_bayes import GaussianNB
from sklearn.neighbors import KNeighborsClassifier, KNeighborsRegressor
from sklearn.neural_network import MLPClassifier, MLPRegressor
from sklearn.pipeline import Pipeline
from sklearn.preprocessing import StandardScaler
from sklearn.svm import SVC, SVR
from sklearn.tree import DecisionTreeClassifier, DecisionTreeRegressor
from sklearn.discriminant_analysis import LinearDiscriminantAnalysis, QuadraticDiscriminantAnalysis
from xgboost import XGBClassifier, XGBRegressor


REGRESSION_MODELS = {
    "baseline_momentum": "Baseline: 1D Momentum",
    "linear": "Linear Regression",
    "ridge": "Ridge",
    "lasso": "Lasso",
    "elastic_net": "ElasticNet",
    "bayesian_ridge": "Bayesian Ridge",
    "huber": "Huber Regressor",
    "sgd_regressor": "SGD Regressor",
    "passive_aggressive_regressor": "Passive Aggressive Regressor",
    "svr_rbf": "SVR (RBF Kernel)",
    "knn_regressor": "KNN Regressor",
    "decision_tree_regressor": "Decision Tree Regressor",
    "random_forest_regressor": "Random Forest Regressor",
    "extra_trees_regressor": "Extra Trees Regressor",
    "adaboost_regressor": "AdaBoost Regressor",
    "bagging_regressor": "Bagging Regressor",
    "gradient_boosting_regressor": "Gradient Boosting Regressor",
    "hist_gradient_boosting_regressor": "Hist Gradient Boosting Regressor",
    "mlp_regressor": "MLP Regressor",
    "xgboost_regressor": "XGBoost Regressor",
}

CLASSIFICATION_MODELS = {
    "baseline_momentum": "Baseline: 1D Momentum",
    "logistic": "Logistic Regression",
    "ridge_classifier": "Ridge Classifier",
    "sgd_classifier": "SGD Classifier",
    "passive_aggressive_classifier": "Passive Aggressive Classifier",
    "lda": "Linear Discriminant Analysis",
    "qda": "Quadratic Discriminant Analysis",
    "gaussian_nb": "Gaussian Naive Bayes",
    "svc_rbf": "SVC (RBF Kernel)",
    "knn_classifier": "KNN Classifier",
    "decision_tree_classifier": "Decision Tree Classifier",
    "random_forest_classifier": "Random Forest Classifier",
    "extra_trees_classifier": "Extra Trees Classifier",
    "adaboost_classifier": "AdaBoost Classifier",
    "bagging_classifier": "Bagging Classifier",
    "gradient_boosting_classifier": "Gradient Boosting Classifier",
    "hist_gradient_boosting_classifier": "Hist Gradient Boosting Classifier",
    "mlp_classifier": "MLP Classifier",
    "xgboost_classifier": "XGBoost Classifier",
}


@dataclass
class ModelResult:
    model_key: str
    model_name: str
    metrics: dict[str, float]
    predictions: pd.DataFrame
    feature_importance: pd.DataFrame


def available_models(target_type: str) -> dict[str, str]:
    if target_type == "direction":
        return CLASSIFICATION_MODELS
    return REGRESSION_MODELS


def train_and_evaluate(
    model_data: pd.DataFrame,
    feature_columns: list[str],
    model_key: str,
    target_type: str,
    train_ratio: float,
) -> ModelResult:
    if len(model_data) < 60:
        raise ValueError(
            "At least about 60 rows are required after cleaning features "
            f"(currently {len(model_data)}). Widen the date range, set Interval to Daily, "
            "or choose Feature Set → Basic to shorten indicator warmup."
        )

    train_size = int(len(model_data) * train_ratio)
    train_size = min(max(train_size, 30), len(model_data) - 10)

    train = model_data.iloc[:train_size]
    test = model_data.iloc[train_size:]

    if model_key == "baseline_momentum":
        raw_predictions = build_baseline_predictions(test, target_type)
        prediction_frame = build_prediction_frame(test, raw_predictions, target_type)
        metrics = calculate_metrics(test["target"], raw_predictions, target_type)
        importance = pd.DataFrame(
            [{"feature": "return_1d", "importance": 1.0, "method": "baseline"}]
        )
        return ModelResult(
            model_key=model_key,
            model_name=available_models(target_type)[model_key],
            metrics=metrics,
            predictions=prediction_frame,
            feature_importance=importance,
        )

    x_train = train[feature_columns]
    y_train = train["target"]
    x_test = test[feature_columns]
    y_test = test["target"]

    model = build_model(model_key, target_type)
    model.fit(x_train, y_train)

    raw_predictions = model.predict(x_test)
    prediction_frame = build_prediction_frame(test, raw_predictions, target_type)
    metrics = calculate_metrics(y_test, raw_predictions, target_type)
    importance = extract_feature_importance(
        model=model,
        feature_columns=feature_columns,
        x_test=x_test,
        y_test=y_test,
        target_type=target_type,
    )

    return ModelResult(
        model_key=model_key,
        model_name=available_models(target_type)[model_key],
        metrics=metrics,
        predictions=prediction_frame,
        feature_importance=importance,
    )


def build_model(model_key: str, target_type: str):
    if target_type == "direction":
        models = {
            "logistic": Pipeline(
                [
                    ("scaler", StandardScaler()),
                    ("model", LogisticRegression(max_iter=1000)),
                ]
            ),
            "ridge_classifier": Pipeline(
                [
                    ("scaler", StandardScaler()),
                    ("model", RidgeClassifier()),
                ]
            ),
            "sgd_classifier": Pipeline(
                [
                    ("scaler", StandardScaler()),
                    ("model", SGDClassifier(loss="log_loss", max_iter=2000, random_state=42)),
                ]
            ),
            "passive_aggressive_classifier": Pipeline(
                [
                    ("scaler", StandardScaler()),
                    (
                        "model",
                        SGDClassifier(
                            loss="hinge",
                            penalty=None,
                            learning_rate="pa1",
                            eta0=1.0,
                            max_iter=2000,
                            random_state=42,
                        ),
                    ),
                ]
            ),
            "lda": Pipeline(
                [
                    ("scaler", StandardScaler()),
                    ("model", LinearDiscriminantAnalysis()),
                ]
            ),
            "qda": Pipeline(
                [
                    ("scaler", StandardScaler()),
                    ("model", QuadraticDiscriminantAnalysis(reg_param=0.05)),
                ]
            ),
            "gaussian_nb": Pipeline(
                [
                    ("scaler", StandardScaler()),
                    ("model", GaussianNB()),
                ]
            ),
            "svc_rbf": Pipeline(
                [
                    ("scaler", StandardScaler()),
                    ("model", SVC(C=1.0, kernel="rbf")),
                ]
            ),
            "knn_classifier": Pipeline(
                [
                    ("scaler", StandardScaler()),
                    ("model", KNeighborsClassifier(n_neighbors=15)),
                ]
            ),
            "decision_tree_classifier": DecisionTreeClassifier(
                random_state=42,
                min_samples_leaf=5,
                max_depth=6,
            ),
            "random_forest_classifier": RandomForestClassifier(
                n_estimators=250,
                random_state=42,
                min_samples_leaf=3,
            ),
            "extra_trees_classifier": ExtraTreesClassifier(
                n_estimators=300,
                random_state=42,
                min_samples_leaf=3,
                n_jobs=-1,
            ),
            "adaboost_classifier": AdaBoostClassifier(
                n_estimators=150,
                learning_rate=0.05,
                random_state=42,
            ),
            "bagging_classifier": BaggingClassifier(
                n_estimators=100,
                random_state=42,
                n_jobs=-1,
            ),
            "gradient_boosting_classifier": GradientBoostingClassifier(random_state=42),
            "hist_gradient_boosting_classifier": HistGradientBoostingClassifier(
                random_state=42,
                max_iter=250,
                learning_rate=0.03,
            ),
            "mlp_classifier": Pipeline(
                [
                    ("scaler", StandardScaler()),
                    (
                        "model",
                        MLPClassifier(
                            hidden_layer_sizes=(64, 32),
                            max_iter=500,
                            random_state=42,
                            early_stopping=True,
                        ),
                    ),
                ]
            ),
            "xgboost_classifier": XGBClassifier(
                n_estimators=300,
                max_depth=3,
                learning_rate=0.03,
                subsample=0.9,
                colsample_bytree=0.9,
                eval_metric="logloss",
                random_state=42,
                n_jobs=-1,
            ),
        }
    else:
        models = {
            "linear": Pipeline(
                [
                    ("scaler", StandardScaler()),
                    ("model", LinearRegression()),
                ]
            ),
            "ridge": Pipeline(
                [
                    ("scaler", StandardScaler()),
                    ("model", Ridge(alpha=1.0)),
                ]
            ),
            "lasso": Pipeline(
                [
                    ("scaler", StandardScaler()),
                    ("model", Lasso(alpha=0.0001, max_iter=10000)),
                ]
            ),
            "elastic_net": Pipeline(
                [
                    ("scaler", StandardScaler()),
                    ("model", ElasticNet(alpha=0.0001, l1_ratio=0.5, max_iter=10000)),
                ]
            ),
            "bayesian_ridge": Pipeline(
                [
                    ("scaler", StandardScaler()),
                    ("model", BayesianRidge()),
                ]
            ),
            "huber": Pipeline(
                [
                    ("scaler", StandardScaler()),
                    ("model", HuberRegressor(max_iter=1000)),
                ]
            ),
            "sgd_regressor": Pipeline(
                [
                    ("scaler", StandardScaler()),
                    ("model", SGDRegressor(max_iter=2000, random_state=42)),
                ]
            ),
            "passive_aggressive_regressor": Pipeline(
                [
                    ("scaler", StandardScaler()),
                    (
                        "model",
                        SGDRegressor(
                            loss="epsilon_insensitive",
                            penalty=None,
                            learning_rate="pa1",
                            eta0=1.0,
                            max_iter=2000,
                            random_state=42,
                        ),
                    ),
                ]
            ),
            "svr_rbf": Pipeline(
                [
                    ("scaler", StandardScaler()),
                    ("model", SVR(C=1.0, epsilon=0.001, kernel="rbf")),
                ]
            ),
            "knn_regressor": Pipeline(
                [
                    ("scaler", StandardScaler()),
                    ("model", KNeighborsRegressor(n_neighbors=15)),
                ]
            ),
            "decision_tree_regressor": DecisionTreeRegressor(
                random_state=42,
                min_samples_leaf=5,
                max_depth=6,
            ),
            "random_forest_regressor": RandomForestRegressor(
                n_estimators=250,
                random_state=42,
                min_samples_leaf=3,
            ),
            "extra_trees_regressor": ExtraTreesRegressor(
                n_estimators=300,
                random_state=42,
                min_samples_leaf=3,
                n_jobs=-1,
            ),
            "adaboost_regressor": AdaBoostRegressor(
                n_estimators=150,
                learning_rate=0.05,
                random_state=42,
            ),
            "bagging_regressor": BaggingRegressor(
                n_estimators=100,
                random_state=42,
                n_jobs=-1,
            ),
            "gradient_boosting_regressor": GradientBoostingRegressor(random_state=42),
            "hist_gradient_boosting_regressor": HistGradientBoostingRegressor(
                random_state=42,
                max_iter=250,
                learning_rate=0.03,
            ),
            "mlp_regressor": Pipeline(
                [
                    ("scaler", StandardScaler()),
                    (
                        "model",
                        MLPRegressor(
                            hidden_layer_sizes=(64, 32),
                            max_iter=500,
                            random_state=42,
                            early_stopping=True,
                        ),
                    ),
                ]
            ),
            "xgboost_regressor": XGBRegressor(
                n_estimators=300,
                max_depth=3,
                learning_rate=0.03,
                subsample=0.9,
                colsample_bytree=0.9,
                objective="reg:squarederror",
                random_state=42,
                n_jobs=-1,
            ),
        }

    if model_key not in models:
        raise ValueError(f"Unsupported model: {model_key}")
    return models[model_key]


def build_baseline_predictions(test: pd.DataFrame, target_type: str) -> np.ndarray:
    momentum = test["return_1d"].fillna(0)
    if target_type == "direction":
        return (momentum > 0).astype(int).to_numpy()
    if target_type == "next_close":
        return (test["Adj Close"] * (1 + momentum)).to_numpy()
    return momentum.to_numpy()


def build_prediction_frame(
    test: pd.DataFrame,
    raw_predictions: np.ndarray,
    target_type: str,
) -> pd.DataFrame:
    predictions = test[["Close", "Adj Close", "forward_return", "target"]].copy()
    predictions["prediction"] = raw_predictions

    if target_type == "direction":
        predictions["actual_return"] = predictions["forward_return"]
        predictions["predicted_return"] = np.where(predictions["prediction"] > 0, 0.01, -0.01)
    elif target_type == "next_close":
        predictions["actual_return"] = predictions["target"] / predictions["Adj Close"] - 1
        predictions["predicted_return"] = predictions["prediction"] / predictions["Adj Close"] - 1
    else:
        predictions["actual_return"] = predictions["target"]
        predictions["predicted_return"] = predictions["prediction"]

    predictions["actual_direction"] = (predictions["actual_return"] > 0).astype(int)
    predictions["predicted_direction"] = (predictions["predicted_return"] > 0).astype(int)
    return predictions.dropna()


def calculate_metrics(
    y_true: pd.Series,
    y_pred: np.ndarray,
    target_type: str,
) -> dict[str, float]:
    if target_type == "direction":
        return {
            "Accuracy": float(accuracy_score(y_true, y_pred)),
            "Precision": float(precision_score(y_true, y_pred, zero_division=0)),
            "Recall": float(recall_score(y_true, y_pred, zero_division=0)),
            "F1": float(f1_score(y_true, y_pred, zero_division=0)),
        }

    direction_accuracy = np.mean(np.sign(y_true) == np.sign(y_pred))
    return {
        "MAE": float(mean_absolute_error(y_true, y_pred)),
        "RMSE": float(np.sqrt(mean_squared_error(y_true, y_pred))),
        "R2": float(r2_score(y_true, y_pred)),
        "Direction Accuracy": float(direction_accuracy),
    }


def extract_feature_importance(
    model,
    feature_columns: list[str],
    x_test: pd.DataFrame,
    y_test: pd.Series,
    target_type: str,
) -> pd.DataFrame:
    rows = []
    estimator = model.named_steps["model"] if isinstance(model, Pipeline) else model
    if hasattr(estimator, "feature_importances_"):
        values = np.ravel(estimator.feature_importances_)
        rows.extend(
            {
                "feature": feature,
                "importance": float(abs(value)),
                "signed_importance": float(value),
                "method": "native_importance",
            }
            for feature, value in zip(feature_columns, values, strict=False)
        )
    elif hasattr(estimator, "coef_"):
        values = np.ravel(estimator.coef_)
        rows.extend(
            {
                "feature": feature,
                "importance": float(abs(value)),
                "signed_importance": float(value),
                "method": "coefficient",
            }
            for feature, value in zip(feature_columns, values, strict=False)
        )

    rows.extend(
        extract_permutation_importance(
            model=model,
            feature_columns=feature_columns,
            x_test=x_test,
            y_test=y_test,
            target_type=target_type,
        )
    )

    if not rows:
        rows = [
            {
                "feature": feature,
                "importance": 0.0,
                "signed_importance": 0.0,
                "method": "unavailable",
            }
            for feature in feature_columns
        ]

    importance = pd.DataFrame(rows)
    return importance.sort_values("importance", ascending=False).head(40)


def extract_permutation_importance(
    model,
    feature_columns: list[str],
    x_test: pd.DataFrame,
    y_test: pd.Series,
    target_type: str,
) -> list[dict]:
    if x_test.empty:
        return []
    sample_size = min(len(x_test), 250)
    sampled_x = x_test.iloc[-sample_size:]
    sampled_y = y_test.iloc[-sample_size:]
    scoring = "f1" if target_type == "direction" else "neg_root_mean_squared_error"
    try:
        result = permutation_importance(
            model,
            sampled_x,
            sampled_y,
            n_repeats=3,
            random_state=42,
            scoring=scoring,
            n_jobs=1,
        )
    except Exception:
        return []
    return [
        {
            "feature": feature,
            "importance": float(abs(value)),
            "signed_importance": float(value),
            "method": "permutation",
        }
        for feature, value in zip(feature_columns, result.importances_mean, strict=False)
    ]


def metrics_to_frame(result: ModelResult) -> pd.DataFrame:
    return pd.DataFrame(
        [{"metric": key, "value": value} for key, value in result.metrics.items()]
    )
