# yfinance ML Dash

This is a dark, green-accented Dash analytics app for downloading market data with yfinance, engineering features, training and evaluating machine learning models, visualizing predictions, benchmarking strategies, and running configurable backtests.

This app is for analysis and learning only. It is not investment advice.

## Features

- Download market data with yfinance
- Switch between daily, weekly, and monthly intervals
- Generate basic features, technical indicators, and lag features
- Predict next trading day return, next trading day close, or next trading day direction
- Train a broad model zoo: linear, regularized linear, robust linear, stochastic linear, SVM, KNN, tree, bagging, random forest, extra trees, AdaBoost, gradient boosting, histogram boosting, neural network, Naive Bayes, discriminant analysis, and XGBoost models
- Explain feature contribution with native feature importance, model coefficients, and model-agnostic permutation importance
- Compare all models and automatically select the best model using the primary metric
- Review executive KPI cards for selected model, return, Sharpe, drawdown, exposure, and sample size
- Use an Executive Review tab with an automated investment committee verdict
- Compare strategy equity against both instrument Buy & Hold and a configurable benchmark
- Monitor institutional risk metrics including Sortino, Calmar, VaR 95, CVaR 95, profit factor, and payoff ratio
- Inspect stress and tail-risk snapshots plus rolling volatility/drawdown regimes
- Generate institutional due-diligence checks, model governance checks, and strategy quality scores
- Use the expanded analytics modules for portfolio construction, efficient frontier research, risk contribution, and model stability review
- Research feature/factor quality with information coefficient, rank IC, t-stat, quantile returns, factor decay, and capacity proxies
- Estimate execution quality with spread, market impact, slippage, ticket cost, participation, liquidity buckets, capacity curves, and implementation shortfall
- Run walk-forward validation utilities, embargo split summaries, performance decay detection, and stability scoring
- Generate research memo objects, audit log records, schema summaries, and markdown-ready reports
- Inspect market data quality, target distribution, rolling strategy diagnostics, and return distributions
- Use preset tickers or enter any yfinance-compatible ticker
- Evaluate models with MAE, RMSE, R2, direction accuracy, Accuracy, Precision, Recall, and F1
- Visualize prediction vs actual, residuals, feature importance, feature correlations, and target distribution with Plotly
- Run long/cash, long/short, or short/cash backtests based on predicted returns
- Compare the strategy equity curve and drawdown against Buy & Hold
- Sort, filter, and export analytical tables as CSV

## Setup

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

## Run

```powershell
python app.py
```

After startup, open the Dash URL shown in the terminal. It is usually `http://127.0.0.1:8050/`.

## Usage

1. Select a preset ticker or enter a yfinance-compatible ticker such as `TSLA` or `8306.T`.
2. Select the date range, prediction target, feature set, and model. The default `Compare All Models` option evaluates all available models.
3. Set the benchmark, interval, training data ratio, trading threshold, strategy mode, transaction cost, slippage, and initial capital.
4. Click `Run Analysis` to download data, generate features, train models, evaluate results, and run the backtest.
5. Review the Executive Review, Data Overview, Feature Analysis, Model Evaluation, Predictions, and Backtest tabs.

## Executive Review

The first tab is designed for a fast investment committee style review:

- Automated verdict based on return, Sharpe, and drawdown profile
- Strategy vs instrument Buy & Hold vs benchmark equity curve
- Risk regime monitor using rolling volatility and drawdown
- Stress and tail-risk snapshot
- Institutional risk table for risk-adjusted return, tail risk, trade quality, and capital usage

## Expanded Architecture

The project is structured as a larger analytics system rather than a single dashboard script:

- `src/data.py`: yfinance download, normalization, and cached retrieval
- `src/features.py`: feature engineering and target construction
- `src/models.py`: model registry, training, prediction, evaluation, and feature importance
- `src/backtest.py`: strategy signals, long/short logic, and institutional risk metrics
- `src/plots.py`: dark-theme Plotly visualizations for executive review and diagnostics
- `src/institutional.py`: institutional scoring, risk limits, due diligence, drawdown events, scenario shocks, and alpha/beta utilities
- `src/model_governance.py`: model card utilities, feature stability, PSI, prediction drift, leakage proxy checks, and governance checklist
- `src/portfolio.py`: portfolio construction, minimum variance, max Sharpe search, efficient frontier, risk contribution, and rebalance drift
- `src/factor_research.py`: factor IC, rank IC, factor t-stat, turnover, quantile returns, decay profile, orthogonalization, and factor capacity proxy
- `src/execution.py`: order schedule, spread/impact/slippage/ticket cost, implementation shortfall, liquidity buckets, and capacity curve
- `src/walk_forward.py`: walk-forward windows, rolling validation, purged split summaries, stability score, and performance decay detection
- `src/reporting.py`: research memo generation, markdown formatting, audit log entries, schema summaries, and report snapshot comparisons
- `assets/style.css`: premium black and green UI system

## Model Comparison

When `Compare All Models` is selected, the app selects the best model using the primary metric for the selected target.

- Next trading day return: Direction Accuracy
- Next trading day close: RMSE
- Next trading day direction: F1

The Model Evaluation tab shows the selected model metrics, model-by-model comparison, and feature importance.

## Available Models

Regression targets:

- Baseline: 1D Momentum
- Linear Regression
- Ridge
- Lasso
- ElasticNet
- Bayesian Ridge
- Huber Regressor
- SGD Regressor
- Passive Aggressive Regressor
- SVR (RBF Kernel)
- KNN Regressor
- Decision Tree Regressor
- Random Forest Regressor
- Extra Trees Regressor
- AdaBoost Regressor
- Bagging Regressor
- Gradient Boosting Regressor
- Hist Gradient Boosting Regressor
- MLP Regressor
- XGBoost Regressor

Classification target:

- Baseline: 1D Momentum
- Logistic Regression
- Ridge Classifier
- SGD Classifier
- Passive Aggressive Classifier
- Linear Discriminant Analysis
- Quadratic Discriminant Analysis
- Gaussian Naive Bayes
- SVC (RBF Kernel)
- KNN Classifier
- Decision Tree Classifier
- Random Forest Classifier
- Extra Trees Classifier
- AdaBoost Classifier
- Bagging Classifier
- Gradient Boosting Classifier
- Hist Gradient Boosting Classifier
- MLP Classifier
- XGBoost Classifier

## Preset Tickers

- `AAPL`
- `MSFT`
- `GOOGL`
- `NVDA`
- `SPY`
- `QQQ`
- `7203.T`
- `6758.T`
- `9984.T`

## Notes

- Data may not be available depending on yfinance availability or ticker notation.
- The backtest is a simplified validation tool. It does not fully account for execution constraints, taxes, liquidity, dividends, or real-world trading limitations.
- The app uses chronological train/test splits and does not randomly shuffle time series data.
- Features are generated from information available before the prediction point to avoid future data leakage.
