# -*- coding: utf-8 -*-
"""
Created on Tue Oct 20 12:09:08 2026

@author: vikeshchellaram
"""

import numpy as np
import pandas as pd
import yfinance as yf
from scipy.stats import norm


def value_at_risk(tickers, weights, capital, period="5y", confidence_levels=(0.95, 0.99),
                  auto_adjust=True, progress=False,):
    """
    Calculate historical VaR, parametric VaR and CVaR for a stock portfolio.

    Parameters
    ----------
    tickers : list -> List of stock ticker symbols.
    weights : list, numpy.ndarray, or pandas.Series -> Portfolio weights. 
        They can be percentages or decimal values. Weights are normalized to sum to 1.
    capital : float -> Total portfolio capital.
    period : str -> Historical data period accepted by yfinance, e.g. "1y", "3y", "5y".
    confidence_levels : tuple -> Confidence levels for VaR and CVaR, e.g. (0.95, 0.99).
    auto_adjust : bool -> Whether to use adjusted prices.
    progress : bool -> Whether to display yfinance download progress.

    Returns
    -------
    dict
        Dictionary containing portfolio returns, VaR results, CVaR results, 
        data coverage, and normalized weights.
    """

    # 1. Validate inputs
    tickers = list(tickers)

    if not tickers:
        raise ValueError("Tickers cannot be empty.")

    if len(tickers) != len(set(tickers)):
        raise ValueError("Duplicate tickers are not allowed.")

    if capital <= 0 or not np.isfinite(capital):
        raise ValueError("Capital must be a positive finite number.")

    weights = np.asarray(weights, dtype=float)

    if len(tickers) != len(weights):
        raise ValueError("The number of tickers and weights must match.")

    if not np.isfinite(weights).all() or (weights < 0).any():
        raise ValueError("Weights must be finite, non-negative numbers.")

    if weights.sum() <= 0:
        raise ValueError("The sum of weights must be positive.")

    if not confidence_levels or any(not 0 < level < 1 for level in confidence_levels):
        raise ValueError("Confidence levels must be between 0 and 1.")

    # 2. Normalize weights
    original_weight_sum = weights.sum()
    weights = weights / original_weight_sum
    weights = pd.Series(weights, index=tickers)

    # 3. Download adjusted closing prices
    downloaded = yf.download(tickers, period=period, auto_adjust=auto_adjust,
                             progress=progress, group_by="column", threads=True)

    if downloaded.empty:
        raise ValueError("No price data was downloaded.")

    if isinstance(downloaded.columns, pd.MultiIndex):
        if "Close" not in downloaded.columns.get_level_values(0):
            raise ValueError("Closing prices are unavailable.")
        prices = downloaded["Close"]
    else:
        if len(tickers) != 1:
            raise ValueError("Unexpected price data format.")
        prices = downloaded[["Close"]].rename(columns={"Close": tickers[0]})

    prices = prices.reindex(columns=tickers)

    # 4. Check data availability
    missing_tickers = prices.columns[prices.isna().all()].tolist()

    if missing_tickers:
        raise ValueError(f"No price data available for: {missing_tickers}")

    data_coverage = pd.DataFrame({"First Available Price": prices.apply(
        lambda col: col.first_valid_index()),
        "Valid Price Observations": prices.notna().sum(),
        "Missing Prices": prices.isna().sum()})
    
    # 5. Calculate daily returns
    returns = prices.pct_change(fill_method=None)

    # Keep only dates with returns available for every stock
    returns = returns.dropna(how="any")

    if returns.empty:
        raise ValueError("Insufficient common return data to calculate VaR.")

    # 6. Calculate daily portfolio returns
    portfolio_returns = returns @ weights

    if len(portfolio_returns) < 2:
        raise ValueError("At least two return observations are required.")

    mean_return = portfolio_returns.mean()
    daily_volatility = portfolio_returns.std(ddof=1)

    # 7. Calculate risk metrics
    results = []

    for confidence_level in confidence_levels:

        # Historical VaR
        historical_threshold = portfolio_returns.quantile(1 - confidence_level)
        historical_var = -historical_threshold * capital

        # Parametric VaR assuming normally distributed returns
        parametric_threshold = (mean_return - norm.ppf(confidence_level) * daily_volatility)
        parametric_var = max(0, -parametric_threshold * capital)

        # Historical CVaR / Expected Shortfall
        tail_returns = portfolio_returns[portfolio_returns <= historical_threshold]
        historical_cvar = -tail_returns.mean() * capital

        results.append({"Confidence Level": confidence_level,
                        "Historical VaR": historical_var,
                        "Parametric VaR": parametric_var,
                        "Historical CVaR": historical_cvar,})

    risk_results = pd.DataFrame(results).set_index("Confidence Level")

    # 8. Print summary
    print("-" * 60)
    print("RISK ANALYSIS")
    print(f"Number of stocks: {len(tickers)}")
    print(f"Capital: ${capital:,.2f}")
    print(f"Historical period: {period}")
    print(f"Original weight sum: {original_weight_sum:.4f}")
    print(f"Number of daily observations: {len(portfolio_returns)}")
    print(f"Sample start date: {portfolio_returns.index.min().date()}")
    print(f"Sample end date: {portfolio_returns.index.max().date()}")

    print("\nRISK METRICS")
    print(risk_results.to_string(formatters={"Historical VaR": lambda x: f"${x:,.2f}",
                                             "Parametric VaR": lambda x: f"${x:,.2f}",
                                             "Historical CVaR": lambda x: f"${x:,.2f}",}))

    # print("\nDATA COVERAGE")
    # print(data_coverage.to_string())

    # 9. Return results for further analysis
    return {"portfolio_returns": portfolio_returns,
            "risk_metrics": risk_results,
            "data_coverage": data_coverage,
            "normalized_weights": weights,
            "prices": prices}

                    
