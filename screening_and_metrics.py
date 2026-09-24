"""
Stock screening and cross-sectional market metrics.
"""

import numpy as np
import pandas as pd


def esg_negative_screening(df: pd.DataFrame, excluded_industries: list[str] | None = None,) -> pd.DataFrame:
    #Remove companies belonging to explicitly excluded industries.

    if excluded_industries is None:
        excluded_industries = ["Gambling", "Resorts & Casinos", "Tobacco"]

    if "industry" not in df.columns:
        raise ValueError("Input data must contain an 'industry' column.")

    return df[~df["industry"].fillna("").astype(str).str.strip().isin(excluded_industries)].reset_index(drop=True)


def sector_average_ratio(df: pd.DataFrame, ratio_column: str,) -> pd.Series:
    #Calculate the average value of a ratio for each sector.
    
    if not {"sector", ratio_column}.issubset(df.columns):
        raise ValueError(f"Input must contain 'sector' and '{ratio_column}' columns.")

    values = pd.to_numeric(df[ratio_column], errors="coerce")
    return values.groupby(df["sector"]).transform("mean")


def classify_stock_beta(beta: float | int | str) -> str:
    #Classify a stock as Defensive or Aggressive based on beta.
    
    if pd.isna(beta) or str(beta).upper() == "N/A":
        return "N/A"

    return "Defensive" if float(beta) < 1 else "Aggressive"


def calculate_effective_number_of_stocks(weights: pd.Series) -> float:
    """
    Calculate the effective number of holdings using the inverse HHI.

    Weights can be expressed either as decimals summing to 1 or as percentages.
    """
    w = pd.to_numeric(weights, errors="coerce").dropna().astype(float)

    if w.empty:
        raise ValueError("No valid weights supplied.")

    if w.sum() > 1.5:
        w = w / 100

    if not np.isclose(w.sum(), 1.0, atol=1e-6):
        w = w / w.sum()

    hhi = (w**2).sum()
    return float(1 / hhi)
