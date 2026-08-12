"""Feature engineering for per-SKU daily demand forecasting."""
import pandas as pd

from ingest import SKU_COLS

LAGS = [1, 7, 14, 28]
ROLLING_WINDOWS = [7, 14, 30]
CLIENT_WINDOWS = [7, 30]
CLIENT_SIGNAL_COLS = (
    [f"n_active_clients_{w}d" for w in CLIENT_WINDOWS]
    + [f"top_client_share_{w}d" for w in CLIENT_WINDOWS]
    + ["days_since_last_client_order"]
)


def add_date_features(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["day_of_week"] = df["date"].dt.dayofweek
    df["day_of_month"] = df["date"].dt.day
    df["month"] = df["date"].dt.month
    df["is_weekend"] = df["day_of_week"].isin([5, 6]).astype(int)
    df["is_month_start"] = (df["day_of_month"] <= 5).astype(int)
    df["is_month_end"] = (df["day_of_month"] >= 26).astype(int)
    return df


def add_lag_and_rolling_features(df: pd.DataFrame) -> pd.DataFrame:
    df = df.sort_values(SKU_COLS + ["date"]).copy()
    grouped = df.groupby(SKU_COLS, sort=False)["quantity"]

    for lag in LAGS:
        df[f"lag_{lag}"] = grouped.shift(lag)

    for window in ROLLING_WINDOWS:
        df[f"roll_mean_{window}"] = (
            df.groupby(SKU_COLS)["quantity"]
            .transform(lambda s: s.shift(1).rolling(window, min_periods=1).mean())
        )
        df[f"roll_std_{window}"] = (
            df.groupby(SKU_COLS)["quantity"]
            .transform(lambda s: s.shift(1).rolling(window, min_periods=2).std())
        )
    return df


def add_client_signal_features(df: pd.DataFrame, client_signal: pd.DataFrame) -> pd.DataFrame:
    """Left-join reseller concentration/recency signals from Market share.xlsx.

    Coverage is partial (only SKUs/dates present in that file) — SKUs or dates
    outside it get 0 active clients / 0 concentration (no known reseller
    activity) and a large "days since" sentinel, which LightGBM can split on
    same as any other value rather than treating it as missing.
    """
    merged = df.merge(client_signal, on=["Vendor", "Product Code", "date"], how="left")
    for col in CLIENT_SIGNAL_COLS:
        if col == "days_since_last_client_order":
            merged[col] = merged[col].fillna(9999)
        else:
            merged[col] = merged[col].fillna(0)
    return merged


def build_feature_frame(daily: pd.DataFrame, client_signal: pd.DataFrame = None) -> pd.DataFrame:
    df = add_date_features(daily)
    df = add_lag_and_rolling_features(df)
    if client_signal is not None:
        df = add_client_signal_features(df, client_signal)
    for col in SKU_COLS:
        df[col] = df[col].astype("category")
    return df


FEATURE_COLS = (
    [c for c in SKU_COLS if c != "Product"]  # "Product" is free text, redundant with Product Code
    + ["day_of_week", "day_of_month", "month", "is_weekend", "is_month_start", "is_month_end"]
    + [f"lag_{lag}" for lag in LAGS]
    + [f"roll_mean_{w}" for w in ROLLING_WINDOWS]
    + [f"roll_std_{w}" for w in ROLLING_WINDOWS]
    + CLIENT_SIGNAL_COLS
)
TARGET_COL = "quantity"


if __name__ == "__main__":
    import os

    data_dir = os.path.join(os.path.dirname(__file__), "..", "data")
    daily = pd.read_parquet(os.path.join(data_dir, "daily_demand.parquet"))

    client_signal_path = os.path.join(data_dir, "client_signal_features.parquet")
    client_signal = pd.read_parquet(client_signal_path) if os.path.exists(client_signal_path) else None
    if client_signal is None:
        print("No client_signal_features.parquet found - run market_share.py first. "
              "Continuing without reseller signal features.")

    feat = build_feature_frame(daily, client_signal)
    out_path = os.path.join(data_dir, "features.parquet")
    feat.to_parquet(out_path, index=False)
    print(f"Built features: {feat.shape}, saved to {out_path}")
