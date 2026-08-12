"""Recursive multi-day-ahead demand forecast per SKU, exported to Excel."""
import os

import joblib
import lightgbm as lgb
import numpy as np
import pandas as pd

from features import CLIENT_SIGNAL_COLS, FEATURE_COLS, LAGS, ROLLING_WINDOWS
from ingest import SKU_COLS

DATA_DIR = os.path.join(os.path.dirname(__file__), "..", "data")
MODEL_DIR = os.path.join(os.path.dirname(__file__), "..", "models")
OUTPUT_DIR = os.path.join(os.path.dirname(__file__), "..", "output")

HORIZON_DAYS = 30
HISTORY_WINDOW = max(max(LAGS), max(ROLLING_WINDOWS)) + 1  # enough tail history to compute all features


def load_model():
    model = lgb.Booster(model_file=os.path.join(MODEL_DIR, "demand_model.txt"))
    cat_cols = joblib.load(os.path.join(MODEL_DIR, "cat_cols.joblib"))
    bias_correction = joblib.load(os.path.join(MODEL_DIR, "bias_correction.joblib"))
    return model, cat_cols, bias_correction


def align_dtypes(batch: pd.DataFrame, feat: pd.DataFrame) -> pd.DataFrame:
    """Match each column's dtype (and category codes, where relevant) to the training frame."""
    for col in batch.columns:
        if col not in feat.columns:
            continue
        if str(feat[col].dtype) == "category":
            batch[col] = pd.Categorical(batch[col], categories=feat[col].cat.categories)
        else:
            batch[col] = batch[col].astype(feat[col].dtype)
    return batch


def build_feature_row(
    history: np.ndarray, meta: dict, forecast_date: pd.Timestamp, client_signal_last: dict = None
) -> dict:
    row = dict(meta)
    row["day_of_week"] = forecast_date.dayofweek
    row["day_of_month"] = forecast_date.day
    row["month"] = forecast_date.month
    row["is_weekend"] = int(forecast_date.dayofweek in (5, 6))
    row["is_month_start"] = int(forecast_date.day <= 5)
    row["is_month_end"] = int(forecast_date.day >= 26)

    for lag in LAGS:
        row[f"lag_{lag}"] = history[-lag] if len(history) >= lag else np.nan

    for window in ROLLING_WINDOWS:
        tail = history[-window:]
        row[f"roll_mean_{window}"] = tail.mean() if len(tail) > 0 else np.nan
        row[f"roll_std_{window}"] = tail.std(ddof=1) if len(tail) > 1 else np.nan

    # reseller concentration/recency signals are exogenous - future values are
    # unknown, so carry forward each SKU's last observed values rather than
    # guess at future reseller behavior. days_since is bumped forward each
    # step since "no new client order" is the default assumption.
    for col in CLIENT_SIGNAL_COLS:
        row[col] = client_signal_last.get(col, 0) if client_signal_last else 0
    if client_signal_last is not None and "days_since_last_client_order" in client_signal_last:
        days_ahead = (forecast_date - client_signal_last["_as_of_date"]).days
        row["days_since_last_client_order"] = client_signal_last["days_since_last_client_order"] + days_ahead

    return row


def correction_for_sku(bias_correction, sku_keys: tuple, vendor: str) -> float:
    """bias_correction is {"global", "by_vendor", "by_sku"} from train.py — prefer the
    most specific factor available, falling back to vendor then global when the
    held-out period didn't have enough of that SKU's volume to estimate reliably."""
    if isinstance(bias_correction, dict):
        if sku_keys in bias_correction.get("by_sku", {}):
            return bias_correction["by_sku"][sku_keys]
        return bias_correction["by_vendor"].get(vendor, bias_correction["global"])
    return bias_correction


def forecast(
    feat: pd.DataFrame, model: lgb.Booster, horizon: int = HORIZON_DAYS, bias_correction=1.0
) -> pd.DataFrame:
    last_date = feat["date"].max()
    sku_cols_no_product = [c for c in SKU_COLS if c != "Product"]

    histories = {}
    metas = {}
    corrections = {}
    client_signal_lasts = {}
    for keys, group in feat.groupby(sku_cols_no_product, observed=True):
        group = group.sort_values("date")
        histories[keys] = group["quantity"].to_numpy()[-HISTORY_WINDOW:]
        metas[keys] = dict(zip(sku_cols_no_product, keys))
        corrections[keys] = correction_for_sku(bias_correction, keys, metas[keys]["Vendor"])

        last_row = group.iloc[-1]
        client_signal_lasts[keys] = {
            col: last_row[col] for col in CLIENT_SIGNAL_COLS if col in group.columns
        }
        client_signal_lasts[keys]["_as_of_date"] = last_row["date"]

    keys_order = list(histories.keys())
    correction_array = np.array([corrections[keys] for keys in keys_order])

    forecasts = []
    for step in range(1, horizon + 1):
        forecast_date = last_date + pd.Timedelta(days=step)
        rows = []
        for keys in keys_order:
            rows.append(build_feature_row(
                histories[keys], metas[keys], forecast_date, client_signal_lasts[keys]
            ))
        batch = pd.DataFrame(rows)
        batch = align_dtypes(batch, feat)

        preds = model.predict(batch[FEATURE_COLS])
        preds = np.clip(preds, 0, None) * correction_array

        for keys, pred in zip(keys_order, preds):
            histories[keys] = np.append(histories[keys], pred)[-HISTORY_WINDOW:]
            forecasts.append({**metas[keys], "date": forecast_date, "forecast_quantity": pred})

    return pd.DataFrame(forecasts)


def main():
    feat = pd.read_parquet(os.path.join(DATA_DIR, "features.parquet"))
    model, _, bias_correction = load_model()

    daily_forecast = forecast(feat, model, bias_correction=bias_correction)

    sku_cols_no_product = [c for c in SKU_COLS if c != "Product"]
    summary = (
        daily_forecast.groupby(sku_cols_no_product, as_index=False)["forecast_quantity"]
        .sum()
        .rename(columns={"forecast_quantity": f"forecast_{HORIZON_DAYS}day_total"})
        .sort_values(f"forecast_{HORIZON_DAYS}day_total", ascending=False)
    )

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    out_path = os.path.join(OUTPUT_DIR, "demand_forecast.xlsx")
    with pd.ExcelWriter(out_path, engine="xlsxwriter") as writer:
        summary.to_excel(writer, sheet_name="Monthly Summary", index=False)
        daily_forecast.to_excel(writer, sheet_name="Daily Forecast", index=False)

    print(f"Forecasted {daily_forecast[sku_cols_no_product].drop_duplicates().shape[0]} SKUs "
          f"for {HORIZON_DAYS} days ahead")
    print(f"Saved to {out_path}")


if __name__ == "__main__":
    main()
