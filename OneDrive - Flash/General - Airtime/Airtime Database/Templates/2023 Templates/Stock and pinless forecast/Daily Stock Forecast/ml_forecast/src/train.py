"""Train a global LightGBM model to forecast daily demand per SKU."""
import os

import joblib
import lightgbm as lgb
import numpy as np
import pandas as pd
from sklearn.metrics import mean_absolute_error, mean_squared_error

from features import FEATURE_COLS, TARGET_COL
from ingest import SKU_COLS

DATA_DIR = os.path.join(os.path.dirname(__file__), "..", "data")
MODEL_DIR = os.path.join(os.path.dirname(__file__), "..", "models")

TEST_DAYS = 60

# a vendor/SKU with little held-out volume gives an unstable actual/predicted
# ratio — only trust a residual correction once volume clears these floors,
# and never let a factor stray far from its parent level's (already-checked)
# factor, so one thin-data SKU can't blow up into a nonsensical forecast.
MIN_VENDOR_TEST_VOLUME = 5000
MIN_SKU_TEST_VOLUME = 2000
MAX_DEVIATION_FROM_PARENT = 2.0


def wmape(y_true: np.ndarray, y_pred: np.ndarray) -> float:
    return np.abs(y_true - y_pred).sum() / max(np.abs(y_true).sum(), 1e-9)


def time_based_split(df: pd.DataFrame, test_days: int = TEST_DAYS):
    cutoff = df["date"].max() - pd.Timedelta(days=test_days)
    train = df[df["date"] <= cutoff]
    test = df[df["date"] > cutoff]
    return train, test


def main():
    feat = pd.read_parquet(os.path.join(DATA_DIR, "features.parquet"))
    train_df, test_df = time_based_split(feat)

    cat_cols = [c for c in FEATURE_COLS if str(feat[c].dtype) == "category"]

    train_set = lgb.Dataset(
        train_df[FEATURE_COLS], label=train_df[TARGET_COL], categorical_feature=cat_cols
    )
    valid_set = lgb.Dataset(
        test_df[FEATURE_COLS], label=test_df[TARGET_COL], categorical_feature=cat_cols,
        reference=train_set,
    )

    # Tweedie targets the conditional MEAN of a zero-inflated, right-skewed
    # distribution directly on the raw scale — unlike log1p+L1 (which targets
    # the median and needed large post-hoc rescaling to recover the mean total
    # that reorder decisions actually depend on).
    params = {
        "objective": "tweedie",
        "tweedie_variance_power": 1.2,
        "metric": "mae",
        "learning_rate": 0.05,
        "num_leaves": 63,
        "min_data_in_leaf": 50,
        "feature_fraction": 0.8,
        "bagging_fraction": 0.8,
        "bagging_freq": 1,
        "verbose": -1,
    }

    model = lgb.train(
        params,
        train_set,
        num_boost_round=2000,
        valid_sets=[train_set, valid_set],
        valid_names=["train", "valid"],
        callbacks=[lgb.early_stopping(100), lgb.log_evaluation(100)],
    )

    preds = np.clip(model.predict(test_df[FEATURE_COLS], num_iteration=model.best_iteration), 0, None)

    actual = test_df[TARGET_COL].to_numpy()
    mae = mean_absolute_error(actual, preds)
    rmse = mean_squared_error(actual, preds) ** 0.5
    wm = wmape(actual, preds)

    naive = test_df["roll_mean_7"].fillna(0).to_numpy()
    naive_wm = wmape(actual, naive)

    global_bias_correction = actual.sum() / max(preds.sum(), 1e-9)

    test_eval = test_df[["Vendor"]].copy()
    test_eval["actual"] = actual
    test_eval["pred"] = preds
    vendor_totals = test_eval.groupby("Vendor", observed=True)[["actual", "pred"]].sum()
    vendor_bias_correction = {}
    for vendor, row in vendor_totals.iterrows():
        if row["actual"] >= MIN_VENDOR_TEST_VOLUME:
            raw_factor = row["actual"] / max(row["pred"], 1e-9)
            clamped = min(
                max(raw_factor, global_bias_correction / MAX_DEVIATION_FROM_PARENT),
                global_bias_correction * MAX_DEVIATION_FROM_PARENT,
            )
            vendor_bias_correction[vendor] = clamped
        else:
            vendor_bias_correction[vendor] = global_bias_correction

    sku_cols_no_product = [c for c in SKU_COLS if c != "Product"]
    test_eval_sku = test_df[sku_cols_no_product].copy()
    test_eval_sku["actual"] = actual
    test_eval_sku["pred"] = preds
    sku_totals = test_eval_sku.groupby(sku_cols_no_product, observed=True)[["actual", "pred"]].sum()
    sku_bias_correction = {}
    for keys, row in sku_totals.iterrows():
        vendor = keys[0] if isinstance(keys, tuple) else keys
        fallback = vendor_bias_correction.get(vendor, global_bias_correction)
        if row["actual"] >= MIN_SKU_TEST_VOLUME:
            raw_factor = row["actual"] / max(row["pred"], 1e-9)
            clamped = min(
                max(raw_factor, fallback / MAX_DEVIATION_FROM_PARENT), fallback * MAX_DEVIATION_FROM_PARENT
            )
            sku_bias_correction[keys] = clamped
        else:
            sku_bias_correction[keys] = fallback

    print(f"Test MAE:   {mae:.2f}")
    print(f"Test RMSE:  {rmse:.2f}")
    print(f"Test WMAPE: {wm:.2%}")
    print(f"Naive (7-day rolling mean) WMAPE: {naive_wm:.2%}")
    print(f"Global bias correction factor (should be close to 1.0 now): {global_bias_correction:.3f}")
    print("Per-vendor bias correction factors:")
    for vendor, factor in vendor_bias_correction.items():
        print(f"  {vendor}: {factor:.3f}")

    importance = pd.Series(
        model.feature_importance(importance_type="gain"), index=FEATURE_COLS
    ).sort_values(ascending=False)
    print("\nTop feature importance (gain):")
    print(importance.head(10))

    os.makedirs(MODEL_DIR, exist_ok=True)
    model_path = os.path.join(MODEL_DIR, "demand_model.txt")
    model.save_model(model_path)
    joblib.dump(cat_cols, os.path.join(MODEL_DIR, "cat_cols.joblib"))
    bias_correction = {
        "global": global_bias_correction,
        "by_vendor": vendor_bias_correction,
        "by_sku": sku_bias_correction,
    }
    joblib.dump(bias_correction, os.path.join(MODEL_DIR, "bias_correction.joblib"))
    print(f"\nModel saved to {model_path}")


if __name__ == "__main__":
    main()
