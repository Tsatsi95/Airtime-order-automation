"""Load BOD - Order Report files and aggregate to daily quantity per SKU."""
import glob
import os

import pandas as pd

from config import BOD_FOLDER

SKU_COLS = ["Vendor", "Product Type", "Product Code", "Product"]


def load_raw(folder: str = BOD_FOLDER) -> pd.DataFrame:
    files = sorted(glob.glob(os.path.join(folder, "*.xlsx")))
    frames = []
    for f in files:
        df = pd.read_excel(f, skiprows=2)
        df["source_file"] = os.path.basename(f)
        frames.append(df)
    return pd.concat(frames, ignore_index=True)


def build_daily_demand(raw: pd.DataFrame) -> pd.DataFrame:
    df = raw[raw["Status"] == "Complete"].copy()

    # rows like "Transfer_2026/07.2853" are internal stock movement between
    # aggregation accounts, not customer demand — exclude them
    is_labeled_transfer = df["User Reference"].astype(str).str.contains("transfer", case=False, na=False)

    # under "Flash Trader VS", most volume is in exact multiples of 10,000 units
    # with sequential numeric references — the same internal-transfer pattern as
    # the labeled rows above, just without the "Transfer" text. Confirmed with
    # the business as also being internal stock movement, not customer demand.
    is_unlabeled_transfer = (
        (df["Aggregation Account"] == "Flash Trader VS") & (df["Quantity"] % 10000 == 0)
    )

    df = df[~is_labeled_transfer & ~is_unlabeled_transfer]
    df["date"] = pd.to_datetime(df["Order Date"]).dt.normalize()

    daily = (
        df.groupby(SKU_COLS + ["date"], as_index=False)
        .agg(quantity=("Delivered Quantity", "sum"))
    )

    # fill in missing calendar days per SKU with 0 quantity so the model sees true gaps
    skus = daily[SKU_COLS].drop_duplicates()
    full_range = pd.date_range(daily["date"].min(), daily["date"].max(), freq="D")
    date_frame = pd.DataFrame({"date": full_range})

    skus["_key"] = 1
    date_frame["_key"] = 1
    grid = skus.merge(date_frame, on="_key").drop(columns="_key")

    daily_full = grid.merge(daily, on=SKU_COLS + ["date"], how="left")
    daily_full["quantity"] = daily_full["quantity"].fillna(0)
    return daily_full.sort_values(SKU_COLS + ["date"]).reset_index(drop=True)


def filter_active_skus(daily: pd.DataFrame, min_active_days: int = 180) -> pd.DataFrame:
    """Keep only SKUs that have had at least one sale within the trailing window.

    Drops discontinued/legacy products that would otherwise pad the training
    set with long stretches of true-zero demand irrelevant to current stock decisions.
    """
    recent_cutoff = daily["date"].max() - pd.Timedelta(days=min_active_days)
    recent = daily[daily["date"] > recent_cutoff]
    active_sales = recent.groupby(SKU_COLS)["quantity"].sum()
    active_keys = active_sales[active_sales > 0].index
    return daily.set_index(SKU_COLS).loc[active_keys].reset_index()


if __name__ == "__main__":
    raw = load_raw()
    print(f"Loaded {len(raw):,} raw order rows")
    daily = build_daily_demand(raw)
    print(f"Built {len(daily):,} daily SKU-date rows across {daily[SKU_COLS].drop_duplicates().shape[0]} SKUs")

    active = filter_active_skus(daily)
    print(f"Kept {len(active):,} rows across {active[SKU_COLS].drop_duplicates().shape[0]} active SKUs")

    out_path = os.path.join(os.path.dirname(__file__), "..", "data", "daily_demand.parquet")
    active.to_parquet(out_path, index=False)
    print(f"Saved to {out_path}")
