"""Client-level order data from Market share.xlsx, used to derive reseller
concentration/recency signals that supplement the market-wide demand forecast.

Large resellers (PEP Stores, Ackermans, etc.) place periodic bulk orders that
create exactly the demand spikes the base model struggles with. Tracking how
many clients are actively ordering a SKU, how concentrated that demand is in
one big reseller, and how long it's been since a reseller last ordered gives
the model a leading signal for those spikes instead of just smoothing over them.
"""
import pandas as pd

from config import MARKET_SHARE_PATH

CLIENT_WINDOWS = [7, 30]


def load_client_daily(path: str = MARKET_SHARE_PATH) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name="Pivot", header=1)
    df = df.drop(columns=[c for c in df.columns if str(c).startswith("Unnamed")])
    df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
    df = df.dropna(subset=["Date"])

    df = df[df["ProductCode"] != "(blank)"]
    df["ProductCode"] = df["ProductCode"].astype("int64")
    df = df.rename(columns={"Network": "Vendor", "ProductCode": "Product Code",
                             "Clients": "Client", "Date": "date", "Total Quantity": "quantity"})

    daily = (
        df.groupby(["Vendor", "Product Code", "Client", "date"], as_index=False)["quantity"]
        .sum()
    )
    return daily


def build_client_signal_features(daily_by_client: pd.DataFrame) -> pd.DataFrame:
    sku_cols = ["Vendor", "Product Code"]
    all_features = []

    for keys, group in daily_by_client.groupby(sku_cols, observed=True):
        full_range = pd.date_range(group["date"].min(), group["date"].max(), freq="D")
        pivoted = (
            group.pivot_table(index="date", columns="Client", values="quantity", aggfunc="sum")
            .reindex(full_range)
            .fillna(0)
        )

        feat = pd.DataFrame(index=full_range)
        total_daily = pivoted.sum(axis=1)

        for window in CLIENT_WINDOWS:
            rolling = pivoted.shift(1).rolling(window, min_periods=1).sum()
            rolling_total = rolling.sum(axis=1)
            feat[f"n_active_clients_{window}d"] = (rolling > 0).sum(axis=1)
            feat[f"top_client_share_{window}d"] = (
                rolling.max(axis=1) / rolling_total.replace(0, pd.NA)
            ).astype("float64").fillna(0)

        had_order = (total_daily.shift(1).fillna(0) > 0)
        last_order_date = pd.Series(pd.NaT, index=full_range)
        last_order_date[had_order] = pd.Series(full_range, index=full_range)[had_order]
        last_order_date = last_order_date.ffill()
        feat["days_since_last_client_order"] = (
            pd.Series(full_range, index=full_range) - last_order_date
        ).dt.days

        feat["Vendor"] = keys[0]
        feat["Product Code"] = keys[1]
        feat = feat.reset_index().rename(columns={"index": "date"})
        all_features.append(feat)

    return pd.concat(all_features, ignore_index=True)


if __name__ == "__main__":
    import os

    daily = load_client_daily()
    print(f"Loaded {len(daily):,} client-daily rows across "
          f"{daily[['Vendor', 'Product Code']].drop_duplicates().shape[0]} SKUs, "
          f"{daily['Client'].nunique()} clients")

    features = build_client_signal_features(daily)
    print(f"Built client-signal features: {features.shape}")

    out_path = os.path.join(os.path.dirname(__file__), "..", "data", "client_signal_features.parquet")
    features.to_parquet(out_path, index=False)
    print(f"Saved to {out_path}")
