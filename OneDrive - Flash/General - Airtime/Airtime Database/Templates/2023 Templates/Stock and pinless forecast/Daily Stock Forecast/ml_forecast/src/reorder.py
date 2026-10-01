"""Turn the demand forecast into a reorder-quantity recommendation per SKU.

Business constraint: holding excess stock hurts the stock-aging report,
inventory turnover, and the working capital cycle. Orders happen once a
month and must stay within that calendar month - not a rolling 31 days that
spills into the next order cycle. So the horizon is however many days are
left in the current month: a full ~31 days if run at month start, shrinking
as the month goes on, so the order always lands exactly at that month's end
and never overlaps with next month's order:

    horizon      = days remaining in the current calendar month
    order_up_to  = forecast demand over that horizon
    reorder_qty  = max(order_up_to - current stock position, 0)
"""
import calendar
import os

import pandas as pd

from ingest import SKU_COLS
from predict import forecast, load_model
from soh import load_soh

DATA_DIR = os.path.join(os.path.dirname(__file__), "..", "data")
OUTPUT_DIR = os.path.join(os.path.dirname(__file__), "..", "output")


def days_to_real_month_end(as_of: pd.Timestamp, today: pd.Timestamp = None) -> int:
    """Days to forecast forward from `as_of` through the end of the CURRENT
    real-world month, determined by `today` (actual calendar date), not by
    `as_of`'s own month. The underlying order data is typically a day or so
    behind - if that lag happens to straddle a month boundary (e.g. today is
    the 1st but the last available data is the 30th of last month), anchoring
    to `as_of`'s month would wrongly target "1 day left in last month"
    instead of the ~31 days actually needed to cover the new month.
    """
    if today is None:
        today = pd.Timestamp.now().normalize()
    last_day_of_month = calendar.monthrange(today.year, today.month)[1]
    end_of_month = pd.Timestamp(year=today.year, month=today.month, day=last_day_of_month)
    return (end_of_month - as_of).days


def main():
    feat = pd.read_parquet(os.path.join(DATA_DIR, "features.parquet"))
    model, _, bias_correction = load_model()

    as_of = feat["date"].max()
    horizon = max(days_to_real_month_end(as_of), 1)
    if horizon > 45:
        print(f"WARNING: horizon is {horizon} days, which is unusually long for a "
              f"within-month forecast - the underlying order data (as of {as_of.date()}) "
              f"may be significantly stale. Consider re-running ingest.py first.")

    daily_forecast = forecast(feat, model, horizon=horizon, bias_correction=bias_correction)

    sku_cols = [c for c in SKU_COLS if c != "Product"]
    forecast_total = (
        daily_forecast.groupby(sku_cols, as_index=False)["forecast_quantity"]
        .sum()
        .rename(columns={"forecast_quantity": "forecast_total"})
    )

    soh = load_soh()
    plan = forecast_total.merge(soh, on=sku_cols, how="left", indicator="_soh_match")
    plan["has_soh_data"] = plan["_soh_match"] == "both"
    plan = plan.drop(columns="_soh_match")

    plan["order_up_to"] = plan["forecast_total"]
    # SKUs outside the VVS_FORECAST SOH sheet (vouchers, other non-cellular lines) have no
    # stock-position data here, so we can't compute a reorder qty — leave it blank rather
    # than assuming zero stock, which would overstate the reorder need.
    plan["reorder_qty"] = (plan["order_up_to"] - plan["Total SOH"]).clip(lower=0)
    plan.loc[~plan["has_soh_data"], "reorder_qty"] = pd.NA

    plan = plan.sort_values("reorder_qty", ascending=False, na_position="last")

    cols = ["Vendor", "Product Type", "FaceValue", "ProductDescription", "Product Code",
            "Total SOH", "forecast_total", "order_up_to", "reorder_qty", "has_soh_data"]
    plan = plan[cols]

    unmatched = (~plan["has_soh_data"]).sum()
    if unmatched:
        print(f"Note: {unmatched} SKUs are outside the VVS_FORECAST SOH sheet's coverage "
              f"(non-cellular product lines) — reorder_qty left blank for these, forecast_total still shown")

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    out_path = os.path.join(OUTPUT_DIR, "reorder_plan.xlsx")
    plan.to_excel(out_path, sheet_name="Reorder Plan", index=False)

    today = pd.Timestamp.now().normalize()
    print(f"Assumption: forecasting {horizon} day(s) from {as_of.date()} through the end of "
          f"{today.strftime('%B %Y')}, no safety-stock buffer - order sized to land exactly at month-end")
    matched = plan[plan["has_soh_data"]]
    print(f"Total recommended reorder across {len(matched)} matched SKUs: {matched['reorder_qty'].sum():,.0f} units")
    print(f"Saved to {out_path}")


if __name__ == "__main__":
    main()
