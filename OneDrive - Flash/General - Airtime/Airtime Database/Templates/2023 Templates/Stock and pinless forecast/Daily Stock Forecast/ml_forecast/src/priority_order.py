"""Export the top-N reorder-priority SKUs, cross-checked against live SOH, to their own sheet."""
import os

import pandas as pd

from soh import VVS_FORECAST_PATH, SHEET_NAME, HEADER_ROW

OUTPUT_DIR = os.path.join(os.path.dirname(__file__), "..", "output")
TOP_N = 15


def status_for(days_of_soh: float) -> str:
    if days_of_soh <= 0:
        return "Stocked out"
    if days_of_soh < 1:
        return "Critical (<1 day)"
    if days_of_soh < 5:
        return "Low"
    return "OK"


def main():
    plan = pd.read_excel(os.path.join(OUTPUT_DIR, "reorder_plan.xlsx"))
    top = plan[plan["has_soh_data"]].sort_values("reorder_qty", ascending=False).head(TOP_N).copy()

    soh = pd.read_excel(VVS_FORECAST_PATH, sheet_name=SHEET_NAME, skiprows=HEADER_ROW)
    soh = soh.dropna(subset=["ProductCode"])
    soh = soh[soh["Vendor"] != "(blank)"]
    soh = soh.rename(columns={"ProductType": "Product Type", "ProductCode": "Product Code"})
    soh["Product Code"] = soh["Product Code"].astype("int64")

    sku_cols = ["Vendor", "Product Type", "Product Code"]
    merged = top.merge(
        soh[sku_cols + ["Pending imports", "Days of SOH", "QTY Processed Last 31 Days",
                         "QTY Processed Last 7 Days"]],
        on=sku_cols, how="left",
    )
    merged["status"] = merged["Days of SOH"].apply(status_for)

    merged = merged.sort_values(["Days of SOH", "reorder_qty"], ascending=[True, False])

    cols = ["status", "Vendor", "Product Type", "FaceValue", "ProductDescription", "Product Code",
            "Total SOH", "Pending imports", "Days of SOH", "QTY Processed Last 31 Days",
            "QTY Processed Last 7 Days", "forecast_total", "order_up_to", "reorder_qty"]
    merged = merged[cols]

    out_path = os.path.join(OUTPUT_DIR, "priority_order.xlsx")
    with pd.ExcelWriter(out_path, engine="xlsxwriter") as writer:
        merged.to_excel(writer, sheet_name="Priority Order", index=False)
        workbook = writer.book
        worksheet = writer.sheets["Priority Order"]

        critical_fmt = workbook.add_format({"bg_color": "#FFC7CE", "font_color": "#9C0006"})
        low_fmt = workbook.add_format({"bg_color": "#FFEB9C", "font_color": "#9C6500"})
        status_col = cols.index("status")
        worksheet.conditional_format(
            1, status_col, len(merged), status_col,
            {"type": "text", "criteria": "containing", "value": "Stocked out", "format": critical_fmt},
        )
        worksheet.conditional_format(
            1, status_col, len(merged), status_col,
            {"type": "text", "criteria": "containing", "value": "Critical", "format": critical_fmt},
        )
        worksheet.conditional_format(
            1, status_col, len(merged), status_col,
            {"type": "text", "criteria": "containing", "value": "Low", "format": low_fmt},
        )
        for i, col in enumerate(cols):
            width = max(12, min(38, merged[col].astype(str).str.len().max() if len(merged) else 12))
            worksheet.set_column(i, i, width)

    print(f"Saved top {len(merged)} priority SKUs to {out_path}")
    print(merged["status"].value_counts())


if __name__ == "__main__":
    main()
