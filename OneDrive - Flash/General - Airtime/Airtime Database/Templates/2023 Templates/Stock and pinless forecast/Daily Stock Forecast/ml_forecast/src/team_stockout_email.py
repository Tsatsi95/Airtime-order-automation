"""Build the HTML stockout/critical alert table for the team email."""
import os
import sys

import pandas as pd

from config import ML_FORECAST_DIR

OUTPUT_DIR_DEFAULT = os.path.join(ML_FORECAST_DIR, "output")

STATUS_ICON = {"Stocked out": "\U0001F534 Stocked out", "Critical (<1 day)": "\U0001F534 Critical",
               "Low": "\U0001F7E1 Low", "OK": "\U0001F7E2 OK"}


def short_product(vendor: str, description: str) -> str:
    prefix = f"{vendor} "
    suffix = f" {vendor} Cellular Airtime"
    name = description
    if name.endswith(suffix):
        name = name[: -len(suffix)] + " Airtime"
    if name.startswith(prefix):
        name = name[len(prefix):]
    return name


def build_html_table(priority_path: str) -> tuple:
    plan = pd.read_excel(priority_path)
    urgent = plan[plan["status"].isin(["Stocked out", "Critical (<1 day)"])].copy()

    if urgent.empty:
        return None, 0, 0

    urgent["Product"] = urgent.apply(lambda r: short_product(r["Vendor"], r["ProductDescription"]), axis=1)
    urgent["StatusLabel"] = urgent["status"].map(STATUS_ICON)

    rows_html = ""
    for _, r in urgent.iterrows():
        rows_html += (
            f"<tr>"
            f"<td>{r['StatusLabel']}</td>"
            f"<td>{r['Vendor']}</td>"
            f"<td>{r['Product']}</td>"
            f"<td>{r['Total SOH']:,.0f}</td>"
            f"<td>{r['Days of SOH']:.1f}</td>"
            f"<td>{r['reorder_qty']:,.0f}</td>"
            f"</tr>"
        )

    table_html = f"""
<table border="1" cellpadding="4" cellspacing="0" style="border-collapse:collapse;font-family:Calibri,Arial,sans-serif;font-size:11pt;">
<tr style="background-color:#D9D9D9;font-weight:bold;">
<td>Status</td><td>Vendor</td><td>Product</td><td>Total SOH</td><td>Days of SOH</td><td>Reorder Qty</td>
</tr>
{rows_html}
</table>
"""
    n_stocked_out = (urgent["status"] == "Stocked out").sum()
    n_critical = (urgent["status"] == "Critical (<1 day)").sum()
    return table_html, n_stocked_out, n_critical


if __name__ == "__main__":
    priority_path = sys.argv[1] if len(sys.argv) > 1 else OUTPUT_DIR_DEFAULT + r"\priority_order.xlsx"
    table_html, n_out, n_crit = build_html_table(priority_path)
    print(f"STOCKED_OUT={n_out}")
    print(f"CRITICAL={n_crit}")
    if table_html:
        with open(OUTPUT_DIR_DEFAULT + r"\_alert_table.html", "w", encoding="utf-8") as f:
            f.write(table_html)
        print("TABLE_WRITTEN=1")
    else:
        print("TABLE_WRITTEN=0")
