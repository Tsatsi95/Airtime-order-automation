"""Load current stock-on-hand snapshot from the VVS_FORECAST workbook."""
import pandas as pd

from config import VVS_FORECAST_PATH

SHEET_NAME = "FORECAST|AGGREGATION"
HEADER_ROW = 4  # 5th row in the sheet holds the real column names


def load_soh(path: str = VVS_FORECAST_PATH) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name=SHEET_NAME, skiprows=HEADER_ROW)
    df = df.dropna(subset=["ProductCode"])
    df = df[df["Vendor"] != "(blank)"]

    df = df.rename(columns={"ProductType": "Product Type", "ProductCode": "Product Code"})
    df["Product Code"] = df["Product Code"].astype("int64")

    keep = ["Vendor", "Product Type", "Product Code", "FaceValue", "ProductDescription",
            "SOH", "Pending imports", "Total SOH"]
    return df[keep].drop_duplicates(subset=["Vendor", "Product Type", "Product Code"])
