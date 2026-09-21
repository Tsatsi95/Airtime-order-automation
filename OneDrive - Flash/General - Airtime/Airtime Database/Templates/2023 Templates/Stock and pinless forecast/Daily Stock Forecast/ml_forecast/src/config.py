"""Shared path configuration.

Resolves the OneDrive root dynamically from the current Windows user's home
directory so this project works on any machine/login without a hardcoded
username baked into every script. See HANDOVER.md before moving this project
to a new machine.
"""
import glob
import os


def _find_onedrive_flash_root() -> str:
    override = os.environ.get("ONEDRIVE_FLASH_ROOT")
    if override:
        return override

    home = os.path.expanduser("~")
    exact = os.path.join(home, "OneDrive - Flash")
    if os.path.isdir(exact):
        return exact

    matches = glob.glob(os.path.join(home, "OneDrive*Flash*"))
    if matches:
        return matches[0]

    raise RuntimeError(
        f"Could not find the 'OneDrive - Flash' folder under your home directory ({home}). "
        "Make sure OneDrive is installed and synced, or set the ONEDRIVE_FLASH_ROOT "
        "environment variable to the correct path."
    )


ONEDRIVE_FLASH_ROOT = _find_onedrive_flash_root()

AIRTIME_DB_TEMPLATES = os.path.join(ONEDRIVE_FLASH_ROOT, "General - Airtime", "Airtime Database", "Templates")
DAILY_STOCK_FORECAST_DIR = os.path.join(
    AIRTIME_DB_TEMPLATES, "2023 Templates", "Stock and pinless forecast", "Daily Stock Forecast"
)
ML_FORECAST_DIR = os.path.join(DAILY_STOCK_FORECAST_DIR, "ml_forecast")

BOD_FOLDER = os.path.join(AIRTIME_DB_TEMPLATES, "Reports for templates", "BOD - Order Report")
VVS_FORECAST_PATH = os.path.join(DAILY_STOCK_FORECAST_DIR, "VVS_FORECAST.xlsx")

# NOTE: this file lives in a personal Documents folder (shared by a colleague),
# not the shared "General - Airtime" library - it will NOT resolve for a new
# user unless they have their own copy at the equivalent path, or (better) it
# gets moved into the shared library and MARKET_SHARE_PATH updated to match.
# See HANDOVER.md "Known risks" section.
MARKET_SHARE_PATH = os.environ.get("MARKET_SHARE_PATH") or os.path.join(
    ONEDRIVE_FLASH_ROOT, "Documents", "Reports", "Nontu", "Markets Shares", "Market share.xlsx"
)
