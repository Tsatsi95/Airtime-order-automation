# Airtime Stock Forecast & Reorder System — Handover

**Written:** September 2026, ahead of Tsatsi Letsoara's last day (16 October 2026).

## What this is

A machine-learning forecast that predicts how much of each airtime/data SKU
will sell before month-end, compares that to current stock (`VVS_FORECAST.xlsx`),
and tells you how much to reorder — sized to land at exactly enough stock to
cover the rest of the month, not more. Over-ordering hurts stock aging,
inventory turnover, and working capital, so the system deliberately does not
add a safety-stock buffer (see "Business logic decisions" below).

Output: [`output/priority_order.xlsx`](output/priority_order.xlsx), a sorted
table of every SKU by urgency (Stocked out → Critical → Low → OK) with the
recommended reorder quantity. This also gets emailed to the team automatically
twice a day (see "Scheduled automation").

## Quick start — running it manually

From `ml_forecast/src`, in order:

```powershell
python ingest.py         # pulls BOD - Order Report files, builds daily demand per SKU
python features.py       # lag/rolling/calendar features + reseller-concentration signals
python train.py          # retrains the LightGBM model (~5-15 min)
python reorder.py        # forecasts demand, joins to current SOH, computes reorder_qty
python priority_order.py # builds the top-15 sorted, color-coded priority sheet
```

`reorder.py`/`priority_order.py` alone (skipping retrain) take under a minute and
are enough for a same-day refresh if you just refreshed SOH and don't need a new
model. Before running any of these, refresh the SOH workbook:

```powershell
powershell -ExecutionPolicy Bypass -File ..\refresh_soh.ps1
```

(Run `refresh_soh.ps1` from the `ml_forecast` folder — it opens
`VVS_FORECAST.xlsx` via Excel COM and does `RefreshAll`. **Close
`VVS_FORECAST.xlsx` yourself first if you have it open** — the automation
can't open a file you already have locked, and fails with a hard-to-read COM
error if you don't.)

To email the team's priority table (built from whatever's currently in
`output/priority_order.xlsx`):

```powershell
python src\team_stockout_email.py     # builds output\_alert_table.html + prints counts
powershell -ExecutionPolicy Bypass -File send_prebuilt_alert.ps1   # sends via Outlook
```

(`send_prebuilt_alert.ps1` has the stocked-out/critical counts hardcoded into
its subject line — update that line to match what `team_stockout_email.py`
printed before sending, or just use `team_stockout_alert.ps1` below, which
does the whole refresh+send in one go.)

## Architecture

```
ml_forecast/
  src/
    config.py          # <- resolves OneDrive root + all data paths, read this first
    ingest.py           # BOD - Order Report -> daily_demand.parquet
    market_share.py     # Market share.xlsx (per-client orders) -> client_signal_features.parquet
    features.py         # daily_demand + client_signal -> features.parquet
    train.py             # features.parquet -> demand_model.txt + bias_correction.joblib
    predict.py           # recursive multi-day forecast (used by reorder.py)
    reorder.py           # forecast + SOH -> reorder_plan.xlsx
    priority_order.py    # reorder_plan.xlsx + live SOH -> priority_order.xlsx (top 15)
    soh.py                # reads VVS_FORECAST.xlsx SOH columns
    team_stockout_email.py  # priority_order.xlsx -> output\_alert_table.html
  data/                 # generated parquet files (gitignored)
  models/               # trained model + bias correction (gitignored)
  output/               # reorder_plan.xlsx, priority_order.xlsx (gitignored)
  logs/                 # every scheduled/manual run writes a timestamped log here
  paths.ps1             # PowerShell equivalent of config.py - dot-source this
  run_pipeline.ps1       # full monthly pipeline (ingest->features->train->reorder->priority_order)
  refresh_soh.ps1        # opens VVS_FORECAST.xlsx, RefreshAll, saves
  team_stockout_alert.ps1  # refresh SOH + reorder + priority_order + email, used by daily tasks
  send_monthly_reminder.ps1  # popup reminder, used by PriorityOrderReminder_Monthly
  send_prebuilt_alert.ps1    # manual one-off: email whatever's in output\_alert_table.html
  watchdog_alert.ps1      # checks today's alert actually sent; pops up if not
  watchdog_monthly.ps1    # checks this month's pipeline actually completed; pops up if not
  requirements.txt
  HANDOVER.md            # this file
```

`config.py` (Python) and `paths.ps1` (PowerShell) are the single source of
truth for every file path — they find the current user's `OneDrive - Flash`
folder automatically rather than hardcoding a Windows username. **If you ever
see a hardcoded `C:\Users\...` path in a script that isn't one of these two
files, that's a bug — route it through config.py/paths.ps1 instead.**

## Business logic decisions (and why)

- **No safety-stock buffer.** Early versions added a statistical safety-stock
  margin on top of the forecast. The business explicitly rejected this:
  excess stock hurts stock aging, inventory turnover, and working capital.
  `reorder_qty = max(forecast_total - Total SOH, 0)`, nothing added.
- **Horizon = days remaining in the current calendar month, not a rolling
  31 days.** Orders happen once a month; a rolling 31-day horizon would
  spill into next month's own order and cause overlap/over-purchase. See
  `days_remaining_in_month()` in `reorder.py`.
- **Internal stock transfers excluded from training data.** The BOD Order
  Report includes bulk "orders" that are actually stock moved between
  internal accounts (originally under "Flash Trader VS"), not real customer
  demand. Two patterns are filtered in `ingest.py`: rows with `"Transfer"` in
  the User Reference, and round-10,000-unit batches under that account
  (confirmed with the business as also being transfers).
- **`Status = "Complete"` only.** The BOD Order Report Power Query in
  `VVS_FORECAST.xlsx` was including Canceled/New orders in "QTY Processed" —
  fixed by adding a filter step (`Filtered Rows3`) to that query.
- **`AVG_31 days` / `AVG_7 days` measures** in `VVS_FORECAST.xlsx` mean
  "average quantity **per transaction**", not "per day" (business request —
  see git history / conversation log for the exact DAX). **`Days of SOH`
  was rewired** to compute its own true daily rate directly from
  `QTY Processed Last 31/7 Days` instead of dividing by those AVG measures,
  since after the redefinition dividing by them would have silently broken
  the "days of stock" meaning. If anyone ever asks "why doesn't Days of SOH
  just use AVG_31/AVG_7" — this is why.
- **Reseller concentration/recency features** (`market_share.py`,
  `days_since_last_client_order`, `top_client_share_Nd`,
  `n_active_clients_Nd`). Large resellers (PEP Stores, Ackermans, etc.) place
  periodic bulk orders that create demand spikes the base model otherwise
  smooths over. This meaningfully improved the model — `days_since_last_client_order`
  usually ranks in the top-10 most important features.
- **Bias correction, computed per-SKU with fallback to per-vendor then
  global.** The LightGBM model (Tweedie objective) can still be slightly
  biased on the held-out test period; `train.py` computes a correction
  factor at SKU/vendor/global granularity (whichever has enough held-out
  volume to be reliable) and `predict.py` applies it. Check the "Global bias
  correction factor" printed by `train.py` — it should be close to 1.0
  (0.9-1.1ish); if it drifts far from that, something about the underlying
  demand pattern has probably shifted and is worth investigating before
  trusting the reorder numbers.

## Scheduled automation

Seven Windows Task Scheduler tasks, all under this Windows user's Task
Scheduler library:

| Task | Schedule | Does |
|---|---|---|
| `AirtimeStockForecast_Monthly` | 1st of month, 06:00 | Full pipeline: ingest→features→train→reorder→priority_order |
| `PriorityOrderReminder_Monthly` | 1st of month, 08:00 | Popup reminder to check the sheet |
| `TeamStockoutAlert_Morning` | Weekdays, 08:00 | Refresh SOH + reorder + priority_order + email team |
| `TeamStockoutAlert_Afternoon` | Weekdays, 15:30 | Same as above |
| `Watchdog_MorningAlert` | Weekdays, 08:20 | Popup if the 08:00 alert didn't actually send |
| `Watchdog_AfternoonAlert` | Weekdays, 15:45 | Popup if the 15:30 alert didn't actually send |
| `Watchdog_MonthlyPipeline` | 2nd of month, 09:00 | Popup if the monthly pipeline hasn't completed yet |

All have `StartWhenAvailable` enabled (catches up a missed trigger next time
the machine is logged in, rather than skipping the month/day entirely).

### ⚠️ Known reliability issue — read this before assuming silence means success

**Over a real 4-week trial (mid-August to mid-September 2026), the Outlook/
Excel-COM-dependent tasks (`TeamStockoutAlert_*`, `AirtimeStockForecast_Monthly`)
succeeded only about 30-40% of the time.** Failure modes observed:

1. **Instant kill** (`STATUS_CONTROL_C_EXIT`, exit code `-1073741510`) — the
   process dies before even opening Excel. Root cause not fully identified;
   suspected to be Task Scheduler's interactive-session launch path being
   flaky on this machine (possibly Azure-AD-join-related).
2. **Stuck mid-refresh** (`0x800AC472` / `RPC_E_CALL_REJECTED` on `Quit()`) —
   Excel COM gets busy/rejects the call, usually because **`VVS_FORECAST.xlsx`
   was already open by a human** at trigger time. This one has a known,
   avoidable cause: don't have the workbook open at 08:00/15:30.
3. **Task simply never fires at all** — no log entry whatsoever for that day,
   not even a failure. Happened on ~8-9 weekdays out of ~20 in the trial period.
4. **Orphaned Excel processes.** When (1) or (2) happens, a hidden `EXCEL.EXE
   /automation -Embedding` process is often left running indefinitely — found
   one that had been running for 3 days. These don't clean themselves up.
   Check periodically:
   ```powershell
   Get-CimInstance Win32_Process -Filter "Name='EXCEL.EXE'" | Select ProcessId, CommandLine, CreationDate
   ```
   Kill anything with `/automation -Embedding` in the command line that's
   older than a few minutes and isn't a legitimate refresh in progress.

**The popup-only tasks (`PriorityOrderReminder_Monthly`, `Watchdog_*`) have
been 100% reliable** — they don't touch Outlook or Excel COM, just
`System.Windows.Forms.MessageBox`. If you're ever adding new automation,
prefer a popup over an email/Excel-COM approach if the popup can do the job.

**Practical implication:** don't assume the scheduled tasks are working.
Check `logs/` periodically (`alert_YYYY-MM-DD_*.log`, `run_YYYY-MM_*.log`) —
a log ending after just the first line (`=== ... run: ... ===`) means it died
instantly; one that stops after "Opened workbook, starting RefreshAll..."
died mid-refresh; one ending with "Email sent to: ..." or "=== Done: ..."
succeeded. The watchdog tasks help catch this via popup, but only if someone
is logged in to see the popup at 08:20/15:45.

### Locale bug worth knowing about (if you ever edit DAX/Power Query via script)

This machine's Windows regional settings have both the decimal separator AND
list separator set to `,` — an unusual, ambiguous combination. Setting a DAX
measure's `.Formula` property via Excel COM automation silently corrupts any
bare digit immediately followed by a comma (e.g. `31,` gets mangled into
`31.`, breaking the syntax). The workaround used throughout this project:
wrap numeric literals in parentheses, e.g. `(-31)` instead of `-31`, so a `)`
sits before the comma instead of a digit. If the new machine's regional
settings are normal (`,` decimal / different list separator, or the reverse),
this bug may not reproduce — but if you ever get a cryptic "syntax for 'X' is
incorrect" error setting a Formula via COM, this is why.

## Known risks / things to fix before or during handover

1. **`Market share.xlsx` lives in a personal folder, not the shared library.**
   `config.py`'s `MARKET_SHARE_PATH` resolves to
   `OneDrive - Flash\Documents\Reports\Nontu\Markets Shares\Market share.xlsx`
   — a colleague's (Nontu's) personal OneDrive Documents folder, not
   `General - Airtime`. This will **not** resolve on a new machine/account
   unless the new person has their own copy at the equivalent path. Either:
   - move the file into the shared `General - Airtime` library and update
     `MARKET_SHARE_PATH` in `config.py`, or
   - have the new person's own OneDrive get a copy at the same relative path
     (`OneDrive - Flash\Documents\Reports\Nontu\Markets Shares\Market share.xlsx`),
     which is fragile and not recommended.
   `MARKET_SHARE_PATH` can also be overridden via the `MARKET_SHARE_PATH`
   environment variable without touching code, if that's easier short-term.
2. **Outlook sender identity.** All emails send via whatever Outlook profile
   is configured on the machine currently running the automation. On the new
   machine, this will naturally become the new person's own mailbox — no
   code change needed, just confirm Outlook is set up and signed in there.
3. **Team recipient list** is hardcoded in `team_stockout_alert.ps1` and
   `send_prebuilt_alert.ps1` (`lutho.posiliti@flash.co.za`,
   `Khanyesa.Khota@flash.co.za`, `Nontuthuko.madondo@flash.co.za`). Update if
   the team changes.
4. **Git repository is rooted at the Windows user's entire home directory**
   (`C:\Users\TsatsiLetsoara`), not just this project — an unusual setup from
   before this project started. Only `ml_forecast/` source files (plus
   `.gitignore`) are tracked; everything else under the home directory is
   untracked noise. Remote: `https://github.com/Tsatsi95/Airtime-order-automation`.
   On a new machine, either re-clone that repo somewhere sensible and copy
   `ml_forecast/` into place, or set up a fresh repo scoped to just this
   project — the current home-directory-rooted setup is not something worth
   replicating.

## New-machine setup checklist

1. Confirm OneDrive is installed, signed in, and `General - Airtime` (shared
   library) is synced locally.
2. Install Python 3.12+, then `pip install -r ml_forecast\requirements.txt`.
3. Resolve the `Market share.xlsx` path issue above (item 1) *before* running
   `features.py` — it'll otherwise fail on `MARKET_SHARE_PATH` not found.
4. Confirm Outlook is installed and signed into the new owner's mailbox.
5. Run `python config.py`-equivalent sanity check (or just run `ingest.py`)
   to confirm `config.py` resolves all paths correctly:
   ```powershell
   cd ml_forecast\src
   python -c "import config; print(config.ONEDRIVE_FLASH_ROOT); print(config.BOD_FOLDER)"
   ```
6. Run the full pipeline manually once (`run_pipeline.ps1` or the five
   `python` commands under "Quick start") to confirm everything works
   end-to-end before relying on automation.
7. Re-create the 7 scheduled tasks (see table above for names/schedules/
   times) via `schtasks /Create` pointing at the new machine's copy of these
   `.ps1` files — they won't carry over automatically since Task Scheduler
   tasks are local to each machine. Set `StartWhenAvailable` on each
   (`Set-ScheduledTask -Settings` with `.StartWhenAvailable = $true`).
8. Read the "Known reliability issue" section above and decide whether to
   invest more time hardening the automation, or accept the ~30-40% success
   rate and rely on manual triggering + the watchdog popups as a safety net
   (this was the last state of that decision as of handover).
9. Update the recipient list in `team_stockout_alert.ps1` /
   `send_prebuilt_alert.ps1` if needed.

## Contacts

Stockout alert recipients (as of handover): lutho.posiliti@flash.co.za,
Khanyesa.Khota@flash.co.za, Nontuthuko.madondo@flash.co.za.
