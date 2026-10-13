# Project Planning — Flash Airtime Order Automation

## 1. Problem Statement

The Flash Group Cellular team processes airtime orders from client Purchase Orders (POs) on a recurring basis. Each cycle requires:

- Reading one or more client PO PDFs manually
- Checking current stock-on-hand (SOH) against what the client has ordered
- Calculating the shortfall (what needs to be purchased from the network)
- Typing those quantities into network-specific order templates (MTN, Telkom)
- Drafting supplier emails

This process was entirely manual, time-consuming, and prone to transcription errors. The goal was to automate it end-to-end.

---

## 2. Goals

### Must-have (MVP)
- [x] Parse VPS client PO PDFs automatically
- [x] Load and aggregate stock-on-hand from VVS Excel report
- [x] Calculate shortfall per product
- [x] Match PO products to rows in network order templates (MTN + Telkom)
- [x] Write order quantities into the correct cells in those templates
- [x] Produce downloadable, ready-to-send template files
- [x] Simple web UI accessible to the whole team (no coding required to use)

### Nice-to-have
- [x] Confidence scores on product matches
- [x] Manual override / review table for low-confidence matches
- [x] Auto-generated email drafts per network
- [x] Flash Group branding in the UI
- [ ] Automatic email sending (future)
- [ ] Historical order log / audit trail (future)
- [ ] Vodacom and Cell C template support (future)

---

## 3. Constraints

- **No backend server** — the tool runs locally on a team member's laptop via `streamlit run`. No cloud hosting or IT involvement required.
- **Existing file formats must be preserved** — the MTN `.xlsm` (macro-enabled) and Telkom `.xlsx` templates contain formulas, validation rules, and macros that must not be broken. Only the quantity column is written.
- **PDF format is fixed** — the VPS PO PDFs have a consistent column structure: `QTY  DESCRIPTION  UNIT_PRICE  DISC%  TOTAL`. The parser anchors on the DISC% field (4 decimal places) as a reliable column marker.
- **No external API access** — all data is local (PDFs, Excel files). No network calls.

---

## 4. Architecture

### Pipeline (left to right)

```
Client PO PDFs
      │
      ▼
 PDF Parser          ← pdfplumber, anchors on DISC% column
      │
      ▼
 PO Line Items       {qty, description} per product
      │
      ├──────────────────────────┐
      ▼                          ▼
 SOH Lookup               Template Loader
 (VVS Excel)              (MTN .xlsm / Telkom .xlsx)
      │                          │
      ▼                          ▼
 Shortfall Calc           Template Row Index
      │                          │
      └──────────┬───────────────┘
                 ▼
          Matching Engine         ← rapidfuzz two-hop match
                 │
                 ▼
          Results Table           pandas DataFrame
                 │
          ┌──────┴──────┐
          ▼             ▼
     Streamlit UI   Template Writer  ← openpyxl
     (review/edit)       │
                         ▼
                   Download (.xlsm/.xlsx)
```

### Two-hop matching

Product names differ between the PO, the SOH report, and the order template — sometimes significantly. A direct string match fails too often. The approach:

1. **PO → SOH**: fuzzy match the PO product description against SOH product names to find the SOH qty
2. **PO → Template**: fuzzy match the PO product description against template row labels to find the cell to write

Both hops use the same scoring function.

### Matching score function (`_score`)

Plain `fuzz.token_set_ratio` is too permissive — it ignores important distinguishing features. The scoring function adds:

| Feature | Logic |
|---|---|
| **Size hard-filter** | If both sides have a data size (e.g. `1gb`, `500mb`) and they don't match → score 0 |
| **Rands hard-filter** | Same for rand values (`R50`, `R150`) |
| **Period mismatch penalty** | `daily` vs `weekly` → -20 |
| **Once-off penalty** | Template row is once-off but query isn't → -15 |
| **Content-type penalty** | One side has `tiktok`/`whatsapp`/`freeme` etc., other doesn't → -15 |
| **Voice+ penalty** | Template row adds voice bundles the query doesn't ask for → -15 |
| **Size match bonus** | Exact size match → +15 |
| **Rands match bonus** | Exact rand match → +15 |
| **Minutes match bonus** | Exact minute match → +10 |

Pre-processing:
- `+` replaced with space so `2.5GB+2.5GB` tokenises as two separate sizes
- Size normalisation: `1.0gb` → `1gb` to avoid false mismatches

---

## 5. Key Technical Decisions

### Why Streamlit?
The audience is non-technical team members. Streamlit turns a Python script into a web app with minimal code, runs locally with one command, and needs no deployment infrastructure. The team can run it themselves.

### Why not a macro / VBA solution?
The SOH report and PO PDFs are external files not inside the Excel template. A Python solution handles cross-file reading more cleanly and is easier to maintain and extend.

### Why openpyxl, not xlwings or win32com?
`openpyxl` reads and writes `.xlsx`/`.xlsm` files without needing Excel installed. This keeps the tool portable. The trade-off: macros in `.xlsm` files are preserved but not executed. Since the templates only need their QTY column filled, this is acceptable.

### Why rapidfuzz over standard difflib?
`rapidfuzz` is significantly faster and provides `token_set_ratio` which handles reordered words well (e.g. "MTN 1GB Daily" vs "Daily 1GB MTN Bundle"). Speed matters less here, but the quality of `token_set_ratio` is the key reason.

### File editing workflow
The app file (`app.py`) must **never** be edited directly in the OneDrive-synced folder. OneDrive's sync process can truncate a file mid-write. All edits are made to a local working copy, syntax-verified with `ast.parse()`, then copied to the OneDrive folder.

---

## 6. Data Flow Detail

### PDF Parsing
- Anchor: DISC% column (regex `\d+\.\d{4}`) is always present and unique per line item
- QTY: extracted from the tokens before DESCRIPTION; handles space-thousands-separator (`4 300` = 4300)
- Duplicate descriptions across pages are deduplicated (same product sometimes appears on multiple pages)

### SOH Loading
- VVS report columns: `[Date, ..., Qty(col 3), ..., Product(col 8), ..., Vendor(col 11)]`
- Same product can appear on multiple rows (different batch dates) — quantities are **summed**
- Result: `{vendor: {product_name: total_qty}}`

### Template Loading
- MTN: scans column B for product descriptions; order quantity written to column D
- Telkom: scans column C for product descriptions; order quantity written to column E
- Rows with blank product cells are skipped

### Shortfall Calculation
```
order_qty = max(0, po_qty - soh_qty)
```
If SOH exceeds the PO quantity, order qty is 0 (no order needed).

---

## 7. Confidence Thresholds

Chosen by manual review of match quality across real PO and template data:

| Threshold | Value | Rationale |
|---|---|---|
| `CONFIDENCE_LOW` | 55 | Below this, match is likely wrong — flag in red |
| `CONFIDENCE_WARN` | 70 | Below this, worth a human glance — flag in yellow |

---

## 8. Future Roadmap

### Near-term
- **Vodacom + Cell C templates** — extend `load_*_template_rows()` and `fill_*_template()` for the remaining two networks
- **Persistent match overrides** — save a user's manual corrections so the same product is matched correctly next time
- **Batch history** — log each order run with a timestamp for audit purposes

### Longer-term
- **Automatic email sending** — integrate with Outlook via `win32com` or the Microsoft Graph API to send the completed templates directly
- **SOH threshold alerts** — flag products where SOH is critically low across all POs, not just the current batch
- **Web deployment** — host on an internal server (e.g. Streamlit Community Cloud or a company VM) so the team doesn't need a local Python install

---

## 9. Development Log

| Phase | What was built |
|---|---|
| v0.1 | PDF parser + SOH loader + basic shortfall calculation |
| v0.2 | Fuzzy matching engine (first version, basic token ratio) |
| v0.3 | Template writer for MTN and Telkom; Streamlit UI with three tabs |
| v0.4 | Matching improvements: size hard-filter, period penalty, content-type penalty, `+` normalisation, `1.0gb` normalisation |
| v0.5 | Flash Group branding (logo + colour scheme); optional templates (only require templates for networks present in uploaded POs); session state bug fix for network selectbox |
