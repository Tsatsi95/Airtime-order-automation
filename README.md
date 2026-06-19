# Flash Group – Airtime Order Automation

> **Cellular Team internal tool** · Built with Python + Streamlit

Automates the end-to-end airtime ordering workflow for the Flash Group Cellular team. Upload client Purchase Order PDFs, check live stock-on-hand, and get network order templates pre-filled and ready to send — in one click.

---

## The Problem It Solves

Every ordering cycle the team had to:
1. Manually read each client PO PDF
2. Cross-reference quantities against the VVS stock-on-hand report
3. Calculate the shortfall for each product
4. Manually type quantities into the MTN and Telkom order templates
5. Draft emails to the network suppliers

This took significant time and was prone to transcription errors. This tool collapses all five steps into a single automated workflow.

---

## Features

- **PDF parsing** — Extracts line items (product + quantity) from VPS purchase order PDFs, handling space-separated thousands (`4 300` → 4 300)
- **SOH lookup** — Aggregates stock-on-hand from the VVS Excel report (summing multiple batches of the same product)
- **Shortfall calculation** — Computes `order qty = PO qty − SOH` per product
- **Fuzzy product matching** — Two-hop matching (PO description → SOH product → template row) using `rapidfuzz` with size/period/content-type aware scoring
- **Template population** — Writes order quantities directly into the MTN `.xlsm` and Telkom `.xlsx` order templates, preserving all formulas and formatting
- **Review & adjust** — Editable results table so the team can override any quantity or fix a wrong match before downloading
- **Email drafts** — Auto-generated email bodies per network, ready to paste
- **Multi-network** — Upload MTN and Telkom POs simultaneously; only the relevant templates are required

---

## Tech Stack

| Library | Purpose |
|---|---|
| `streamlit` | Web UI |
| `pdfplumber` | PDF text extraction |
| `openpyxl` | Read/write Excel templates |
| `pandas` | Data manipulation |
| `rapidfuzz` | Fuzzy string matching |

---

## Setup

### Prerequisites

- Python 3.10 or newer — [python.org/downloads](https://www.python.org/downloads/)
  - During install, tick **"Add Python to PATH"**

### Install

Open a terminal in this folder and run:

```bash
pip install -r requirements.txt
```

### Run

```bash
streamlit run app.py
```

Your browser will open at `http://localhost:8501`. Press **Ctrl+C** to stop.

---

## Usage

### 1 — Upload files (sidebar)

| File | What it is |
|---|---|
| **SOH Report** | VVS stock-on-hand Excel (`VVS - SOH Report *.xlsx`) |
| **MTN Order Template** | Flash Logical Catalogue `.xlsm` *(only needed for MTN POs)* |
| **Telkom Order Template** | Airtime Order Form `.xlsx` *(only needed for Telkom POs)* |
| **Client PO PDFs** | One or more VPS purchase order PDFs |

The app auto-detects the network from the filename. You can override it with the dropdown.

### 2 — Analyse Orders

Click **Analyse Orders**. The app extracts every line item, checks SOH, calculates shortfalls, and matches each product to the correct row in the order template.

### 3 — Tabs

| Tab | What you do |
|---|---|
| **Stock Analysis** | Review the full results. Green = needs ordering, Yellow = low-confidence match |
| **Review & Adjust** | Edit quantities or fix wrong matches. Click **Save changes** when done |
| **Download Templates** | Download the completed order templates + copy email drafts |

---

## Confidence Scores

| Score | Meaning |
|---|---|
| 80 – 100 | High confidence — match is almost certainly correct |
| 70 – 79 | Good — worth a quick glance |
| 55 – 69 | Low — review recommended before sending |
| < 55 | Very low — manually assign in the Review tab |

---

## Project Structure

```
Automation - ordering/
├── app.py               # Main Streamlit application
├── requirements.txt     # Python dependencies
├── README.md            # This file
├── PLANNING.md          # Architecture and planning notes
└── SETUP.md             # Quick setup guide for the team
```

---

## Contributing

This is an internal team tool. To suggest changes or report a bug, raise it with the Cellular team directly.

---

*Built for Flash Group Cellular Team*
