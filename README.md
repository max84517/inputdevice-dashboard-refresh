# InputDevice Dashboard Refresh

A Python desktop tool that consolidates scattered **Spending & Rebate shipment reports** from multiple GTK suppliers into clean Excel files for Power BI consumption.

## Features

- **Dark-mode Tkinter UI** — browse the base folder, select suppliers with checkboxes, and navigate between file versions with ◀▶ buttons
- **Persistent settings** — base path and supplier selection are saved to `config.json` and restored on next launch
- **Smart file detection** — automatically picks the most recently modified Excel inside each supplier's `Spending and rebate` folder
- **Wide → long pivot** — reads `FYXX` sheets (header on row 2), melts value columns (`Table Price`, `Unit Rebate`, `Q'ty`, `Rebate Amount`) into a long format per month
- **Derived columns** — computes `HP Cost`, `ODM Cost`, `Spending Amount`, `Actual Spending`, `Month`, `Year`, `FY` (e.g. `FY26 Q2`)
- **Merged-cell handling** — forward-fills feature columns to fill gaps from merged cells
- **Data quality** — drops rows where `Platforms` is blank, fills missing value cells with `0`, removes FY/month/empty/supplier-specific columns
- **History files** — saves each processed FY as `FYXX_Rebate & Spending Shipment Report.xlsx` under `data/history/`
- **Power BI export** — merges selected FY history files into `data/output/InputDevice_Shipment.xlsx`
- **Real-time log panel** — progress messages appear live as each supplier is processed

## Project Structure

```
inputdevice-dashboard-refresh/
├── main.py                          # Entry point
├── pyproject.toml                   # Poetry package definition
├── poetry.toml                      # Poetry local venv config
├── poetry.lock
├── config.json                      # Auto-generated; stores UI state
├── data/
│   ├── source_data/                 # Copied source Excels (cleared on each run)
│   ├── history/                     # FYXX_Rebate & Spending Shipment Report.xlsx
│   └── output/                      # InputDevice_Shipment.xlsx (Power BI)
└── src/
    ├── config.py                    # load/save config.json
    ├── ingestion/
    │   └── fetcher.py               # Scan supplier folders, find latest Excel
    ├── processing/
    │   └── transformer.py           # Read sheet, wide→long, derive columns
    ├── export/
    │   └── exporter.py              # Save history, merge for Power BI
    └── ui/
        └── app.py                   # Tkinter dark-mode UI
```

## Requirements

- Python ≥ 3.11
- [Poetry](https://python-poetry.org/) ≥ 1.8

## Setup

```bash
# Install dependencies (creates .venv in project folder)
poetry install

# Run the app
poetry run python main.py
```

## Expected Folder Structure (source data)

```
<Base Folder>/
├── Acrox/
│   └── Spending and rebate/
│       └── Rebate and Spending report Acrox_2026 03.xlsx
├── Chicony/
│   └── Spending and rebate/
│       └── Rebate and Spending report Chicony_2026 03.xlsx
└── ...
```

## Workflow

1. **Browse** — select the root folder containing supplier sub-folders
2. **Select** — tick the suppliers to include; use ◀▶ to pick a different file version
3. **Fetch Data** — scans each supplier's `Spending and rebate` folder and displays the latest file
4. **Process** — choose the common FY sheet → data is transformed and saved to `data/history/`
5. **Power BI Export** — select which FY files to merge → outputs `data/output/InputDevice_Shipment.xlsx`
