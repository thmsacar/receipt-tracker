# Receipt Tracker

CLI tool to scan store receipts using Gemini Flash, track expenses in JSON, and export to Excel.

## Requirements

- Python 3.10+
- Google Gemini API key

## Setup

```bash
pip install -r requirements.txt
cp .env.example .env
# Set your GEMINI_API_KEY in .env
```

## Commands

### Scan a receipt
```bash
python cli.py add /path/to/receipt.jpg
```
Extracts items, normalizes product names, categorizes each item, appends to `receipts.json`, and updates `depenses.xlsx`.

### List receipts
```bash
python cli.py list
```

### View receipt details
```bash
python cli.py show 1
# or by ID
python cli.py show 2026-09-04_01
```

### View stats
```bash
python cli.py stats
```

### Export to Excel
```bash
python cli.py export
```

## Excel Output (`depenses.xlsx`)

Generates 4 sheets:
- Synthèse Globale: KPIs, food vs non-food breakdown, store totals, and charts.
- Articles Alimentaires: Food items table with auto-filters.
- Tickets de Caisse: Chronological receipt log.
- Tous les Articles: Complete item-level log.

## Files

- `cli.py`: Main CLI tool (scan, list, show, stats, export).
- `generate_excel.py`: Excel workbook generator using XlsxWriter.
- `receipts.json`: Local database of parsed receipts.
