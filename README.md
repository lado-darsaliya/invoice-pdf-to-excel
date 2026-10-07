# Invoice PDF to Excel

A Python script that reads text-based PDF invoices and saves the key data to an Excel file.

## Extracted fields
Invoice number, date, company, subtotal, tax, total, currency (USD, EUR, GBP).

## Setup (Windows)
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt

## Run
Put PDF invoices into the `invoices` folder, then:
.\.venv\Scripts\python.exe extract_invoice.py

The result is saved as `invoices.xlsx`.

## Limitations
- Works with text-based PDFs only (not scanned images).
- Supports common invoice layouts; new layouts may need small adjustments.