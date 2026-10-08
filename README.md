# Invoice PDF to Excel

A Python script that reads invoices (text PDFs, scanned PDFs and photos) and saves the key data to an Excel file.

## Extracted fields

Invoice number, date, company, subtotal, tax, total, currency (USD, EUR, GBP).

## Setup (Windows)

1. Install Tesseract OCR (Windows installer from the UB Mannheim page) to the default folder `C:\Program Files\Tesseract-OCR`.
2. Run:

```
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

## Run

Put PDF, JPG or PNG invoices into the `invoices` folder, then:

```
.\.venv\Scripts\python.exe extract_invoice.py
```

The result is saved as `invoices.xlsx`.

## Limitations

- English invoices only.
- Results from photos depend on image quality: always check the numbers.
- Supports common invoice layouts; new layouts may need small adjustments.