import re
from pathlib import Path

import pdfplumber
from openpyxl import Workbook

MONEY = r"(?:[$€£]\s?\d[\d,]*\.\d{2})|(?:\d[\d,]*\.\d{2}\s?[$€£])"
MONTH = r"(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\.?"
DATE_PATTERNS = [
    rf"{MONTH} \d{{1,2}}, \d{{4}}",
    rf"\d{{1,2}} {MONTH} \d{{4}}",
    r"\d{4}-\d{2}-\d{2}",
    r"\d{1,2}[./-]\d{1,2}[./-]\d{2,4}",
]
CURRENCIES = {"$": "USD", "€": "EUR", "£": "GBP"}
GENERIC_TITLES = {"invoice", "tax invoice"}


def last_money(line):
    found = re.findall(MONEY, line)
    return found[-1].strip() if found else ""


def find_amount(lines, labels):
    pattern = rf"^\s*(?:{labels})\b"
    for line in lines:
        if re.search(pattern, line, re.IGNORECASE):
            value = last_money(line)
            if value:
                return value
    return ""


def find_number(text):
    match = re.search(
        r"Invoice\s*(?:Number|No\.?|#)\s*:?\s*([A-Z0-9][\w\-/]*)", text, re.IGNORECASE
    )
    if match:
        return match.group(1).strip()
    match = re.search(r"\b(INV[-/]?\d[\w\-/]*)", text)
    return match.group(1).strip() if match else ""


def find_date(text):
    candidates = []
    for p in DATE_PATTERNS:
        for m in re.finditer(p, text):
            before = text[max(0, m.start() - 20):m.start()].lower()
            if "due" in before:
                continue
            candidates.append((m.start(), m.group(0)))
    return min(candidates)[1] if candidates else ""


def find_company(lines):
    for i, line in enumerate(lines):
        if line.strip() == "From:" and i + 1 < len(lines):
            return lines[i + 1].split(" Order Number")[0].strip()
    for line in lines:
        if line.strip() and line.strip().lower() not in GENERIC_TITLES:
            return line.strip()
    return ""


def parse_invoice(text, file_name):
    lines = text.split("\n")
    total = find_amount(lines, "Grand Total|Total Due|Amount Due|Total Amount|Total")
    code = re.search(r"\b(USD|EUR|GBP)\b", text)
    symbol = re.search(r"[$€£]", total)
    if code:
        currency = code.group(1)
    elif symbol:
        currency = CURRENCIES[symbol.group(0)]
    else:
        currency = ""
    return {
        "File": file_name,
        "Invoice Number": find_number(text),
        "Date": find_date(text),
        "Company": find_company(lines),
        "Subtotal": find_amount(lines, "Sub ?total|Net amount"),
        "Tax": find_amount(lines, "Sales Tax|Tax|VAT|GST"),
        "Total": total,
        "Currency": currency,
    }


def read_pdf(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        return "\n".join(page.extract_text() or "" for page in pdf.pages)


if __name__ == "__main__":
    rows = []
    for pdf_path in sorted(Path("invoices").glob("*.pdf")):
        row = parse_invoice(read_pdf(pdf_path), pdf_path.name)
        missing = [k for k, v in row.items() if not v]
        if missing:
            print(f"ВНИМАНИЕ: в {pdf_path.name} не найдено: {', '.join(missing)}")
        rows.append(row)

    wb = Workbook()
    ws = wb.active
    ws.append(list(rows[0].keys()))
    for row in rows:
        ws.append(list(row.values()))
    for col in ws.columns:
        width = max(len(str(cell.value or "")) for cell in col) + 2
        ws.column_dimensions[col[0].column_letter].width = width
    wb.save("invoices.xlsx")
    print(f"Готово! Обработано счетов: {len(rows)}. Файл: invoices.xlsx")