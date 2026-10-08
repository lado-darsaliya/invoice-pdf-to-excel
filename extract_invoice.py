import re
from pathlib import Path

import pdfplumber
import pytesseract
from openpyxl import Workbook
from PIL import Image, ImageOps

WINDOWS_TESSERACT = Path(r"C:\Program Files\Tesseract-OCR\tesseract.exe")
if WINDOWS_TESSERACT.exists():
    pytesseract.pytesseract.tesseract_cmd = str(WINDOWS_TESSERACT)

IMAGE_TYPES = {".jpg", ".jpeg", ".png"}

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
    start_of_line = rf"^\s*(?:{labels})\b"
    for line in lines:
        if re.search(start_of_line, line, re.IGNORECASE):
            value = last_money(line)
            if value:
                return value
    # OCR sometimes glues columns together, so the label may be in the middle of a line
    inside_line = rf"(?<![A-Za-z])(?<!Sub )(?<!Sub)(?:{labels})\b[^$€£\d]*({MONEY})"
    for line in lines:
        match = re.search(inside_line, line, re.IGNORECASE)
        if match:
            return match.group(1).strip()
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


def looks_like_text(line):
    letters = sum(ch.isalpha() for ch in line)
    return letters >= 5 and letters / max(len(line.strip()), 1) >= 0.6


def find_company(lines):
    for i, line in enumerate(lines):
        if line.strip() == "From:" and i + 1 < len(lines):
            return lines[i + 1].split(" Order Number")[0].strip()
    for line in lines:
        name = re.sub(r"\s+(Tax Invoice|Invoice)$", "", line.strip(), flags=re.IGNORECASE)
        if looks_like_text(name) and name.lower() not in GENERIC_TITLES:
            return name
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


def ocr_image(image):
    image = ImageOps.grayscale(image)
    if image.width < 1600:
        scale = 1600 / image.width
        image = image.resize((int(image.width * scale), int(image.height * scale)))
    return pytesseract.image_to_string(image, config="--psm 6")


def read_pdf(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        text = "\n".join(page.extract_text() or "" for page in pdf.pages)
        if len(text.strip()) >= 30:
            return text
        # no selectable text: this is a scan, read it with OCR
        return "\n".join(ocr_image(page.to_image(resolution=300).original) for page in pdf.pages)


def read_file(path):
    if path.suffix.lower() == ".pdf":
        return read_pdf(path)
    return ocr_image(Image.open(path))


if __name__ == "__main__":
    rows = []
    files = [p for p in sorted(Path("invoices").iterdir())
             if p.suffix.lower() == ".pdf" or p.suffix.lower() in IMAGE_TYPES]
    for pdf_path in files:
        try:
            text = read_file(pdf_path)
        except pytesseract.TesseractNotFoundError:
            print("Не найдена программа Tesseract OCR. Установи её (см. инструкцию) и запусти снова.")
            raise SystemExit(1)
        row = parse_invoice(text, pdf_path.name)
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