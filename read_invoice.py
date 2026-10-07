import pdfplumber

with pdfplumber.open("invoices/invoice1.pdf") as pdf:
    page = pdf.pages[0]
    text = page.extract_text()

print(text)