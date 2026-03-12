# PDF Invoice Extractor

Python script that reads invoice PDF files, extracts the invoice number and date using regular expressions, and automatically generates an Excel report.

## Technologies
- Python
- pdfplumber
- openpyxl
- regex (re)

## How to Run

1. Install dependencies:

pip install -r requirements.txt

2. Place the invoice PDFs inside the `pdf_invoice` folder.

3. Run the script:

python leitorPdf.py

## Output

The script generates an Excel file with:

- Invoice number  
- Invoice date  
- File name  
- Status