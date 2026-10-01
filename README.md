# Excel Letterhead Generator

A Flask web app that converts Excel workbooks to PDF with LibreOffice and applies a PDF letterhead using PyMuPDF.

## Local run

```bash
python -m venv venv
venv\\Scripts\\activate
pip install -r requirements.txt
py app.py
```

## Render

Deploy as a Docker Web Service on Render. The Docker image installs LibreOffice, so Microsoft Excel is not required on the server.

The service accepts an Excel workbook and a letterhead PDF, then returns the merged PDF.
