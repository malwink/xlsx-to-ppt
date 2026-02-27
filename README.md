# Excel to PowerPoint Automation

A Python script that reads structured data from an Excel workbook and automatically generates a formatted PowerPoint presentation.

This is a simplified sample adapted from a larger project where Excel-based data inputs were programmatically transformed into standardized PowerPoint reports.

---

## Overview

The script:

- Reads configuration settings from the first worksheet
- Reads tabular data from the second worksheet
- Automatically generates:
  - A styled title slide
  - Multiple content slides
  - A final slide containing a formatted table
- Saves the result as a `.pptx` file

This demonstrates automation of structured reporting workflows using Python.

---

## Use Case

This sample reflects a real-world scenario where:

- Business or operational data is maintained in Excel
- Reports must be generated regularly in PowerPoint
- Manual copy-paste is inefficient and error-prone
- Formatting and layout must remain consistent

The full project automated the pipeline from Excel data to presentation-ready output.

---

## Technologies Used

- **Python 3.10+**
- `openpyxl` – Excel file parsing
- `python-pptx` – PowerPoint generation
- Type hints and structured class design
- Basic validation and formatting logic

---

## Project Structure

excel-to-ppt/
│
├── excel_to_ppt.py
├── requirements.txt
├── data.xlsx
├── README.md
└── .gitignore

---

## Excel Structure

The workbook must contain at least two worksheets:

### Sheet 1 – Settings (row 2 used)

| Cell | Purpose |
|------|----------|
| A2 | Number of slides |
| B2 | Title font size (pt) |
| C2 | Title color (hex, e.g. `#1F4E79`) |
| D2 | Body text font size (pt) |

---

### Sheet 2 – Data

The entire worksheet is rendered as a table on the final slide.

---

## Installation & Usage

Create a virtual environment (recommended):

```bash
python -m venv venv
source venv/bin/activate      # macOS/Linux
venv\Scripts\activate         # Windows
pip install -r requirements.txt

---

## Usage

python excel_to_ppt.py

## Or, as import:

from excel_to_ppt import ExcelToPPT

converter = ExcelToPPT("data.xlsx", "presentation.pptx")
converter.create_presentation()