# Excel to Word to PDF Automation

This Python script automates the process of generating PDFs from Excel data by filling a Word template with data from an Excel sheet, then converting it to PDF.
This repository contains a Python script designed to automate the process of generating PDFs from Excel data, which is particularly useful for my current work of automating document generation tasks. The script fills out a Word template using data from an Excel sheet and converts the completed Word documents into PDFs. This automation significantly reduces manual work and improves efficiency in managing bulk document creation.

## Prerequisites

- Python 3.x
- Install required packages by running:
  ```bash
  pip install pandas python-docx docx2pdf


## How To Use
- Update the following variables in generate_pdfs_from_excel.py to match your files:

    - excel_file_path: Path to your Excel file.
    - word_template_path: Path to your Word template file.
    - Adjust column names in the script to match those in your Excel file.
  
## Run The Script
```
python generate_pdfs_from_excel.py
```

The generated PDFs will be saved in the generated_pdfs/ folder.
