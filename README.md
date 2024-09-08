# Excel to Word to PDF Automation

This Python script automates the process of generating PDFs from Excel data by filling a Word template with data from an Excel sheet, then converting it to PDF.
This repository contains a Python script designed to automate the process of generating PDFs from Excel data, which is particularly useful for my current work of automating document generation tasks. The script fills out a Word template using data from an Excel sheet and converts the completed Word documents into PDFs. This automation significantly reduces manual work and improves efficiency in managing bulk document creation.

Use Case:
- Document Automation for Work: I needed to automate repetitive tasks related to generating documents such as invoices, reports, and certificates, where data is already stored in Excel. By merging this data with a Word template, I can quickly generate personalized PDFs.
- Scalable Workflow: The script is capable of handling large datasets, making it ideal for workflows that involve generating hundreds of PDFs from structured data in an Excel file.
  
Features:
- Automatically reads rows from an Excel file and populates a Word template.
- Customizes the Word document based on Excel data, including names, addresses, and dates.
- Converts the populated Word document into a PDF format, eliminating the need for manual conversion.
- Helps streamline repetitive tasks, improving efficiency in daily operations.

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
