# QR Code Generation from Excel File

This script generates QR codes from the data in an Excel file. Each QR code is created for each row in the file and classified into a folder based on the category from the "Prom" column.

## Requirements

Before using this script, you need to install the required packages and have a correctly formatted Excel file.

### Required Packages

You must have Python installed on your machine along with the following packages:

- **pandas**: used to read and manipulate Excel data.
- **qrcode**: used to generate QR codes.
- **openpyxl**: used by `pandas` to read `.xlsx` Excel files.

You can install these packages with `pip` by running the following command:

```bash
pip install pandas qrcode[pil] openpyxl
