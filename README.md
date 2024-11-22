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

# QR Code Scanner with Webcam Integration

This Python script scans QR codes using your computer's webcam, processes the data, and updates an Excel file with scan information. The QR code data is expected to contain "Name" and "First name" fields, and the script checks and updates the scan count for each entry.

## Features

- Real-time scanning of QR codes from a webcam.
- Detects QR codes, decodes their data, and extracts relevant information (Name and First name).
- Keeps track of scanned QR codes and their count in an Excel file.
- Prevents multiple scans of the same QR code within a specified time interval.
- Draws a rectangle around the detected QR codes in the webcam feed.

## Requirements

Before using this script, ensure that you have Python and the necessary dependencies installed.

### Dependencies

You need to install the following Python packages:

- `cv2` (OpenCV): For webcam capture and image processing.
- `pyzbar`: For decoding QR codes.
- `pandas`: For handling Excel files and data processing.
- `numpy`: For image processing tasks.
- `openpyxl`: Required by `pandas` to read/write Excel `.xlsx` files.

To install these packages, run the following command:

```bash
pip install opencv-python pyzbar pandas numpy openpyxl
