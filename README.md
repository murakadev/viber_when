# Viber Attachment Auto-Renamer

This script automatically monitors a folder for Viber attachments, takes a screenshot of the Viber conversation when a new attachment is detected, and renames the attachment using information extracted from the screenshot (contact name, transaction details, date/time).

## Features

- Monitors a folder for new Viber attachment downloads
- Brings Viber to the foreground when a new attachment is detected
- Takes a screenshot of the Viber window
- Uses OCR (Optical Character Recognition) to extract text from the screenshot
- Renames attachments using the extracted information
- Saves a copy of the screenshot for reference
- Handles temporary files and ensures files are fully downloaded before processing

## Requirements

- Python 3.6 or higher
- Tesseract OCR (for text recognition)
- Python packages:
  - pyautogui
  - pywin32
  - pytesseract
  - pillow (PIL)
  - opencv-python

## Installation

1. Install Python 3.6 or higher
2. Install Tesseract OCR:
   - Download from [Tesseract OCR for Windows](https://github.com/UB-Mannheim/tesseract/wiki)
   - Install to the default location (or update the path in the script)
3. Run the provided `install_requirements.bat` file to install required Python packages

## Configuration

Before running the script, update the following settings in `viber_file_rename.py`:

1. Tesseract OCR path (if different from default):
   ```python
   pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
   ```

2. Viber window title (if different):
   ```python
   VIBER_WINDOW_TITLE = "Rakuten Viber"
   ```

3. Input and output directories:
   ```python
   folder_to_watch = r"C:\Users\Admin\Documents\ViberDownloads"
   OUTPUT_DIR = r"C:\Users\Admin\Documents\Viber_Attachments"
   ```

## Usage

1. Start Viber
2. Run the script:
   ```
   python viber_file_rename.py
   ```
3. The script will continuously monitor the specified folder for new Viber attachments
4. When a new attachment is detected, the script will:
   - Bring Viber to the foreground
   - Take a screenshot
   - Extract information from the screenshot
   - Rename the attachment using the extracted information
   - Save a copy of the screenshot for reference

## Troubleshooting

- If text extraction is not accurate, try adjusting the regions of interest in the `extract_specific_regions` function
- Make sure Viber is running before starting the script
- Check that the Viber window title matches the one in the script
- Ensure the script has permission to access and modify files in the specified directories

## Notes

- The script is designed to work with Viber on Windows
- Screenshot OCR works best with clear, high-contrast text
- The script may require adjustments based on your specific Viber layout and language settings 