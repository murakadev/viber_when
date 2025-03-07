@echo off
echo Installing required Python packages for Viber file rename script...
pip install pyautogui pywin32 pytesseract pillow opencv-python
echo.
echo Note: You also need to install Tesseract OCR from:
echo https://github.com/UB-Mannheim/tesseract/wiki
echo.
echo After installing Tesseract OCR, make sure to update the path in the script to match your installation:
echo pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
echo.
pause 