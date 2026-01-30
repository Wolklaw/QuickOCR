@echo off
echo Building QuickOCR...
pip install pyinstaller
py -m PyInstaller --noconsole --onefile --name "QuickOCR" --icon="aa.ico" --add-data "Tesseract-OCR;Tesseract-OCR" --add-data "aa.ico;." quickocr.py
echo.
echo Build Complete! Check the 'dist' folder.
pause
