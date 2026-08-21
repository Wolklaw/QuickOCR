@echo off
echo Building QuickOCR...

py -m pip install -r requirements.txt pyinstaller
if errorlevel 1 goto :failed

rem Stage only the part of Tesseract the app actually needs (94 MB -> ~34 MB).
py make_bundle.py
if errorlevel 1 goto :failed

py -m PyInstaller --noconsole --onefile --name "QuickOCR" --icon="aa.ico" ^
  --add-data "build\bundle\Tesseract-OCR;Tesseract-OCR" ^
  --add-data "aa.ico;." ^
  quickocr.py
if errorlevel 1 goto :failed

echo.
echo Build Complete! Check the 'dist' folder.
goto :end

:failed
echo.
echo BUILD FAILED - see the output above.
exit /b 1

:end
pause
