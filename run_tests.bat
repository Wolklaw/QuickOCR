@echo off
echo Running QuickOCR tests...
echo NOTE: these open real windows and take the foreground for about 20 seconds.
echo.
py -m unittest discover -s tests -t tests -v
pause
