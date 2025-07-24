@echo off
echo Installing dependencies...
echo pandas
python -m pip install --quiet pandas
echo PyMuPDF
python -m pip install --quiet PyMuPDF
echo openpyxl
python -m pip install --quiet openpyxl

echo.
echo Running the script...
python autodxcc.py

pause
