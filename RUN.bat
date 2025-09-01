@echo off
echo Installing dependencies...
echo pandas
python -m pip install --quiet pandas
echo requests
python -m pip install --quiet requests
echo PyMuPDF
python -m pip install --quiet pymupdf

echo.
echo Running the script...
python autodxcc.py

pause