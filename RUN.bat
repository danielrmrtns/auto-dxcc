@echo off
echo Installing dependencies...
echo pandas
python -m pip install --quiet pandas
echo requests
python -m pip install --quiet requests
echo fitz
python -m pip install --quiet fitz
echo shutil
python -m pip install --quiet shutil

echo.
echo Running the script...
python autodxcc.py

pause