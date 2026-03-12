@echo off
setlocal enabledelayedexpansion

echo Installing required packages...
C:\Users\mikel.huete.ext\AppData\Local\Programs\Python\Python311\python.exe -m pip install pymupdf pdfplumber reportlab Pillow

if errorlevel 1 (
    echo Installation failed!
    exit /b 1
)

echo.
echo Running PDF analysis...
C:\Users\mikel.huete.ext\AppData\Local\Programs\Python\Python311\python.exe C:\Users\mikel.huete.ext\Desktop\Word to Pdf\doc2pdf\pdfCreation\analyze_pdf.py

if errorlevel 1 (
    echo Analysis failed!
    exit /b 1
)

echo.
echo All done!
