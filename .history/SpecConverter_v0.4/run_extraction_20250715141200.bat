@echo off
echo SpecConverter v0.4 - Specification Content Extractor
echo ==================================================

if "%1"=="" (
    echo Usage: run_extraction.bat [document.docx]
    echo.
    echo Examples:
    echo   run_extraction.bat examples\SECTION 26 05 00.docx
    echo.
    echo Note: Using fixed template: templates\test_template_cleaned.docx
    echo.
    pause
    exit /b 1
)

echo Extracting content from: %1
echo Using template: templates\test_template_cleaned.docx

python src\extract_spec_content_v3.py "%1"

echo.
echo Extraction complete! Check the output directory for results.
pause 