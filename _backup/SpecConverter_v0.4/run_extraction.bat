@echo off
echo SpecConverter v0.4 - Specification Content Extractor
echo ==================================================

if "%1"=="" (
    echo Usage: run_extraction.bat [document.docx] [template.docx]
    echo.
    echo Examples:
    echo   run_extraction.bat examples\SECTION 26 05 00.docx templates\test_template_cleaned.docx
    echo.
    echo Available templates:
    echo   - templates\test_template_cleaned.docx (recommended)
    echo   - templates\test_template.docx
    echo   - templates\test_template_orig.docx
    echo.
    pause
    exit /b 1
)

if "%2"=="" (
    echo Error: Template file must be specified as the second argument.
    echo Usage: run_extraction.bat [document.docx] [template.docx]
    echo.
    pause
    exit /b 1
)

echo Extracting content from: %1
echo Using template: %2

python src\extract_spec_content_v3.py "%1" . "%2"

echo.
echo Extraction complete! Check the output directory for results.
pause 