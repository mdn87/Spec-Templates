@echo off
echo SpecConverter v0.4 - Specification Content Extractor
echo ==================================================

if "%1"=="" (
    echo Usage: run_extraction.bat [document.docx] [template.docx]
    echo.
    echo Examples:
    echo   run_extraction.bat examples\SECTION 26 05 00.docx
    echo   run_extraction.bat examples\SECTION 26 05 00.docx templates\test_template_cleaned.docx
    echo.
    pause
    exit /b 1
)

echo Extracting content from: %1

if "%2"=="" (
    echo No template specified, will auto-detect cleaned template...
    python src\extract_spec_content_v3.py "%1"
) else (
    echo Using template: %2
    python src\extract_spec_content_v3.py "%1" . "%2"
)

echo.
echo Extraction complete! Check the output directory for results.
pause 