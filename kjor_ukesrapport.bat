@echo off
setlocal

set "SCRIPT_DIR=%~dp0"
set "PYTHON_EXE=%SCRIPT_DIR%.venv\Scripts\python.exe"

if not exist "%PYTHON_EXE%" (
  echo Fant ikke Python i .venv: "%PYTHON_EXE%"
  echo Opprett venv eller juster filen kjor_ukesrapport.bat.
  pause
  exit /b 1
)

pushd "%SCRIPT_DIR%"
"%PYTHON_EXE%" blabla.py --refresh
set "EXIT_CODE=%ERRORLEVEL%"
popd

if not "%EXIT_CODE%"=="0" (
  echo.
  echo Rapportkjoring feilet med exit code %EXIT_CODE%.
  pause
  exit /b %EXIT_CODE%
)

set "HTML_FILE=%SCRIPT_DIR%report_2027.html"
set "PDF_FILE=%SCRIPT_DIR%report_2027.pdf"

if exist "%HTML_FILE%" (
  where wkhtmltopdf >nul 2>&1
  if "%ERRORLEVEL%"=="0" (
    echo.
    echo Lager PDF med wkhtmltopdf...
    wkhtmltopdf --enable-local-file-access "%HTML_FILE%" "%PDF_FILE%"
    if not "%ERRORLEVEL%"=="0" (
      echo Advarsel: Klarte ikke lage PDF med wkhtmltopdf.
    ) else (
      echo PDF klar: "%PDF_FILE%"
    )
  ) else (
    echo.
    echo Hopper over PDF: wkhtmltopdf ble ikke funnet i PATH.
  )
)

echo.
echo Ferdig. Genererte filer: report_2027.html og report_2027.pdf
pause
