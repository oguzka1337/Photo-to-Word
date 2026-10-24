@echo off
setlocal
cd /d "%~dp0"

if exist "PhotoToWord.exe" (
    start "" "PhotoToWord.exe"
    exit /b 0
)

set "PY=python"
where python >nul 2>&1
if errorlevel 1 set "PY=py"

set "PYW=pythonw"
if /i "%PY%"=="py" set "PYW=pyw"

%PY% -c "import PIL, docx, customtkinter" >nul 2>&1
if errorlevel 1 (
    echo Installing dependencies...
    %PY% -m pip install --user --disable-pip-version-check Pillow python-docx customtkinter
    if errorlevel 1 (
        echo.
        echo Could not install dependencies. Check that Python and pip are available.
        pause
        exit /b 1
    )
)

start "" %PYW% "photo_to_word.py"
