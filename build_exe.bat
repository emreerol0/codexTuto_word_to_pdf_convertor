@echo off
setlocal

set "KEEP_WINDOW_OPEN=1"
if /I "%~1"=="--no-pause" set "KEEP_WINDOW_OPEN=0"

set "SCRIPT_DIR=%~dp0"
set "DEFAULT_OUTPUT_DIR=%SCRIPT_DIR%dist"
set /p OUTPUT_DIR=Enter output folder for the executable [default: %DEFAULT_OUTPUT_DIR%]: 
if "%OUTPUT_DIR%"=="" set "OUTPUT_DIR=%DEFAULT_OUTPUT_DIR%"

if not exist "%OUTPUT_DIR%" mkdir "%OUTPUT_DIR%"

set "LOG_FILE=%OUTPUT_DIR%\build.log"
echo Build started at %DATE% %TIME% > "%LOG_FILE%"

echo Upgrading pip... >> "%LOG_FILE%"
python -m pip install --upgrade pip >> "%LOG_FILE%" 2>&1
if errorlevel 1 goto :build_failed

echo Installing project dependencies... >> "%LOG_FILE%"
python -m pip install -r "%SCRIPT_DIR%requirements.txt" >> "%LOG_FILE%" 2>&1
if errorlevel 1 goto :build_failed

echo Installing pyinstaller... >> "%LOG_FILE%"
python -m pip install pyinstaller >> "%LOG_FILE%" 2>&1
if errorlevel 1 goto :build_failed

echo Running pyinstaller... >> "%LOG_FILE%"
pyinstaller --noconfirm --clean --onefile --windowed ^
  --name WordToPdfConverter ^
  --distpath "%OUTPUT_DIR%" ^
  --hidden-import pythoncom ^
  --hidden-import pywintypes ^
  --collect-submodules win32com ^
  "%SCRIPT_DIR%app.py" >> "%LOG_FILE%" 2>&1

if errorlevel 1 goto :build_failed

set "EXE_PATH=%OUTPUT_DIR%\WordToPdfConverter.exe"
if exist "%EXE_PATH%" goto :build_success

for /f "delims=" %%F in ('dir /b /s "%OUTPUT_DIR%\WordToPdfConverter*.exe" 2^>nul') do (
    set "EXE_PATH=%%F"
    goto :build_success
)

for /f "delims=" %%F in ('dir /b /s "%SCRIPT_DIR%dist\WordToPdfConverter*.exe" 2^>nul') do (
    set "EXE_PATH=%%F"
    goto :build_success
)

echo Build finished, but no executable was found.
echo Check the log at "%LOG_FILE%" for details.
call :finalize 1
exit /b 1

:build_success
echo Build complete. Executable is at "%EXE_PATH%"
echo Build complete. Executable is at "%EXE_PATH%" >> "%LOG_FILE%"
call :finalize 0
exit /b 0

:build_failed
echo Build failed. See log for details: "%LOG_FILE%"
call :finalize 1
exit /b 1

:finalize
if "%KEEP_WINDOW_OPEN%"=="1" (
    echo.
    echo Press any key to close this window...
    pause >nul
)
exit /b %~1
