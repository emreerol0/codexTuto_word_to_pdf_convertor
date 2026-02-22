@echo off
setlocal

set "SCRIPT_DIR=%~dp0"
set "DEFAULT_OUTPUT_DIR=%SCRIPT_DIR%dist"
set "KEEP_WINDOW_OPEN=1"
if /I "%~1"=="--no-pause" set "KEEP_WINDOW_OPEN=0"
if /I "%BUILD_NO_PAUSE%"=="1" set "KEEP_WINDOW_OPEN=0"

set /p OUTPUT_DIR=Enter output folder for the executable [default: %DEFAULT_OUTPUT_DIR%]: 
if "%OUTPUT_DIR%"=="" set "OUTPUT_DIR=%DEFAULT_OUTPUT_DIR%"

if not exist "%OUTPUT_DIR%" mkdir "%OUTPUT_DIR%"

set "LOG_FILE=%OUTPUT_DIR%\build.log"
call :log ==================================================
call :log WordToPdfConverter build started
call :log Output directory: %OUTPUT_DIR%
call :log Log file: %LOG_FILE%
call :log ==================================================

echo [1/4] Upgrading pip...
python -m pip install --upgrade pip >> "%LOG_FILE%" 2>&1
if errorlevel 1 (
    call :log ERROR: pip upgrade failed.
    goto :build_failed
)

echo [2/4] Installing requirements...
python -m pip install -r "%SCRIPT_DIR%requirements.txt" >> "%LOG_FILE%" 2>&1
if errorlevel 1 (
    call :log ERROR: requirements install failed.
    goto :build_failed
)

echo [3/4] Installing pyinstaller...
python -m pip install pyinstaller >> "%LOG_FILE%" 2>&1
if errorlevel 1 (
    call :log ERROR: pyinstaller install failed.
    goto :build_failed
)

echo [4/4] Running pyinstaller...
call :log Running pyinstaller...
pyinstaller --noconfirm --clean --onefile --windowed ^
  --name WordToPdfConverter ^
  --distpath "%OUTPUT_DIR%" ^
  --hidden-import pythoncom ^
  --hidden-import pywintypes ^
  --collect-submodules win32com ^
  "%SCRIPT_DIR%app.py" >> "%LOG_FILE%" 2>&1

if errorlevel 1 (
    call :log ERROR: pyinstaller build failed.
    goto :build_failed
)

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

call :log ERROR: Build finished, but no executable was found.
call :log Check log file for details: %LOG_FILE%
goto :build_failed

:build_success
call :log SUCCESS: Build complete. Executable is at "%EXE_PATH%"
echo.
echo Build complete. Executable is at "%EXE_PATH%"
echo Full log: "%LOG_FILE%"
call :maybe_pause
exit /b 0

:build_failed
echo.
echo Build failed. See log for details: "%LOG_FILE%"
if exist "%LOG_FILE%" (
    echo.
    echo --- Last 40 log lines ---
    powershell -NoProfile -Command "if (Test-Path '%LOG_FILE%') { Get-Content '%LOG_FILE%' -Tail 40 }" 2>nul
)
call :maybe_pause
exit /b 1

:maybe_pause
if "%KEEP_WINDOW_OPEN%"=="1" (
    echo.
    pause
)
goto :eof

:log
echo [%date% %time%] %~1>> "%LOG_FILE%"
goto :eof
