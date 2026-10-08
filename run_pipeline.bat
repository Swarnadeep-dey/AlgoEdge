@echo off
echo ===================================================
echo Starting MarketMentor Quantitative Pipeline
echo ===================================================

:: Always run from the folder this batch file lives in
cd /d "%~dp0"

where python >nul 2>nul
if %errorlevel% neq 0 (
    echo [ERROR] Python not found. Install Python 3.9+ and tick "Add to PATH".
    pause
    exit /b 1
)

echo.
echo [1/2] Running Daily Screener...
python python_codes\Screeners.py
if %errorlevel% neq 0 (
    echo [ERROR] Screener failed to execute.
    pause
    exit /b %errorlevel%
)

echo.
echo [2/2] Running Backtester on latest database...
python python_codes\Backtester.py
if %errorlevel% neq 0 (
    echo [ERROR] Backtester failed to execute.
    pause
    exit /b %errorlevel%
)

echo.
echo ===================================================
echo Pipeline Complete! Reports saved in the results folder.
echo ===================================================
pause
