@echo off
echo ===================================================
echo Starting AlgoEdge Grid Search Optimizer
echo ===================================================

cd /d "%~dp0.."

where python >nul 2>nul
if %errorlevel% neq 0 (
    echo [ERROR] Python not found. Install Python 3.9+ and tick "Add to PATH".
    pause
    exit /b 1
)

echo.
echo Running Optimization Engine (This may take a few minutes)...
python python_codes\Optimeser.py
if %errorlevel% neq 0 (
    echo [ERROR] Optimizer failed to execute.
    pause
    exit /b %errorlevel%
)

echo.
echo ===================================================
echo Optimization Complete! Check results\Optimization.xlsx
echo ===================================================
pause
