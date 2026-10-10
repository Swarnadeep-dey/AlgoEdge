#!/usr/bin/env bash
# Mac / Linux equivalent of run_pipeline.bat  (first time: chmod +x run_pipeline.sh)
set -e
cd "$(dirname "$0")/.."
PY=$(command -v python3 || command -v python)
echo "[1/2] Running Daily Screener..."
"$PY" python_codes/Screeners.py
echo "[2/2] Running Backtester on latest database..."
"$PY" python_codes/Backtester.py
echo "Pipeline complete! Reports saved in ./results"
echo "Opening the dashboard in your browser... (press Ctrl+C to stop it)"
"$PY" -m streamlit run python_codes/Dashboard.py
