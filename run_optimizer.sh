#!/usr/bin/env bash
# Mac / Linux equivalent of run_optimizer.bat  (first time: chmod +x run_optimizer.sh)
set -e
cd "$(dirname "$0")"
PY=$(command -v python3 || command -v python)
echo "Running Optimization Engine (this may take a few minutes)..."
"$PY" python_codes/Optimeser.py
echo "Done! Check results/Optimization.xlsx"
