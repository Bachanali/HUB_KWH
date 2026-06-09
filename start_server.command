#!/bin/bash
# Double-click on Mac (you may need: right-click > Open the first time).
cd "$(dirname "$0")"
echo "=== VOLTGRID Energy Dashboard ==="
echo "Installing required libraries (first run only)..."
python3 -m pip install -r requirements.txt
echo ""
echo "Starting dashboard -> open http://localhost:5010  (login: admin / admin123)"
echo "Keep this window open. Press Ctrl+C to stop."
echo ""
python3 app.py
echo ""
echo "Server stopped."
read -p "Press Enter to close..."
