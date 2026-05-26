#!/bin/bash
# Quick script to run app_starter.py

echo "🚀 Starting Recon Engine..."
echo ""
echo "Features:"
echo "  ✅ All 18 Operations"
echo "  ✅ Workflow Builder"
echo "  ✅ Save/Load (JSON)"
echo "  ✅ NO win32com!"
echo ""

cd /Users/lshrinivaasan/Documents/Recon

# Check if streamlit is installed
if ! command -v streamlit &> /dev/null; then
    echo "⚠️  Streamlit not found. Installing..."
    pip install streamlit pandas openpyxl
fi

echo "Starting app..."
streamlit run app_starter.py


