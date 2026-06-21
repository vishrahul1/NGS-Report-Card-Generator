@echo off
echo Starting NEO GEETANJALI Report Card Generator UI...
cd /d "%~dp0"
python -m streamlit run app.py --server.port 8501 --browser.gatherUsageStats false
pause
