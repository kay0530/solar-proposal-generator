@echo off
cd /d "%~dp0"
python -m streamlit run proposal_generator/app.py --server.port 8502
pause
