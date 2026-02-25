@echo off
cd /d "%~dp0"
echo Avvio Affitti Brevi...
echo L'app si aprira' nel browser a: http://localhost:8501
echo Per chiudere: premi Ctrl+C in questa finestra
echo.
streamlit run app.py
pause
