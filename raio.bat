@echo off
cd /d "C:\Users\anderson.marley\Documents\Projeta\7 - Desenvolvimento\DWG-EXCEL\scripts"
call venv\Scripts\activate
start "" /min cmd /c "python -m streamlit run app.py"