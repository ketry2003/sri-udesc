@echo off
cd /d "%~dp0.."

if not exist "venv\Scripts\python.exe" (
    echo Criando ambiente virtual...
    python -m venv venv
)

call "venv\Scripts\activate.bat"

python -m pip install -r requirements.txt

echo Iniciando sistema...
start "" "http://localhost:8501"

python -m streamlit run app.py