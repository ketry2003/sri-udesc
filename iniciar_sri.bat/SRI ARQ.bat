@echo off
cd d %~dp0

if not exist venv (
    echo Criando ambiente virtual...
    python -m venv venv
)

call venvScriptsactivate

pip install -r requirements.txt

echo Iniciando sistema...
start httplocalhost8501

python -m streamlit run app.py