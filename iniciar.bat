@echo off
cd /d "%~dp0"
python -m pip install -r requirements.txt --quiet --no-warn-script-location

rem Evita a pergunta de e-mail do Streamlit na primeira execucao
if not exist "%USERPROFILE%\.streamlit\credentials.toml" (
  mkdir "%USERPROFILE%\.streamlit" 2>nul
  > "%USERPROFILE%\.streamlit\credentials.toml" echo [general]
  >> "%USERPROFILE%\.streamlit\credentials.toml" echo email = ""
)

python -m streamlit run app.py --browser.gatherUsageStats false
pause
