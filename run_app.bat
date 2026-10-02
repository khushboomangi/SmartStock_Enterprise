@echo off
TITLE SmartStock AI Engine
echo Starting SmartStock AI...

cd /d "%~dp0"

:: Server run hoga aur browser me ek hi baar app auto-open ho jayegi
streamlit run app.py