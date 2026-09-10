@echo off
cd /d "%~dp0"
for /f "usebackq tokens=1,* delims==" %%A in (".env") do (
    if not "%%A"=="" set "%%A=%%B"
)
"%~dp0.venv\Scripts\uvicorn.exe" app.main:app --reload --port 3400
