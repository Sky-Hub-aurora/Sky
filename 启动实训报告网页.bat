@echo off
chcp 65001 >nul
cd /d "%~dp0"

set "BUNDLED_PY=C:\Users\SKY\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe"

if exist "%BUNDLED_PY%" (
    "%BUNDLED_PY%" "%~dp0实训报告生成器网页.py"
) else (
    python "%~dp0实训报告生成器网页.py"
)

pause
