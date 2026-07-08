@echo off
setlocal
chcp 65001 >nul
cd /d "%~dp0"

set "APP=实训报告生成器网页.py"
set "LOCAL_PY=%~dp0.venv\Scripts\python.exe"

echo ========================================
echo 实训报告生成器
echo ========================================
echo.

if not exist "%APP%" (
    echo 未找到 %APP%
    echo 请确认本 bat 文件和 Python 脚本在同一个文件夹中。
    echo.
    pause
    exit /b 1
)

if exist "%LOCAL_PY%" (
    echo 使用项目本地 Python 环境。
    "%LOCAL_PY%" -m pip install -r requirements.txt
    "%LOCAL_PY%" "%APP%"
    goto END
)

where py >nul 2>nul
if %errorlevel%==0 (
    echo 使用 Windows Python Launcher。
    py -3 -m pip install -r requirements.txt
    py -3 "%APP%"
    goto END
)

where python >nul 2>nul
if %errorlevel%==0 (
    echo 使用系统 Python。
    python -m pip install -r requirements.txt
    python "%APP%"
    goto END
)

echo 没有找到 Python。
echo.
echo 请先安装 Python 3.10 或以上版本：
echo https://www.python.org/downloads/
echo.
echo 安装时请勾选：Add python.exe to PATH
echo.

:END
echo.
pause
