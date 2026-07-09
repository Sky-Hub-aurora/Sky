@echo off
setlocal EnableExtensions
cd /d "%~dp0"

set "APP=start_report_web.py"
set "REQ=requirements.txt"

if not exist "%APP%" (
    echo Missing %APP%.
    echo Please keep this BAT file in the report generator folder.
    pause
    exit /b 1
)

set "RUN_KIND="
set "RUN_EXE="
set "RUN_ARGS="
set "INSTALL_KIND="
set "INSTALL_EXE="
set "INSTALL_ARGS="

call :CHECK_EXE "%~dp0.venv\Scripts\python.exe"
call :CHECK_EXE "D:\ruanjian\python.exe"
call :CHECK_PY "-3"
call :CHECK_CMD "python"
call :CHECK_EXE "%LocalAppData%\Programs\Python\Python312\python.exe"
call :CHECK_EXE "%LocalAppData%\Programs\Python\Python311\python.exe"
call :CHECK_EXE "%LocalAppData%\Programs\Python\Python310\python.exe"

if defined RUN_KIND goto RUN_APP

echo No Python with required modules was found.
if not defined INSTALL_KIND (
    echo Python was not found. Please install Python 3.10 or newer, then run this file again.
    pause
    exit /b 1
)

if not exist "%REQ%" (
    echo Missing %REQ%. Cannot install required modules.
    pause
    exit /b 1
)

echo Installing missing modules for the first available Python...
call :RUN_PIP install -r "%REQ%"
if errorlevel 1 (
    echo Dependency installation failed.
    echo If your network blocks pip, install modules manually:
    echo python -m pip install -r requirements.txt
    pause
    exit /b 1
)

set "RUN_KIND=%INSTALL_KIND%"
set "RUN_EXE=%INSTALL_EXE%"
set "RUN_ARGS=%INSTALL_ARGS%"

:RUN_APP
echo Starting local report generator...
echo Browser address: http://127.0.0.1:8765/
echo.
call :RUN_PYTHON "%APP%" %*
set "EXIT_CODE=%ERRORLEVEL%"
if not "%EXIT_CODE%"=="0" pause
exit /b %EXIT_CODE%

:CHECK_EXE
if defined RUN_KIND exit /b
set "CAND=%~1"
if not exist "%CAND%" exit /b
if not defined INSTALL_KIND (
    set "INSTALL_KIND=EXE"
    set "INSTALL_EXE=%CAND%"
    set "INSTALL_ARGS="
)
"%CAND%" -c "import docx, pptx, PIL" >nul 2>nul
if not errorlevel 1 (
    set "RUN_KIND=EXE"
    set "RUN_EXE=%CAND%"
    set "RUN_ARGS="
)
exit /b

:CHECK_PY
if defined RUN_KIND exit /b
set "CAND_ARGS=%~1"
py %CAND_ARGS% -c "import sys" >nul 2>nul
if errorlevel 1 exit /b
if not defined INSTALL_KIND (
    set "INSTALL_KIND=PY"
    set "INSTALL_EXE=py"
    set "INSTALL_ARGS=%CAND_ARGS%"
)
py %CAND_ARGS% -c "import docx, pptx, PIL" >nul 2>nul
if not errorlevel 1 (
    set "RUN_KIND=PY"
    set "RUN_EXE=py"
    set "RUN_ARGS=%CAND_ARGS%"
)
exit /b

:CHECK_CMD
if defined RUN_KIND exit /b
set "CAND_CMD=%~1"
where "%CAND_CMD%" >nul 2>nul
if errorlevel 1 exit /b
if not defined INSTALL_KIND (
    set "INSTALL_KIND=CMD"
    set "INSTALL_EXE=%CAND_CMD%"
    set "INSTALL_ARGS="
)
%CAND_CMD% -c "import docx, pptx, PIL" >nul 2>nul
if not errorlevel 1 (
    set "RUN_KIND=CMD"
    set "RUN_EXE=%CAND_CMD%"
    set "RUN_ARGS="
)
exit /b

:RUN_PIP
if /I "%INSTALL_KIND%"=="EXE" (
    "%INSTALL_EXE%" -m pip %*
) else if /I "%INSTALL_KIND%"=="PY" (
    py %INSTALL_ARGS% -m pip %*
) else (
    %INSTALL_EXE% -m pip %*
)
exit /b %ERRORLEVEL%

:RUN_PYTHON
if /I "%RUN_KIND%"=="EXE" (
    "%RUN_EXE%" %*
) else if /I "%RUN_KIND%"=="PY" (
    py %RUN_ARGS% %*
) else (
    %RUN_EXE% %*
)
exit /b %ERRORLEVEL%
