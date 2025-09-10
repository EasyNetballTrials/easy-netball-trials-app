@echo off
setlocal

rem --- Always start from the folder this .bat lives in ---
cd /d "%~dp0"

rem --- Prefer local venv if it exists ---
if exist "venv\Scripts\python.exe" (
  set "PY_CMD=%CD%\venv\Scripts\python.exe"
) else (
  rem --- Try the py launcher (usually installed on Windows) ---
  where py >nul 2>&1 && set "PY_CMD=py"
  if not defined PY_CMD (
    rem --- Try python from PATH ---
    where python >nul 2>&1 && set "PY_CMD=python"
  )
  if not defined PY_CMD (
    rem --- Fall back to common install locations ---
    for %%P in (
      "%LocalAppData%\Programs\Python\Python312\python.exe"
      "%ProgramFiles%\Python312\python.exe"
      "%ProgramFiles(x86)%\Python312\python.exe"
    ) do (
      if exist %%~fP set "PY_CMD=%%~fP"
    )
  )
)

if not defined PY_CMD (
  echo Could not find Python. Please install Python 3.12+ and tick "Add to PATH".
  echo Or create a venv with:  py -3.12 -m venv venv ^& venv\Scripts\pip install -r requirements.txt
  pause
  exit /b 1
)

rem --- Small diagnostics so we can see what's going on ---
echo Using Python: %PY_CMD%
if not exist "app.py" (
  echo ERROR: app.py not found in %CD%
  pause
  exit /b 1
)

rem --- Open browser and start Flask ---
start "" http://127.0.0.1:5000/
"%PY_CMD%" app.py
echo Flask exited with code %errorlevel%
pause
