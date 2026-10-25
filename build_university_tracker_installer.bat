@echo off
setlocal
cd /d "%~dp0"
echo ==========================================
echo   University Tracker - Installer Builder
echo ==========================================

where python >nul 2>nul
if not errorlevel 1 (
  set "PYTHON_CMD=python"
) else (
  where py >nul 2>nul
  if not errorlevel 1 (
    set "PYTHON_CMD=py -3"
  ) else (
    echo Python was not found. Install 64-bit Python 3.11 or newer.
    goto error
  )
)

%PYTHON_CMD% installer\build_installer.py %*
if errorlevel 1 goto error
if not defined UT_NO_PAUSE pause
exit /b 0

:error
echo.
echo Installer build failed. See the error above.
if not defined UT_NO_PAUSE pause
exit /b 1
