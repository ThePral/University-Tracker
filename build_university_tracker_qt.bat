@echo off
setlocal
cd /d "%~dp0"
echo ==========================================
echo   University Tracker - PySide6 Builder
echo ==========================================

where python >nul 2>nul
if not errorlevel 1 (
  set "PYTHON_CMD=python"
) else (
  where py >nul 2>nul
  if not errorlevel 1 (
    set "PYTHON_CMD=py -3"
  ) else (
    echo Python was not found. Install Python 3.11 or newer and enable Add Python to PATH.
    goto error
  )
)

%PYTHON_CMD% --version
if errorlevel 1 goto error
%PYTHON_CMD% -m pip install -r requirements_qt.txt
if errorlevel 1 goto error
%PYTHON_CMD% -m PyInstaller --noconfirm --clean --windowed --onefile --name UniversityTrackerQt --distpath dist_qt --workpath build_qt --specpath build_qt --collect-all PySide6 university_tracker_qt.py
if errorlevel 1 goto error

rem Seed a fresh build folder with the existing workbook, without replacing user data.
if exist "Excel_Sheets\universities.xlsx" (
  if not exist "dist_qt\Excel_Sheets\universities.xlsx" (
    if not exist "dist_qt\Excel_Sheets" mkdir "dist_qt\Excel_Sheets"
    copy "Excel_Sheets\universities.xlsx" "dist_qt\Excel_Sheets\universities.xlsx" >nul
  )
)

echo.
echo Build complete!
echo EXE location: %~dp0dist_qt\UniversityTrackerQt.exe
echo Keep the Excel_Sheets folder beside the EXE to use your workbook.
pause
exit /b 0

:error
echo.
echo Build failed. Check the error above and make sure Python and pip are available.
pause
exit /b 1
