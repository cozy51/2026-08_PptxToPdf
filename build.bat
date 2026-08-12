@echo off
setlocal
cd /d "%~dp0"

echo [1/4] Checking Python...
where py >nul 2>&1
if %errorlevel%==0 (
    set "PYTHON=py -3"
) else (
    where python >nul 2>&1
    if errorlevel 1 goto :no_python
    set "PYTHON=python"
)

echo [2/4] Creating the build environment...
if not exist ".venv-build\Scripts\python.exe" (
    %PYTHON% -m venv .venv-build || goto :failed
)

echo [3/4] Installing build dependencies...
".venv-build\Scripts\python.exe" -m pip install --upgrade pip || goto :failed
".venv-build\Scripts\python.exe" -m pip install -r requirements.txt || goto :failed

echo [4/4] Building dist\OfficePdfConverter.exe...
if exist "dist" rmdir /s /q "dist"
if exist "build" rmdir /s /q "build"
".venv-build\Scripts\python.exe" -m PyInstaller --noconfirm --clean OfficePdfConverter.spec || goto :failed

echo.
echo Build completed successfully: dist\OfficePdfConverter.exe
pause
exit /b 0

:no_python
echo.
echo ERROR: Python 3 was not found. Install Python on this build PC and try again.
pause
exit /b 1

:failed
echo.
echo ERROR: The build failed. Review the messages above.
pause
exit /b 1
