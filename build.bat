@echo off
echo ============================================================
echo   SIDH Automator - Standalone Executable Builder (Windows)
echo ============================================================
echo.

:: Check if Python is installed
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python was not found on your system!
    echo.
    echo Please download and install Python from https://www.python.org/
    echo *IMPORTANT*: Make sure to check the box "Add Python to PATH"
    echo during the installation wizard.
    echo.
    pause
    exit /b
)

echo [1/3] Installing python dependencies...
python -m pip install --upgrade pip
python -m pip install playwright openpyxl pyinstaller
if %errorlevel% neq 0 (
    echo [ERROR] Failed to install python dependencies.
    echo.
    pause
    exit /b
)

echo.
echo [2/3] Compiling package into a single executable...
echo (This may take 1-2 minutes, please wait...)
python -m PyInstaller --onefile --noconsole --collect-all playwright --name "SIDH_Automator" gui.py
if %errorlevel% neq 0 (
    echo [ERROR] PyInstaller compilation failed!
    echo.
    pause
    exit /b
)

echo.
echo ============================================================
echo   SUCCESS! Standalone Executable Built Successfully
echo ============================================================
echo.
echo Opening output directory...
explorer dist
echo.
echo You can find your finished software at:
echo   dist\SIDH_Automator.exe
echo.
echo Copy that single 'SIDH_Automator.exe' file to sell or run on
echo any computer (no python installation needed by the buyer!).
echo ============================================================
echo.
pause

