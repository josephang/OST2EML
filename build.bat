@echo off
echo =========================================
echo Building Premium OST to EML Extractor App
echo =========================================

REM Check if Python is installed
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo Python is not installed or not in your PATH. Please install Python 3.
    pause
    exit /b
)

echo.
echo Installing requirements...
python -m pip install --upgrade pip
python -m pip install -r requirements.txt

echo.
echo Packaging the application into a standalone Executable...
REM PyInstaller will bundle the Python script and customtkinter into one .exe file
REM We use python -m pyinstaller to ensure it finds the locally installed module
python -m PyInstaller --noconfirm --onedir --windowed --add-data "%LOCALAPPDATA%\Programs\Python\Python310\Lib\site-packages\customtkinter;customtkinter/" app.py

echo.
if exist "dist\app\OST_Extractor.exe" (
    echo Build complete. The OST_Extractor.exe is located in the "dist\app" folder.
) else if exist "dist\app\app.exe" (
    echo Build complete. The app.exe is located in the "dist\app" folder.
    rename "dist\app\app.exe" "OST_Extractor.exe"
) else (
    echo WARNING: Build might have failed. Please check the logs above.
)

pause
