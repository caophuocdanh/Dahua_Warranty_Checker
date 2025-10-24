@echo off
echo Building DahuaWarrantyChecker.exe...

REM Install/Update dependencies from requirements.txt
echo Installing/Updating Python dependencies...
pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo Failed to install dependencies. Please check your Python and pip installation.
    pause
    exit /b 1
)

REM Clean up previous build artifacts (before build)
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist DahuaWarrantyChecker.spec del DahuaWarrantyChecker.spec

REM Build the executable
pyinstaller --onefile --windowed --name "DahuaWarrantyChecker" --icon "icon.ico" --add-data "icon.ico;." check_warranty_gui.py

if %errorlevel% equ 0 (
    echo.
    echo Build successful!
    echo The executable is located in the 'dist' folder:
    echo   dist\DahuaWarrantyChecker.exe
    echo.
    echo Cleaning up temporary build files...
    if exist build rmdir /s /q build
    del *.spec
    echo. 
    echo You can now run the application from there.
) else (
    echo.
    echo Build failed! Please check the error messages above.
)

pause