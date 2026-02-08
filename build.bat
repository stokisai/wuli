@echo off
chcp 65001 >nul
echo ========================================
echo    图片处理工具 - 打包脚本
echo ========================================
echo.

REM Check PyInstaller
py -c "import PyInstaller" 2>nul
if errorlevel 1 (
    echo Installing PyInstaller...
    pip install pyinstaller
)

echo Start building...
echo.

REM Clean old build artifacts
if exist "dist" rmdir /s /q dist
if exist "build" rmdir /s /q build
if exist "图片处理工具.spec" del "图片处理工具.spec"

REM Build command
pyinstaller --noconfirm ^
    --onefile ^
    --windowed ^
    --name "图片处理工具" ^
    --icon "NONE" ^
    --add-data "config.ini;." ^
    --add-data "workflow_i2i.json;." ^
    --add-data "fonts;fonts" ^
    --add-data "styles;styles" ^
    --add-data "credentials.json;." ^
    --hidden-import "pandas" ^
    --hidden-import "openpyxl" ^
    --hidden-import "PIL" ^
    --hidden-import "cv2" ^
    --hidden-import "numpy" ^
    --hidden-import "google.oauth2" ^
    --hidden-import "googleapiclient" ^
    --hidden-import "requests" ^
    --hidden-import "PySide6" ^
    --hidden-import "PySide6.QtCore" ^
    --hidden-import "PySide6.QtGui" ^
    --hidden-import "PySide6.QtWidgets" ^
    --hidden-import "shiboken6" ^
    gui_app.py

if errorlevel 1 (
    echo.
    echo ========================================
    echo    Build failed
    echo ========================================
    pause
    exit /b 1
)

echo.
echo ========================================
echo    Build success
echo    Output: dist\图片处理工具.exe
echo ========================================
echo.

REM Copy runtime files
echo Copying runtime files...
copy "config.ini" "dist\" >nul
copy "workflow_i2i.json" "dist\" >nul
copy "credentials.json" "dist\" >nul 2>nul
xcopy "fonts" "dist\fonts\" /E /I /Y >nul
xcopy "styles" "dist\styles\" /E /I /Y >nul

echo.
echo Done.
pause
