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
    --onedir ^
    --windowed ^
    --name "图片处理工具" ^
    --icon "NONE" ^
    --add-data "config.ini;." ^
    --add-data "workflow_i2i.json;." ^
    --add-data "fonts;fonts" ^
    --add-data "styles;styles" ^
    --add-data "templates;templates" ^
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
echo    Output: dist\图片处理工具\图片处理工具.exe
echo ========================================
echo.

REM Copy runtime files INTO the app folder (not dist root!)
REM This ensures the zip has a single top-level folder for OTA updater compatibility.
set "APP_DIR=dist\图片处理工具"
echo Copying runtime files to %APP_DIR% ...
copy "config.ini" "%APP_DIR%\" >nul
copy "workflow_i2i.json" "%APP_DIR%\" >nul
copy "credentials.json" "%APP_DIR%\" >nul 2>nul
xcopy "fonts" "%APP_DIR%\fonts\" /E /I /Y >nul
xcopy "styles" "%APP_DIR%\styles\" /E /I /Y >nul
xcopy "templates" "%APP_DIR%\templates\" /E /I /Y >nul

echo.
echo Done. To create release zip, run:
echo   cd dist ^&^& powershell Compress-Archive -Path '图片处理工具' -DestinationPath '..\wuli_vX.X.X.zip'
echo.
pause

