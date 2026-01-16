@echo off
chcp 65001 >nul
echo ========================================
echo    图片处理工具 - 打包脚本
echo ========================================
echo.

REM 检查 PyInstaller 是否安装
py -c "import PyInstaller" 2>nul
if errorlevel 1 (
    echo 正在安装 PyInstaller...
    pip install pyinstaller
)

echo 开始打包...
echo.

REM 清理旧的构建文件
if exist "dist" rmdir /s /q dist
if exist "build" rmdir /s /q build
if exist "图片处理工具.spec" del "图片处理工具.spec"

REM 打包命令
pyinstaller --noconfirm ^
    --onefile ^
    --windowed ^
    --name "图片处理工具" ^
    --icon "NONE" ^
    --add-data "config.ini;." ^
    --add-data "workflow_i2i.json;." ^
    --add-data "fonts;fonts" ^
    --add-data "credentials.json;." ^
    --hidden-import "pandas" ^
    --hidden-import "openpyxl" ^
    --hidden-import "PIL" ^
    --hidden-import "cv2" ^
    --hidden-import "numpy" ^
    --hidden-import "google.oauth2" ^
    --hidden-import "googleapiclient" ^
    --hidden-import "requests" ^
    gui_app.py

if errorlevel 1 (
    echo.
    echo ========================================
    echo    打包失败！请检查错误信息
    echo ========================================
    pause
    exit /b 1
)

echo.
echo ========================================
echo    打包成功！
echo    输出文件: dist\图片处理工具.exe
echo ========================================
echo.

REM 复制必要文件到dist目录
echo 正在复制配置文件...
copy "config.ini" "dist\" >nul
copy "workflow_i2i.json" "dist\" >nul
copy "credentials.json" "dist\" >nul 2>nul
xcopy "fonts" "dist\fonts\" /E /I /Y >nul

echo.
echo 所有文件已复制到 dist 目录
echo 请将 dist 目录整体发送给客户
echo.
pause
