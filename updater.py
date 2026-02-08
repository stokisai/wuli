# -*- coding: utf-8 -*-
"""
OTA 自动更新模块
- 通过 GitHub Releases API 检查新版本
- 下载 ZIP 更新包并解压覆盖
- 重启应用完成更新
"""

import os
import sys
import re
import shutil
import zipfile
import tempfile
import subprocess
import logging
from dataclasses import dataclass

import requests
from PySide6.QtCore import QThread, Signal, Qt, QTimer
from PySide6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel,
    QPushButton, QProgressBar, QTextEdit, QMessageBox,
)

logger = logging.getLogger(__name__)


# ============================================================
# 版本比较
# ============================================================

def parse_version(version_str: str) -> tuple:
    """解析版本号字符串为 (major, minor, patch) 元组"""
    cleaned = version_str.strip().lstrip("v")
    match = re.match(r"^(\d+)\.(\d+)\.(\d+)", cleaned)
    if not match:
        raise ValueError(f"无效的版本号: {version_str}")
    return (int(match.group(1)), int(match.group(2)), int(match.group(3)))


def is_newer(remote_version: str, local_version: str) -> bool:
    """判断远程版本是否比本地版本更新"""
    try:
        return parse_version(remote_version) > parse_version(local_version)
    except ValueError:
        return False


# ============================================================
# 数据结构
# ============================================================

@dataclass
class ReleaseInfo:
    version: str
    download_url: str       # ZIP 资源的直接下载链接
    changelog: str
    html_url: str           # Release 页面链接（备用）


# ============================================================
# 运行环境检测
# ============================================================

def is_frozen() -> bool:
    """是否以 PyInstaller 打包的 exe 运行"""
    return getattr(sys, 'frozen', False)


def get_app_dir() -> str:
    """获取应用程序所在目录"""
    if is_frozen():
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


# ============================================================
# 检查更新线程
# ============================================================

class UpdateCheckWorker(QThread):
    """后台检查 GitHub Releases 是否有新版本"""
    update_available = Signal(object)   # ReleaseInfo
    no_update = Signal()
    check_failed = Signal(str)

    def __init__(self, current_version: str, repo: str, parent=None):
        super().__init__(parent)
        self.current_version = current_version
        self.repo = repo
        self.api_url = f"https://api.github.com/repos/{repo}/releases/latest"

    def run(self):
        try:
            resp = requests.get(self.api_url, timeout=10, headers={
                "Accept": "application/vnd.github.v3+json"
            })
            if resp.status_code == 403:
                self.check_failed.emit("GitHub API 请求频率超限，请稍后再试")
                return
            if resp.status_code == 404:
                self.check_failed.emit("未找到发布版本")
                return
            if resp.status_code != 200:
                self.check_failed.emit(f"API 返回状态码 {resp.status_code}")
                return

            data = resp.json()
            tag = data.get("tag_name", "")
            remote_ver = tag.lstrip("v")

            if not is_newer(remote_ver, self.current_version):
                self.no_update.emit()
                return

            # 查找 .zip 资源
            download_url = ""
            for asset in data.get("assets", []):
                if asset["name"].lower().endswith(".zip"):
                    download_url = asset["browser_download_url"]
                    break

            if not download_url:
                self.check_failed.emit("Release 中未找到 ZIP 更新包")
                return

            info = ReleaseInfo(
                version=remote_ver,
                download_url=download_url,
                changelog=data.get("body", "") or "暂无更新说明",
                html_url=data.get("html_url", ""),
            )
            self.update_available.emit(info)

        except requests.exceptions.ConnectionError:
            self.check_failed.emit("无法连接网络")
        except requests.exceptions.Timeout:
            self.check_failed.emit("连接超时")
        except Exception as e:
            self.check_failed.emit(f"检查更新失败: {e}")


# ============================================================
# 下载更新线程
# ============================================================

class DownloadWorker(QThread):
    """后台下载 ZIP 更新包，带进度回报"""
    progress = Signal(int, int)         # downloaded_bytes, total_bytes
    download_complete = Signal(str)     # 下载文件路径
    download_failed = Signal(str)

    def __init__(self, url: str, save_path: str, parent=None):
        super().__init__(parent)
        self.url = url
        self.save_path = save_path
        self._cancelled = False

    def cancel(self):
        self._cancelled = True

    def run(self):
        try:
            resp = requests.get(self.url, stream=True, timeout=30)
            resp.raise_for_status()
            total = int(resp.headers.get("content-length", 0))
            downloaded = 0

            with open(self.save_path, "wb") as f:
                for chunk in resp.iter_content(chunk_size=8192):
                    if self._cancelled:
                        f.close()
                        self._cleanup()
                        self.download_failed.emit("下载已取消")
                        return
                    f.write(chunk)
                    downloaded += len(chunk)
                    self.progress.emit(downloaded, total)

            self.download_complete.emit(self.save_path)
        except Exception as e:
            self._cleanup()
            self.download_failed.emit(f"下载失败: {e}")

    def _cleanup(self):
        try:
            if os.path.exists(self.save_path):
                os.remove(self.save_path)
        except OSError:
            pass


# ============================================================
# 解压与覆盖
# ============================================================

def extract_update(zip_path: str) -> str:
    """解压 ZIP 到临时目录，返回解压后的根目录路径"""
    extract_dir = os.path.join(tempfile.gettempdir(), "_update_extract")
    if os.path.exists(extract_dir):
        shutil.rmtree(extract_dir)

    with zipfile.ZipFile(zip_path, 'r') as zf:
        zf.extractall(extract_dir)

    # 如果 ZIP 内只有一个顶层文件夹，返回该文件夹
    items = os.listdir(extract_dir)
    if len(items) == 1:
        single = os.path.join(extract_dir, items[0])
        if os.path.isdir(single):
            return single
    return extract_dir


def apply_files(source_dir: str, target_dir: str):
    """将 source_dir 中的所有文件覆盖到 target_dir"""
    for item in os.listdir(source_dir):
        src = os.path.join(source_dir, item)
        dst = os.path.join(target_dir, item)
        if os.path.isdir(src):
            if os.path.exists(dst):
                shutil.rmtree(dst)
            shutil.copytree(src, dst)
        else:
            shutil.copy2(src, dst)


# ============================================================
# 更新应用并重启
# ============================================================

def apply_update_and_restart(zip_path: str):
    """
    启动更新脚本并立即退出当前程序。
    所有文件操作（解压、覆盖、重启）都由 bat 脚本在终端中完成。
    """
    app_dir = get_app_dir()
    pid = os.getpid()

    if is_frozen():
        exe_name = os.path.basename(sys.executable)
    else:
        exe_name = ""

    bat_path = os.path.join(tempfile.gettempdir(), "_ota_update.cmd")

    # bat 脚本：等待退出 → 解压 ZIP → 覆盖文件 → 重启
    bat = f'''@echo off
chcp 65001 >nul
title 图片处理工具 - 自动更新
color 0A
echo.
echo  ========================================
echo    图片处理工具 - 自动更新
echo  ========================================
echo.
echo  [1/4] 等待旧程序退出...

set /a c=0
:WAIT
tasklist /FI "PID eq {pid}" 2>nul | find /I "{pid}" >nul
if errorlevel 1 goto :EXTRACT
set /a c+=1
if %c% geq 30 (
    echo.
    echo  超时！请手动关闭程序后重试。
    pause
    goto :END
)
timeout /t 1 /nobreak >nul
goto :WAIT

:EXTRACT
timeout /t 1 /nobreak >nul
echo  [OK] 旧程序已退出
echo.
echo  [2/4] 正在解压更新包...

set "TEMP_DIR=%TEMP%\\_ota_extract"
if exist "%TEMP_DIR%" rmdir /S /Q "%TEMP_DIR%"
powershell -NoProfile -Command "Expand-Archive -Path '{zip_path}' -DestinationPath '%TEMP_DIR%' -Force" 2>nul
if errorlevel 1 (
    echo.
    echo  解压失败！更新包: {zip_path}
    pause
    goto :END
)
echo  [OK] 解压完成

REM 检查是否有单层文件夹包裹
set "SRC_DIR=%TEMP_DIR%"
for /f "tokens=*" %%i in ('dir /b /ad "%TEMP_DIR%" 2^>nul') do (
    set "SINGLE=%%i"
)
REM 如果只有一个子文件夹，进入它
for /f %%n in ('dir /b /ad "%TEMP_DIR%" 2^>nul ^| find /c /v ""') do (
    if %%n==1 (
        for /f "tokens=*" %%i in ('dir /b /ad "%TEMP_DIR%"') do set "SRC_DIR=%TEMP_DIR%\\%%i"
    )
)

echo.
echo  [3/4] 正在覆盖程序文件...
echo         目标: {app_dir}

xcopy "%SRC_DIR%\\*" "{app_dir}\\" /E /Y /I /Q >nul 2>nul
if errorlevel 1 (
    echo.
    echo  文件覆盖失败！
    echo  更新文件在: %SRC_DIR%
    echo  请手动复制到: {app_dir}
    pause
    goto :END
)
echo  [OK] 文件覆盖完成

echo.
echo  [4/4] 正在启动新版本...
echo.
echo  ========================================
echo    更新完成！程序即将重启...
echo  ========================================
echo.
timeout /t 2 /nobreak >nul

start "" "{os.path.join(app_dir, exe_name) if exe_name else 'echo done'}"

:END
REM 清理临时文件
if exist "%TEMP_DIR%" rmdir /S /Q "%TEMP_DIR%" >nul 2>nul
if exist "{zip_path}" del /F /Q "{zip_path}" >nul 2>nul
timeout /t 1 /nobreak >nul
del /F /Q "%~f0" >nul 2>nul
'''

    with open(bat_path, "w", encoding="utf-8") as f:
        f.write(bat)

    # 以可见终端窗口运行（CREATE_NEW_CONSOLE 让用户看到进度）
    CREATE_NEW_CONSOLE = 0x00000010
    subprocess.Popen(
        ["cmd.exe", "/C", bat_path],
        creationflags=CREATE_NEW_CONSOLE,
    )


# ============================================================
# 更新对话框
# ============================================================

class UpdateDialog(QDialog):
    """更新提示对话框：版本信息 + 更新日志 + 下载进度"""

    def __init__(self, release_info: ReleaseInfo, current_version: str, parent=None):
        super().__init__(parent)
        self.release_info = release_info
        self.current_version = current_version
        self.download_worker = None
        self._zip_path = None

        self.setWindowTitle("软件更新")
        self.setMinimumSize(500, 400)
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)
        self._build_ui()
        self._apply_style()

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(12)
        layout.setContentsMargins(24, 20, 24, 20)

        header = QLabel(f"发现新版本 v{self.release_info.version}")
        header.setObjectName("updateHeader")
        layout.addWidget(header)

        ver = QLabel(
            f"当前版本: v{self.current_version}  →  "
            f"最新版本: v{self.release_info.version}"
        )
        ver.setObjectName("versionLabel")
        layout.addWidget(ver)

        layout.addWidget(QLabel("更新内容:"))

        self.changelog_text = QTextEdit()
        self.changelog_text.setReadOnly(True)
        self.changelog_text.setPlainText(self.release_info.changelog)
        self.changelog_text.setMaximumHeight(160)
        layout.addWidget(self.changelog_text)

        self.progress_bar = QProgressBar()
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setFormat("下载中... %p%")
        self.progress_bar.setVisible(False)
        layout.addWidget(self.progress_bar)

        self.status_label = QLabel("")
        self.status_label.setObjectName("updateStatus")
        layout.addWidget(self.status_label)

        btn_layout = QHBoxLayout()
        self.skip_btn = QPushButton("跳过")
        self.skip_btn.clicked.connect(self.reject)
        btn_layout.addWidget(self.skip_btn)
        btn_layout.addStretch()
        self.update_btn = QPushButton("立即更新")
        self.update_btn.setObjectName("updateNowBtn")
        self.update_btn.clicked.connect(self._start_download)
        btn_layout.addWidget(self.update_btn)
        layout.addLayout(btn_layout)

    def _apply_style(self):
        self.setStyleSheet("""
            QDialog { background: #1e293b; }
            QLabel { color: #e7eef8; font-size: 13px; }
            QLabel#updateHeader {
                font-size: 18px; font-weight: bold; color: #8bd3ff;
            }
            QLabel#versionLabel { color: #c7d3e2; }
            QLabel#updateStatus { font-size: 12px; color: #d7e7ff; }
            QTextEdit {
                background: rgba(15,23,42,0.82); color: #f0f6ff;
                border: 1px solid rgba(148,163,184,0.35);
                border-radius: 6px; padding: 8px; font-size: 12px;
            }
            QProgressBar {
                border: none; border-radius: 5px;
                background: rgba(15,23,42,0.78);
                text-align: center; color: white;
                font-weight: bold; font-size: 11px; min-height: 22px;
            }
            QProgressBar::chunk {
                background: qlineargradient(x1:0,y1:0,x2:1,y2:0,
                    stop:0 #6366f1, stop:1 #d946ef);
                border-radius: 5px;
            }
            QPushButton {
                font-size: 13px; font-weight: bold; color: #f8fbff;
                border: 1px solid #5a677a; border-radius: 8px; padding: 10px 24px;
                background: #3f4b5d;
            }
            QPushButton:hover { background: #52627a; border-color: #71809a; }
            QPushButton#updateNowBtn { background: #16a34a; }
            QPushButton#updateNowBtn:hover { background: #22c55e; }
            QPushButton#updateNowBtn:disabled {
                background: rgba(100,116,139,0.4);
                color: rgba(255,255,255,0.4);
            }
        """)

    # ---- 下载流程 ----

    def _start_download(self):
        self.update_btn.setEnabled(False)
        self.skip_btn.setText("取消")
        self.progress_bar.setVisible(True)
        self.progress_bar.setValue(0)
        self.status_label.setText("正在下载更新包...")

        filename = os.path.basename(self.release_info.download_url)
        if not filename.lower().endswith(".zip"):
            filename = f"update_v{self.release_info.version}.zip"
        self._zip_path = os.path.join(tempfile.gettempdir(), filename)

        self.download_worker = DownloadWorker(
            self.release_info.download_url, self._zip_path
        )
        self.download_worker.progress.connect(self._on_progress)
        self.download_worker.download_complete.connect(self._on_complete)
        self.download_worker.download_failed.connect(self._on_failed)
        self.download_worker.start()

        self.skip_btn.clicked.disconnect()
        self.skip_btn.clicked.connect(self._cancel)

    def _cancel(self):
        if self.download_worker:
            self.download_worker.cancel()
        self.reject()

    def _on_progress(self, downloaded, total):
        if total > 0:
            self.progress_bar.setMaximum(total)
            self.progress_bar.setValue(downloaded)
            mb_d = downloaded / (1024 * 1024)
            mb_t = total / (1024 * 1024)
            self.status_label.setText(f"下载中: {mb_d:.1f} MB / {mb_t:.1f} MB")
        else:
            self.progress_bar.setMaximum(0)
            mb_d = downloaded / (1024 * 1024)
            self.status_label.setText(f"下载中: {mb_d:.1f} MB")

    def _on_complete(self, zip_path):
        self.progress_bar.setValue(self.progress_bar.maximum())
        self.status_label.setText("下载完成，正在启动更新...")
        try:
            apply_update_and_restart(zip_path)
            # 立即强制退出程序，让 bat 脚本接管
            from PySide6.QtWidgets import QApplication
            QApplication.quit()
        except Exception as e:
            QMessageBox.warning(
                self, "更新失败",
                f"无法自动应用更新:\n{e}\n\n"
                f"更新包已下载到:\n{zip_path}\n请手动解压覆盖。"
            )
            self.accept()

    def _on_failed(self, error):
        self.status_label.setText(f"下载失败: {error}")
        self.update_btn.setEnabled(True)
        self.update_btn.setText("重试")
        self.skip_btn.setText("跳过")
        self.skip_btn.clicked.disconnect()
        self.skip_btn.clicked.connect(self.reject)
        self.progress_bar.setVisible(False)
