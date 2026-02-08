# -*- coding: utf-8 -*-
"""OTA update module.

Flow:
1) Check latest GitHub Release
2) Download .zip package
3) Run update script (wait old process -> extract -> copy -> restart)
"""

import os
import re
import sys
import shutil
import zipfile
import tempfile
import subprocess
import logging
from dataclasses import dataclass

import requests
from PySide6.QtCore import QThread, Signal, Qt
from PySide6.QtWidgets import (
    QDialog,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QProgressBar,
    QTextEdit,
    QMessageBox,
    QApplication,
)

logger = logging.getLogger(__name__)


def parse_version(version_str: str) -> tuple:
    """Parse semantic version into (major, minor, patch)."""
    cleaned = (version_str or "").strip().lstrip("v")
    match = re.match(r"^(\d+)\.(\d+)\.(\d+)", cleaned)
    if not match:
        raise ValueError(f"Invalid version: {version_str}")
    return (int(match.group(1)), int(match.group(2)), int(match.group(3)))


def is_newer(remote_version: str, local_version: str) -> bool:
    """Return True if remote_version > local_version."""
    try:
        return parse_version(remote_version) > parse_version(local_version)
    except ValueError:
        return False


@dataclass
class ReleaseInfo:
    version: str
    download_url: str
    changelog: str
    html_url: str


def is_frozen() -> bool:
    """Whether running from PyInstaller executable."""
    return getattr(sys, "frozen", False)


def get_app_dir() -> str:
    """Application directory (exe dir or source dir)."""
    if is_frozen():
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


class UpdateCheckWorker(QThread):
    """Background worker checking latest release."""

    update_available = Signal(object)  # ReleaseInfo
    no_update = Signal()
    check_failed = Signal(str)

    def __init__(self, current_version: str, repo: str, parent=None):
        super().__init__(parent)
        self.current_version = current_version
        self.repo = repo
        self.api_url = f"https://api.github.com/repos/{repo}/releases/latest"

    def run(self):
        try:
            headers = {
                "Accept": "application/vnd.github.v3+json",
                "User-Agent": "ImageProcessingTool-Updater",
            }
            token = os.getenv("GITHUB_TOKEN") or os.getenv("GH_TOKEN")
            if token:
                headers["Authorization"] = f"Bearer {token}"

            resp = requests.get(self.api_url, timeout=10, headers=headers)
            if resp.status_code == 403:
                self.check_failed.emit("GitHub API rate limit exceeded. Please retry later.")
                return
            if resp.status_code == 404:
                self.check_failed.emit(
                    "Latest release not found (private repo unauthenticated or no Release)."
                )
                return
            if resp.status_code != 200:
                self.check_failed.emit(f"GitHub API returned status {resp.status_code}")
                return

            data = resp.json()
            tag = data.get("tag_name", "")
            remote_ver = tag.lstrip("v")

            if not is_newer(remote_ver, self.current_version):
                self.no_update.emit()
                return

            assets = data.get("assets", [])
            download_url = ""
            for asset in assets:
                name = (asset.get("name") or "").lower()
                if name.endswith(".zip"):
                    download_url = asset.get("browser_download_url", "")
                    break

            if not download_url:
                asset_names = [a.get("name", "") for a in assets]
                if asset_names:
                    self.check_failed.emit(
                        "Release has no .zip asset. Assets: " + ", ".join(asset_names)
                    )
                else:
                    self.check_failed.emit(
                        "Release has no downloadable assets. Please upload a .zip package."
                    )
                return

            info = ReleaseInfo(
                version=remote_ver,
                download_url=download_url,
                changelog=data.get("body", "") or "No changelog provided.",
                html_url=data.get("html_url", ""),
            )
            self.update_available.emit(info)

        except requests.exceptions.ConnectionError:
            self.check_failed.emit("Network unavailable.")
        except requests.exceptions.Timeout:
            self.check_failed.emit("Connection timeout.")
        except Exception as e:
            self.check_failed.emit(f"Update check failed: {e}")


class DownloadWorker(QThread):
    """Background ZIP downloader with progress."""

    progress = Signal(int, int)  # downloaded_bytes, total_bytes
    download_complete = Signal(str)
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
                        self._cleanup()
                        self.download_failed.emit("Download cancelled")
                        return
                    if not chunk:
                        continue
                    f.write(chunk)
                    downloaded += len(chunk)
                    self.progress.emit(downloaded, total)

            self.download_complete.emit(self.save_path)
        except Exception as e:
            self._cleanup()
            self.download_failed.emit(f"Download failed: {e}")

    def _cleanup(self):
        try:
            if os.path.exists(self.save_path):
                os.remove(self.save_path)
        except OSError:
            pass


def extract_update(zip_path: str) -> str:
    """Extract zip to temp and return extraction root."""
    extract_dir = os.path.join(tempfile.gettempdir(), "_ota_extract")
    if os.path.exists(extract_dir):
        shutil.rmtree(extract_dir)

    with zipfile.ZipFile(zip_path, "r") as zf:
        zf.extractall(extract_dir)

    items = os.listdir(extract_dir)
    if len(items) == 1:
        single = os.path.join(extract_dir, items[0])
        if os.path.isdir(single):
            return single
    return extract_dir


def apply_files(source_dir: str, target_dir: str):
    """Copy all files from source_dir to target_dir."""
    for item in os.listdir(source_dir):
        src = os.path.join(source_dir, item)
        dst = os.path.join(target_dir, item)
        if os.path.isdir(src):
            if os.path.exists(dst):
                shutil.rmtree(dst)
            shutil.copytree(src, dst)
        else:
            shutil.copy2(src, dst)


def _build_update_cmd_script(zip_path: str, app_dir: str, pid: int) -> str:
    """Build cmd script text that performs update after current process exits."""
    if is_frozen():
        restart_cmd = f'start "" "{os.path.join(app_dir, os.path.basename(sys.executable))}"'
    else:
        restart_cmd = f'start "" "{sys.executable}" "{os.path.join(app_dir, "gui_app.py")}"'

    script = f'''@echo off
chcp 65001 >nul
setlocal

set "PID={pid}"
set "ZIP={zip_path}"
set "APP_DIR={app_dir}"
set "TMP_DIR=%TEMP%\_ota_extract"

echo [Update] Waiting process exit: %PID%
:WAIT_PROC
tasklist /FI "PID eq %PID%" 2>nul | find /I "%PID%" >nul
if errorlevel 1 goto EXTRACT
timeout /t 1 /nobreak >nul
goto WAIT_PROC

:EXTRACT
echo [Update] Extracting package...
if exist "%TMP_DIR%" rmdir /S /Q "%TMP_DIR%"
powershell -NoProfile -Command "Expand-Archive -Path '%ZIP%' -DestinationPath '%TMP_DIR%' -Force"
if errorlevel 1 (
  echo [Update] Extract failed: %ZIP%
  pause
  goto END
)

set "SRC_DIR=%TMP_DIR%"
for /f %%n in ('dir /b /ad "%TMP_DIR%" 2^>nul ^| find /c /v ""') do (
  if %%n==1 (
    for /f "tokens=*" %%i in ('dir /b /ad "%TMP_DIR%"') do set "SRC_DIR=%TMP_DIR%\%%i"
  )
)

echo [Update] Copying files...
xcopy "%SRC_DIR%\*" "%APP_DIR%\" /E /Y /I /Q
if errorlevel 1 (
  echo [Update] File copy failed.
  pause
  goto END
)

echo [Update] Restarting app...
{restart_cmd}

:END
if exist "%TMP_DIR%" rmdir /S /Q "%TMP_DIR%" >nul 2>nul
if exist "%ZIP%" del /F /Q "%ZIP%" >nul 2>nul
del /F /Q "%~f0" >nul 2>nul
'''
    return script


def apply_update_and_restart(zip_path: str):
    """Launch update script and let it handle extraction/copy/restart."""
    app_dir = get_app_dir()
    pid = os.getpid()
    bat_path = os.path.join(tempfile.gettempdir(), "_ota_update.cmd")

    script = _build_update_cmd_script(zip_path, app_dir, pid)
    with open(bat_path, "w", encoding="utf-8") as f:
        f.write(script)

    CREATE_NEW_CONSOLE = 0x00000010
    subprocess.Popen(["cmd.exe", "/C", bat_path], creationflags=CREATE_NEW_CONSOLE)


class UpdateDialog(QDialog):
    """Update dialog: version info + changelog + download progress."""

    def __init__(self, release_info: ReleaseInfo, current_version: str, parent=None):
        super().__init__(parent)
        self.release_info = release_info
        self.current_version = current_version
        self.download_worker = None
        self._zip_path = None

        self.setWindowTitle("发现新版本")
        self.setMinimumSize(520, 420)
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowContextHelpButtonHint)
        self._build_ui()
        self._apply_style()

    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(12)
        layout.setContentsMargins(24, 20, 24, 20)

        header = QLabel(f"检测到新版本 v{self.release_info.version}")
        header.setObjectName("updateHeader")
        layout.addWidget(header)

        ver = QLabel(
            f"当前版本: v{self.current_version}  ->  最新版本: v{self.release_info.version}"
        )
        ver.setObjectName("versionLabel")
        layout.addWidget(ver)

        layout.addWidget(QLabel("更新日志:"))

        self.changelog_text = QTextEdit()
        self.changelog_text.setReadOnly(True)
        self.changelog_text.setPlainText(self.release_info.changelog)
        self.changelog_text.setMaximumHeight(170)
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
        self.skip_btn = QPushButton("稍后")
        self.skip_btn.clicked.connect(self.reject)
        btn_layout.addWidget(self.skip_btn)
        btn_layout.addStretch()

        self.update_btn = QPushButton("立即更新")
        self.update_btn.setObjectName("updateNowBtn")
        self.update_btn.clicked.connect(self._start_download)
        btn_layout.addWidget(self.update_btn)
        layout.addLayout(btn_layout)

    def _apply_style(self):
        self.setStyleSheet(
            """
            QDialog { background: #1e293b; }
            QLabel { color: #e7eef8; font-size: 13px; }
            QLabel#updateHeader { font-size: 18px; font-weight: bold; color: #8bd3ff; }
            QLabel#versionLabel { color: #c7d3e2; }
            QLabel#updateStatus { font-size: 12px; color: #d7e7ff; }
            QTextEdit {
                background: rgba(15,23,42,0.82);
                color: #f0f6ff;
                border: 1px solid rgba(148,163,184,0.35);
                border-radius: 6px;
                padding: 8px;
                font-size: 12px;
            }
            QProgressBar {
                border: none;
                border-radius: 5px;
                background: rgba(15,23,42,0.78);
                text-align: center;
                color: white;
                font-weight: bold;
                font-size: 11px;
                min-height: 22px;
            }
            QProgressBar::chunk {
                background: #4f8cff;
                border-radius: 5px;
            }
            QPushButton {
                font-size: 13px;
                font-weight: bold;
                color: #f8fbff;
                border: 1px solid #5a677a;
                border-radius: 8px;
                padding: 10px 24px;
                background: #3f4b5d;
            }
            QPushButton:hover { background: #52627a; border-color: #71809a; }
            QPushButton#updateNowBtn { background: #16a34a; }
            QPushButton#updateNowBtn:hover { background: #22c55e; }
            QPushButton#updateNowBtn:disabled {
                background: rgba(100,116,139,0.4);
                color: rgba(255,255,255,0.4);
            }
            """
        )

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

        self.download_worker = DownloadWorker(self.release_info.download_url, self._zip_path)
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
            self.status_label.setText(f"下载进度: {mb_d:.1f} MB / {mb_t:.1f} MB")
        else:
            self.progress_bar.setMaximum(0)
            mb_d = downloaded / (1024 * 1024)
            self.status_label.setText(f"已下载: {mb_d:.1f} MB")

    def _on_complete(self, zip_path):
        self.progress_bar.setValue(self.progress_bar.maximum())
        self.status_label.setText("下载完成，正在准备安装...")
        try:
            apply_update_and_restart(zip_path)
            QApplication.quit()
        except Exception as e:
            QMessageBox.warning(
                self,
                "自动更新失败",
                f"自动更新启动失败:\n{e}\n\n更新包位置:\n{zip_path}\n请手动解压覆盖后重启程序。",
            )
            self.accept()

    def _on_failed(self, error):
        self.status_label.setText(f"下载失败: {error}")
        self.update_btn.setEnabled(True)
        self.update_btn.setText("重试")
        self.skip_btn.setText("取消")
        self.skip_btn.clicked.disconnect()
        self.skip_btn.clicked.connect(self.reject)
        self.progress_bar.setVisible(False)

