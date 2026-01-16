# -*- coding: utf-8 -*-
"""
图片处理工具 GUI - v1.2.0
为客户提供简单易用的图片处理工具
"""

# 版本信息
APP_VERSION = "1.2.0"
GITHUB_REPO = "stokisai/ImageProcessingTool"
UPDATE_CHECK_URL = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"

import sys
import os
import subprocess
import threading
from datetime import datetime
from pathlib import Path

from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QFileDialog, QTableWidget, QTableWidgetItem,
    QProgressBar, QTextEdit, QFrame, QSplitter, QMessageBox,
    QHeaderView, QGroupBox, QSizePolicy, QScrollArea, QCheckBox
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QUrl, QTimer
from PyQt5.QtGui import QFont, QColor, QPalette, QDesktopServices, QIcon, QBrush

import pandas as pd
import configparser
import logging
import requests

# 导入处理模块
from image_processor import ImageProcessor
from drive_uploader import DriveUploader
from comfyui_client import ComfyUIClient
from utils import setup_logging, ensure_dir

# 设置日志
logger = setup_logging()


def check_update(parent_window):
    """检查是否有新版本可用"""
    try:
        response = requests.get(UPDATE_CHECK_URL, timeout=5)
        if response.status_code != 200:
            return
        
        data = response.json()
        latest_version = data.get("tag_name", "").lstrip("v")
        download_url = data.get("html_url", "")
        
        # 比较版本号
        if latest_version and latest_version > APP_VERSION:
            reply = QMessageBox.question(
                parent_window,
                "发现新版本",
                f"发现新版本 v{latest_version}！\n当前版本: v{APP_VERSION}\n\n是否前往下载？",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.Yes
            )
            if reply == QMessageBox.Yes:
                QDesktopServices.openUrl(QUrl(download_url))
    except Exception as e:
        # 静默失败，不影响正常使用
        logger.debug(f"检查更新失败: {e}")


class WorkerThread(QThread):
    """后台工作线程"""
    progress_updated = pyqtSignal(int, int, str)  # current, total, message
    log_message = pyqtSignal(str)  # 日志消息
    result_added = pyqtSignal(str, str, str, str)  # folder, filename, status, output_path
    stage_completed = pyqtSignal(str, str, bool)  # stage_name, output_dir, success
    error_occurred = pyqtSignal(str)  # error message
    report_saved = pyqtSignal(str)  # report file path
    
    def __init__(self, mode, task_file, manual_stage2_dir=None, parent=None):
        super().__init__(parent)
        self.mode = mode  # 'stage1', 'stage2', 'full_auto', 'manual_stage2'
        self.task_file = task_file
        self.manual_stage2_dir = manual_stage2_dir  # 手动阶段2的输入目录
        self.stage1_results = {}
        self.stage1_output_dir = None
        self.should_stop = False
        self.report_aggregator = {}  # {folder_name: {"Image 1": link, "Image 2": link, ...}}
        self.folder_image_counts = {}  # {folder_name: current_count}
        
    def log(self, message):
        """发送日志消息"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_message.emit(f"[{timestamp}] {message}")
        logger.info(message)
        
    def run(self):
        try:
            if self.mode == 'stage1':
                self.run_stage1()
            elif self.mode == 'stage2':
                self.run_stage2()
            elif self.mode == 'full_auto':
                self.run_stage1()
                if not self.should_stop and self.stage1_results:
                    self.run_stage2()
            elif self.mode == 'manual_stage2':
                self.run_manual_stage2()
        except Exception as e:
            self.error_occurred.emit(f"处理出错: {str(e)}")
            logger.exception("Worker thread error")
    
    def run_stage1(self):
        """执行阶段1: ComfyUI图生图处理"""
        self.log("开始 阶段1: ComfyUI 图生图处理")
        
        # 读取任务文件
        try:
            df_tasks = pd.read_excel(self.task_file)
        except Exception as e:
            self.error_occurred.emit(f"无法读取任务文件: {e}")
            return
        
        # 收集所有任务
        all_tasks = []
        grouped = df_tasks.groupby(['Folder Name', 'Source Path'], sort=False)
        
        for (folder_name, source_path), group_df in grouped:
            if not os.path.exists(source_path):
                self.log(f"⚠ 源路径不存在: {source_path}")
                continue
                
            folder_images = self._collect_images(source_path)
            task_rows = [row for _, row in group_df.iterrows()]
            
            for idx, (folder_rel, images) in enumerate(folder_images):
                for img_idx, img_path in enumerate(images):
                    row_data = task_rows[min(idx, len(task_rows)-1)] if task_rows else {}
                    
                    task_info = {
                        'source_path': img_path,
                        'img_name': os.path.basename(img_path),
                        'folder_rel_path': folder_rel,
                        'comfyui_url': str(row_data.get('comfyui', '')) if pd.notna(row_data.get('comfyui')) else None,
                        'stage1_dir': str(row_data.get('Processed image 1stage', '')) if pd.notna(row_data.get('Processed image 1stage')) else None,
                        'jp_top': str(row_data.get('Top Text JP', '')) if pd.notna(row_data.get('Top Text JP')) else '',
                        'jp_bottom': str(row_data.get('Bottom Text JP', '')) if pd.notna(row_data.get('Bottom Text JP')) else '',
                        'top_size': int(float(row_data.get('Top Font Size', 0))) if pd.notna(row_data.get('Top Font Size')) else 0,
                        'bottom_size': int(float(row_data.get('Bottom Font Size', 0))) if pd.notna(row_data.get('Bottom Font Size')) else 0,
                        'font_name': str(row_data.get('fonts', '')) if pd.notna(row_data.get('fonts')) else None,
                    }
                    
                    if task_info['jp_top'].lower() == 'nan': task_info['jp_top'] = ''
                    if task_info['jp_bottom'].lower() == 'nan': task_info['jp_bottom'] = ''
                    if task_info['comfyui_url'] and task_info['comfyui_url'].lower() == 'nan': task_info['comfyui_url'] = None
                    if task_info['stage1_dir'] and task_info['stage1_dir'].lower() == 'nan': task_info['stage1_dir'] = None
                    
                    all_tasks.append(task_info)
        
        if not all_tasks:
            self.error_occurred.emit("未找到任何有效任务!")
            return
            
        # 获取全局配置
        global_comfyui_url = None
        global_stage1_dir = None
        for task in all_tasks:
            if task['comfyui_url'] and task['stage1_dir']:
                global_comfyui_url = task['comfyui_url']
                global_stage1_dir = task['stage1_dir']
                break
        
        if not global_comfyui_url or not global_stage1_dir:
            self.error_occurred.emit("未找到ComfyUI配置！请确保Excel中配置了 comfyui 和 Processed image 1stage 列")
            return
        
        self.stage1_output_dir = global_stage1_dir
        self.log(f"输出目录: {global_stage1_dir}")
        self.log(f"总任务数: {len(all_tasks)}")
        
        # 初始化ComfyUI客户端
        try:
            comfyui_client = ComfyUIClient.from_url(global_comfyui_url)
            self.log(f"✓ 已连接ComfyUI")
        except Exception as e:
            self.error_occurred.emit(f"无法连接ComfyUI服务器: {e}")
            return
        
        # 处理图片
        success_count = 0
        for idx, task in enumerate(all_tasks, 1):
            if self.should_stop:
                self.log("用户取消操作")
                return
                
            stage1_subfolder = os.path.join(global_stage1_dir, task['folder_rel_path'])
            ensure_dir(stage1_subfolder)
            stage1_output = os.path.join(stage1_subfolder, task['img_name'])
            
            self.progress_updated.emit(idx, len(all_tasks), f"{task['folder_rel_path']}/{task['img_name']}")
            
            try:
                if comfyui_client.process_image(task['source_path'], stage1_output):
                    self.log(f"✓ ({idx}/{len(all_tasks)}) {task['img_name']}")
                    self.result_added.emit(task['folder_rel_path'], task['img_name'], "成功", stage1_output)
                    self.stage1_results[task['source_path']] = {
                        'output': stage1_output,
                        'task': task
                    }
                    success_count += 1
                else:
                    self.log(f"✗ ({idx}/{len(all_tasks)}) {task['img_name']}")
                    self.result_added.emit(task['folder_rel_path'], task['img_name'], "失败", "")
            except Exception as e:
                self.log(f"✗ ({idx}/{len(all_tasks)}) {task['img_name']} - {str(e)}")
                self.result_added.emit(task['folder_rel_path'], task['img_name'], "错误", "")
        
        self.log(f"阶段1完成: {success_count}/{len(all_tasks)} 成功")
        self.stage_completed.emit("stage1", global_stage1_dir, success_count == len(all_tasks))
    
    def run_stage2(self):
        """执行阶段2: 添加文字标签并上传"""
        self.log("开始 阶段2: 添加文字标签")
        
        if not self.stage1_results:
            self.error_occurred.emit("没有阶段1的处理结果！请先运行阶段1")
            return
        
        processor = ImageProcessor()
        uploader = DriveUploader()
        drive_enabled = uploader.authenticate()
        if drive_enabled:
            self.log("✓ Google Drive认证成功")
        else:
            self.log("⚠ Google Drive认证失败，将跳过上传")
        
        tasks = list(self.stage1_results.values())
        success_count = 0
        self.report_data = []
        
        for idx, item in enumerate(tasks, 1):
            if self.should_stop:
                self.log("用户取消操作")
                return
                
            task = item['task']
            current_img_path = item['output']
            
            self.progress_updated.emit(idx, len(tasks), f"{task['folder_rel_path']}/{task['img_name']}")
            
            # 输出路径
            output_filename = f"{task['folder_rel_path']}_{task['img_name']}".replace(os.sep, "_")
            temp_output_dir = "temp_processed"
            ensure_dir(temp_output_dir)
            processed_path = os.path.join(temp_output_dir, output_filename)
            
            result_link = ""
            try:
                success = processor.process_image(
                    current_img_path, processed_path,
                    task['jp_top'], task['jp_bottom'],
                    top_size=task['top_size'],
                    bottom_size=task['bottom_size'],
                    font_name=task['font_name']
                )
                
                if success:
                    # 上传到Google Drive
                    if drive_enabled:
                        try:
                            # 使用清理过的文件夹名（替换反斜杠）
                            folder_name = task['folder_rel_path'].replace("\\", "_").replace("/", "_")
                            drive_folder_id = uploader.create_folder(folder_name)
                            self.log(f"  Drive文件夹: {folder_name} -> {drive_folder_id}")
                            
                            if drive_folder_id:
                                file_obj = uploader.upload_file(processed_path, drive_folder_id)
                                if file_obj:
                                    result_link = uploader.get_direct_link(file_obj['id'])
                                    self.log(f"  ✓ 已上传: {result_link}")
                                else:
                                    self.log(f"  ⚠ 上传失败")
                            else:
                                self.log(f"  ⚠ 创建Drive文件夹失败")
                        except Exception as upload_err:
                            self.log(f"  ⚠ Drive错误: {str(upload_err)}")
                    
                    self.log(f"✓ ({idx}/{len(tasks)}) {task['img_name']}")
                    self.result_added.emit(task['folder_rel_path'], task['img_name'], "完成", result_link or processed_path)
                    success_count += 1
                else:
                    self.log(f"✗ ({idx}/{len(tasks)}) {task['img_name']}")
                    self.result_added.emit(task['folder_rel_path'], task['img_name'], "失败", "")
            except Exception as e:
                self.log(f"✗ ({idx}/{len(tasks)}) {task['img_name']} - {str(e)}")
                self.result_added.emit(task['folder_rel_path'], task['img_name'], "错误", "")
            
            # 记录报告数据 - 横向格式
            folder_key = task['folder_rel_path'].replace("\\", "_").replace("/", "_")
            if folder_key not in self.report_aggregator:
                self.report_aggregator[folder_key] = {}
                self.folder_image_counts[folder_key] = 0
            self.folder_image_counts[folder_key] += 1
            img_col = f"Image {self.folder_image_counts[folder_key]}"
            self.report_aggregator[folder_key][img_col] = result_link or "Upload Failed"
        
        # 保存报告
        self._save_report()
        
        self.log(f"阶段2完成: {success_count}/{len(tasks)} 成功")
        self.stage_completed.emit("stage2", "temp_processed", success_count == len(tasks))
    
    def run_manual_stage2(self):
        """手动阶段2: 直接从指定目录处理图片"""
        self.log("开始 手动阶段2: 从现有图片添加文字")
        
        if not self.manual_stage2_dir or not os.path.exists(self.manual_stage2_dir):
            self.error_occurred.emit(f"目录不存在: {self.manual_stage2_dir}")
            return
        
        # 读取任务文件获取文案配置
        try:
            df_tasks = pd.read_excel(self.task_file)
        except Exception as e:
            self.error_occurred.emit(f"无法读取任务文件: {e}")
            return
        
        # 收集目录中的图片
        folder_images = self._collect_images(self.manual_stage2_dir)
        if not folder_images:
            self.error_occurred.emit(f"目录中未找到图片: {self.manual_stage2_dir}")
            return
        
        # 获取任务配置
        task_rows = [row for _, row in df_tasks.iterrows()]
        if not task_rows:
            self.error_occurred.emit("Excel中没有任务配置")
            return
        
        processor = ImageProcessor()
        uploader = DriveUploader()
        drive_enabled = uploader.authenticate()
        if drive_enabled:
            self.log("✓ Google Drive认证成功")
        else:
            self.log("⚠ Google Drive认证失败，将跳过上传")
        
        # 构建任务列表
        all_tasks = []
        for folder_rel, images in folder_images:
            for img_path in images:
                row_data = task_rows[0]  # 使用第一行配置
                task_info = {
                    'source_path': img_path,
                    'img_name': os.path.basename(img_path),
                    'folder_rel_path': folder_rel,
                    'jp_top': str(row_data.get('Top Text JP', '')) if pd.notna(row_data.get('Top Text JP')) else '',
                    'jp_bottom': str(row_data.get('Bottom Text JP', '')) if pd.notna(row_data.get('Bottom Text JP')) else '',
                    'top_size': int(float(row_data.get('Top Font Size', 0))) if pd.notna(row_data.get('Top Font Size')) else 0,
                    'bottom_size': int(float(row_data.get('Bottom Font Size', 0))) if pd.notna(row_data.get('Bottom Font Size')) else 0,
                    'font_name': str(row_data.get('fonts', '')) if pd.notna(row_data.get('fonts')) else None,
                }
                if task_info['jp_top'].lower() == 'nan': task_info['jp_top'] = ''
                if task_info['jp_bottom'].lower() == 'nan': task_info['jp_bottom'] = ''
                all_tasks.append(task_info)
        
        self.log(f"找到 {len(all_tasks)} 张图片")
        success_count = 0
        self.report_data = []
        
        for idx, task in enumerate(all_tasks, 1):
            if self.should_stop:
                self.log("用户取消操作")
                return
            
            self.progress_updated.emit(idx, len(all_tasks), f"{task['folder_rel_path']}/{task['img_name']}")
            
            output_filename = f"{task['folder_rel_path']}_{task['img_name']}".replace(os.sep, "_")
            temp_output_dir = "temp_processed"
            ensure_dir(temp_output_dir)
            processed_path = os.path.join(temp_output_dir, output_filename)
            
            result_link = ""
            try:
                success = processor.process_image(
                    task['source_path'], processed_path,
                    task['jp_top'], task['jp_bottom'],
                    top_size=task['top_size'],
                    bottom_size=task['bottom_size'],
                    font_name=task['font_name']
                )
                
                if success:
                    if drive_enabled:
                        try:
                            folder_name = task['folder_rel_path'].replace("\\", "_").replace("/", "_")
                            drive_folder_id = uploader.create_folder(folder_name)
                            if drive_folder_id:
                                file_obj = uploader.upload_file(processed_path, drive_folder_id)
                                if file_obj:
                                    result_link = uploader.get_direct_link(file_obj['id'])
                        except Exception as upload_err:
                            self.log(f"  ⚠ Drive错误: {str(upload_err)}")
                    
                    self.log(f"✓ ({idx}/{len(all_tasks)}) {task['img_name']}")
                    self.result_added.emit(task['folder_rel_path'], task['img_name'], "完成", result_link or processed_path)
                    success_count += 1
                else:
                    self.log(f"✗ ({idx}/{len(all_tasks)}) {task['img_name']}")
                    self.result_added.emit(task['folder_rel_path'], task['img_name'], "失败", "")
            except Exception as e:
                self.log(f"✗ ({idx}/{len(all_tasks)}) {task['img_name']} - {str(e)}")
                self.result_added.emit(task['folder_rel_path'], task['img_name'], "错误", "")
            
            # 记录报告数据 - 横向格式
            folder_key = task['folder_rel_path'].replace("\\", "_").replace("/", "_")
            if folder_key not in self.report_aggregator:
                self.report_aggregator[folder_key] = {}
                self.folder_image_counts[folder_key] = 0
            self.folder_image_counts[folder_key] += 1
            img_col = f"Image {self.folder_image_counts[folder_key]}"
            self.report_aggregator[folder_key][img_col] = result_link or "Upload Failed"
        
        self._save_report()
        self.log(f"手动阶段2完成: {success_count}/{len(all_tasks)} 成功")
        self.stage_completed.emit("manual_stage2", "temp_processed", success_count == len(all_tasks))
    
    def _save_report(self):
        """保存报告到Excel - 横向格式"""
        if not self.report_aggregator:
            return
        
        report_file = "final_report.xlsx"
        try:
            # 转换为横向格式: Folder Name | Image 1 | Image 2 | ...
            final_rows = []
            for folder_name, links_dict in self.report_aggregator.items():
                row_dict = {"Folder Name": folder_name}
                row_dict.update(links_dict)
                final_rows.append(row_dict)
            
            df = pd.DataFrame(final_rows)
            df.to_excel(report_file, index=False)
            self.log(f"✓ 报告已保存: {report_file}")
            self.report_saved.emit(os.path.abspath(report_file))
        except Exception as e:
            self.log(f"⚠ 保存报告失败: {e}")
    
    def _collect_images(self, root_path):
        """收集文件夹中的图片"""
        folder_images = []
        valid_exts = ('.jpg', '.jpeg', '.png')
        
        def process_folder(folder_path):
            images = []
            subdirs = []
            
            try:
                items = sorted(os.listdir(folder_path))
            except Exception:
                return
            
            for item in items:
                full_path = os.path.join(folder_path, item)
                if os.path.isfile(full_path):
                    if item.lower().endswith(valid_exts):
                        if "副本" not in item and "copy" not in item.lower() and "._" not in item and not item.startswith("$"):
                            images.append(full_path)
                elif os.path.isdir(full_path):
                    subdirs.append(full_path)
            
            if images:
                rel_folder = os.path.relpath(folder_path, root_path)
                if rel_folder == ".":
                    rel_folder = os.path.basename(root_path)
                folder_images.append((rel_folder, sorted(images)))
            
            for subdir in subdirs:
                process_folder(subdir)
        
        process_folder(root_path)
        return folder_images
    
    def stop(self):
        self.should_stop = True


class MainWindow(QMainWindow):
    """主窗口"""
    
    def __init__(self):
        super().__init__()
        self.worker = None
        self.task_file = None
        self.current_output_dir = None
        self.report_file = None
        self.init_ui()
        
        # 启动后延迟2秒检查更新，不影响UI加载
        QTimer.singleShot(2000, lambda: check_update(self))
        
    def init_ui(self):
        """初始化界面"""
        self.setWindowTitle(f"📷 图片处理工具 v{APP_VERSION}")
        self.setMinimumSize(850, 650)
        self.resize(950, 700)
        
        # 应用样式
        self.apply_styles()
        
        # 主窗口部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        main_layout = QVBoxLayout(central_widget)
        main_layout.setSpacing(12)
        main_layout.setContentsMargins(20, 15, 20, 15)
        
        # 标题
        title_label = QLabel("📷 图片处理工具")
        title_label.setObjectName("titleLabel")
        title_label.setAlignment(Qt.AlignCenter)
        main_layout.addWidget(title_label)
        
        # 文件选择区域
        file_layout = QHBoxLayout()
        file_layout.setSpacing(10)
        
        file_icon = QLabel("📁")
        file_icon.setStyleSheet("font-size: 18px;")
        file_layout.addWidget(file_icon)
        
        self.file_label = QLabel("请选择Excel任务文件...")
        self.file_label.setObjectName("fileLabel")
        self.file_label.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        file_layout.addWidget(self.file_label)
        
        browse_btn = QPushButton("浏览...")
        browse_btn.setObjectName("browseBtn")
        browse_btn.clicked.connect(self.browse_file)
        file_layout.addWidget(browse_btn)
        
        main_layout.addLayout(file_layout)
        
        # 操作按钮区域 - 上排
        btn_layout1 = QHBoxLayout()
        btn_layout1.setSpacing(15)
        
        self.stage1_btn = QPushButton("阶段1\nComfyUI处理")
        self.stage1_btn.setObjectName("stage1Btn")
        self.stage1_btn.setMinimumHeight(55)
        self.stage1_btn.clicked.connect(self.run_stage1)
        self.stage1_btn.setEnabled(False)
        btn_layout1.addWidget(self.stage1_btn)
        
        self.stage2_btn = QPushButton("阶段2\n添加文字")
        self.stage2_btn.setObjectName("stage2Btn")
        self.stage2_btn.setMinimumHeight(55)
        self.stage2_btn.clicked.connect(self.run_stage2)
        self.stage2_btn.setEnabled(False)
        btn_layout1.addWidget(self.stage2_btn)
        
        self.auto_btn = QPushButton("全流程自动\n无需确认")
        self.auto_btn.setObjectName("autoBtn")
        self.auto_btn.setMinimumHeight(55)
        self.auto_btn.clicked.connect(self.run_full_auto)
        self.auto_btn.setEnabled(False)
        btn_layout1.addWidget(self.auto_btn)
        
        main_layout.addLayout(btn_layout1)
        
        # 操作按钮区域 - 下排（手动阶段2）
        btn_layout2 = QHBoxLayout()
        btn_layout2.setSpacing(15)
        
        btn_layout2.addStretch()
        
        self.manual_stage2_btn = QPushButton("手动阶段2\n选择已有图片文件夹")
        self.manual_stage2_btn.setObjectName("manualStage2Btn")
        self.manual_stage2_btn.setMinimumHeight(45)
        self.manual_stage2_btn.setMinimumWidth(200)
        self.manual_stage2_btn.clicked.connect(self.run_manual_stage2)
        self.manual_stage2_btn.setEnabled(False)
        btn_layout2.addWidget(self.manual_stage2_btn)
        
        btn_layout2.addStretch()
        
        main_layout.addLayout(btn_layout2)
        
        # 进度条和状态
        progress_layout = QHBoxLayout()
        
        # 运行状态指示器
        self.running_indicator = QLabel("●")
        self.running_indicator.setObjectName("runningIndicator")
        self.running_indicator.setFixedWidth(25)
        self.running_indicator.setVisible(False)
        progress_layout.addWidget(self.running_indicator)
        
        # 动画计时器
        self.indicator_timer = QTimer()
        self.indicator_timer.timeout.connect(self.animate_indicator)
        self.indicator_colors = ["#22c55e", "#4ade80", "#86efac", "#4ade80"]
        self.indicator_index = 0
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setObjectName("progressBar")
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setFormat("%p%")
        self.progress_bar.setMinimumHeight(22)
        progress_layout.addWidget(self.progress_bar)
        
        self.status_label = QLabel("等待开始...")
        self.status_label.setObjectName("statusLabel")
        self.status_label.setMinimumWidth(200)
        progress_layout.addWidget(self.status_label)
        
        # 停止按钮
        self.stop_btn = QPushButton("⏹ 停止")
        self.stop_btn.setObjectName("stopBtn")
        self.stop_btn.setFixedWidth(70)
        self.stop_btn.clicked.connect(self.stop_processing)
        self.stop_btn.setVisible(False)
        progress_layout.addWidget(self.stop_btn)
        
        main_layout.addLayout(progress_layout)
        
        # 处理结果表格
        result_label = QLabel("📋 处理结果")
        result_label.setObjectName("sectionLabel")
        main_layout.addWidget(result_label)
        
        self.result_table = QTableWidget()
        self.result_table.setObjectName("resultTable")
        self.result_table.setColumnCount(4)
        self.result_table.setHorizontalHeaderLabels(["序号", "文件", "状态", "输出/链接"])
        self.result_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Fixed)
        self.result_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.result_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Fixed)
        self.result_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.Stretch)
        self.result_table.setColumnWidth(0, 50)
        self.result_table.setColumnWidth(2, 70)
        self.result_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.result_table.verticalHeader().setVisible(False)
        main_layout.addWidget(self.result_table, 1)
        
        # 完成状态栏
        self.complete_frame = QFrame()
        self.complete_frame.setObjectName("completeFrame")
        self.complete_frame.setVisible(False)
        complete_layout = QHBoxLayout(self.complete_frame)
        complete_layout.setContentsMargins(15, 12, 15, 12)
        
        complete_left = QVBoxLayout()
        self.complete_label = QLabel()
        self.complete_label.setObjectName("completeLabel")
        complete_left.addWidget(self.complete_label)
        
        self.output_path_label = QLabel()
        self.output_path_label.setObjectName("outputPathLabel")
        complete_left.addWidget(self.output_path_label)
        
        self.report_label = QLabel()
        self.report_label.setObjectName("reportLabel")
        complete_left.addWidget(self.report_label)
        
        complete_layout.addLayout(complete_left)
        complete_layout.addStretch()
        
        btn_layout = QVBoxLayout()
        self.open_folder_btn = QPushButton("📂 打开输出文件夹")
        self.open_folder_btn.setObjectName("openFolderBtn")
        self.open_folder_btn.clicked.connect(self.open_output_folder)
        btn_layout.addWidget(self.open_folder_btn)
        
        self.open_report_btn = QPushButton("📊 打开报告Excel")
        self.open_report_btn.setObjectName("openReportBtn")
        self.open_report_btn.clicked.connect(self.open_report)
        btn_layout.addWidget(self.open_report_btn)
        
        self.open_report_folder_btn = QPushButton("📁 打开报告文件夹")
        self.open_report_folder_btn.setObjectName("openReportFolderBtn")
        self.open_report_folder_btn.clicked.connect(self.open_report_folder)
        btn_layout.addWidget(self.open_report_folder_btn)
        
        complete_layout.addLayout(btn_layout)
        
        main_layout.addWidget(self.complete_frame)
        
    def apply_styles(self):
        """应用样式表"""
        style = """
        QMainWindow {
            background: qlineargradient(x1:0, y1:0, x2:1, y2:1,
                stop:0 #1e293b, stop:0.5 #1e3a5f, stop:1 #172554);
        }
        
        QLabel#titleLabel {
            font-size: 22px;
            font-weight: bold;
            color: #ffffff;
            padding: 3px;
        }
        
        QLabel#sectionLabel {
            font-size: 13px;
            font-weight: bold;
            color: #94a3b8;
            padding: 3px 0;
        }
        
        QLabel#fileLabel {
            font-size: 12px;
            color: #cbd5e1;
            padding: 8px 12px;
            background: rgba(15, 23, 42, 0.6);
            border-radius: 6px;
            border: 1px solid rgba(148, 163, 184, 0.2);
        }
        
        QLabel#statusLabel {
            font-size: 11px;
            color: #94a3b8;
            padding-left: 10px;
        }
        
        QLabel#completeLabel {
            font-size: 15px;
            font-weight: bold;
            color: #4ade80;
        }
        
        QLabel#outputPathLabel, QLabel#reportLabel {
            font-size: 12px;
            color: #93c5fd;
        }
        
        QPushButton {
            font-size: 12px;
            font-weight: bold;
            color: white;
            background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                stop:0 #6366f1, stop:1 #4f46e5);
            border: none;
            border-radius: 8px;
            padding: 10px 18px;
        }
        
        QPushButton:hover {
            background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                stop:0 #818cf8, stop:1 #6366f1);
        }
        
        QPushButton:pressed {
            background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                stop:0 #4f46e5, stop:1 #4338ca);
        }
        
        QPushButton:disabled {
            background: rgba(100, 116, 139, 0.4);
            color: rgba(255, 255, 255, 0.4);
        }
        
        QPushButton#browseBtn {
            padding: 8px 14px;
            font-size: 11px;
        }
        
        QPushButton#stage1Btn {
            background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                stop:0 #0ea5e9, stop:1 #0284c7);
        }
        QPushButton#stage1Btn:hover {
            background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                stop:0 #38bdf8, stop:1 #0ea5e9);
        }
        
        QPushButton#stage2Btn {
            background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                stop:0 #f59e0b, stop:1 #d97706);
        }
        QPushButton#stage2Btn:hover {
            background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                stop:0 #fbbf24, stop:1 #f59e0b);
        }
        
        QPushButton#autoBtn {
            background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                stop:0 #22c55e, stop:1 #16a34a);
        }
        QPushButton#autoBtn:hover {
            background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                stop:0 #4ade80, stop:1 #22c55e);
        }
        
        QPushButton#manualStage2Btn {
            background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                stop:0 #a855f7, stop:1 #9333ea);
            font-size: 11px;
        }
        QPushButton#manualStage2Btn:hover {
            background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                stop:0 #c084fc, stop:1 #a855f7);
        }
        
        QPushButton#openFolderBtn, QPushButton#openReportBtn, QPushButton#openReportFolderBtn {
            background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                stop:0 #3b82f6, stop:1 #2563eb);
            padding: 8px 12px;
            font-size: 11px;
            margin: 2px;
        }
        
        QPushButton#stopBtn {
            background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                stop:0 #ef4444, stop:1 #dc2626);
            padding: 5px 10px;
            font-size: 11px;
        }
        QPushButton#stopBtn:hover {
            background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                stop:0 #f87171, stop:1 #ef4444);
        }
        
        QLabel#runningIndicator {
            font-size: 18px;
            color: #22c55e;
        }
        
        QTableWidget#resultTable {
            background: rgba(15, 23, 42, 0.7);
            color: #e2e8f0;
            border: 1px solid rgba(148, 163, 184, 0.15);
            border-radius: 6px;
            gridline-color: rgba(148, 163, 184, 0.1);
            font-size: 12px;
        }
        
        QTableWidget#resultTable::item {
            padding: 6px 8px;
            border-bottom: 1px solid rgba(148, 163, 184, 0.1);
        }
        
        QTableWidget#resultTable::item:selected {
            background: rgba(99, 102, 241, 0.3);
        }
        
        QHeaderView::section {
            background: rgba(51, 65, 85, 0.8);
            color: #f1f5f9;
            font-weight: bold;
            font-size: 11px;
            padding: 8px;
            border: none;
            border-bottom: 2px solid rgba(99, 102, 241, 0.5);
        }
        
        QProgressBar {
            border: none;
            border-radius: 5px;
            background: rgba(15, 23, 42, 0.6);
            text-align: center;
            color: white;
            font-weight: bold;
            font-size: 11px;
        }
        
        QProgressBar::chunk {
            background: qlineargradient(x1:0, y1:0, x2:1, y2:0,
                stop:0 #6366f1, stop:0.5 #8b5cf6, stop:1 #d946ef);
            border-radius: 5px;
        }
        
        QFrame#completeFrame {
            background: rgba(34, 197, 94, 0.12);
            border: 2px solid rgba(34, 197, 94, 0.4);
            border-radius: 10px;
        }
        
        QScrollBar:vertical {
            background: rgba(15, 23, 42, 0.4);
            width: 10px;
            border-radius: 5px;
        }
        
        QScrollBar::handle:vertical {
            background: rgba(148, 163, 184, 0.3);
            border-radius: 5px;
            min-height: 20px;
        }
        
        QScrollBar::handle:vertical:hover {
            background: rgba(148, 163, 184, 0.5);
        }
        
        QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
            height: 0px;
        }
        """
        self.setStyleSheet(style)
        
    def browse_file(self):
        """浏览并选择任务文件"""
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择任务文件", "", "Excel文件 (*.xlsx *.xls)"
        )
        
        if file_path:
            self.task_file = file_path
            self.file_label.setText(file_path)
            self.stage1_btn.setEnabled(True)
            self.auto_btn.setEnabled(True)
            self.manual_stage2_btn.setEnabled(True)
            self.result_table.setRowCount(0)
            self.complete_frame.setVisible(False)
            
    def run_stage1(self):
        """运行阶段1"""
        if not self.task_file:
            QMessageBox.warning(self, "警告", "请先选择任务文件！")
            return
        self.result_table.setRowCount(0)
        self.start_worker('stage1')
        
    def run_stage2(self):
        """运行阶段2"""
        if not self.worker or not self.worker.stage1_results:
            QMessageBox.warning(self, "警告", "请先完成阶段1！")
            return
        self.start_worker('stage2')
        
    def run_full_auto(self):
        """运行全自动流程"""
        if not self.task_file:
            QMessageBox.warning(self, "警告", "请先选择任务文件！")
            return
        if not self.check_old_report():
            return
        self.result_table.setRowCount(0)
        self.start_worker('full_auto')
    
    def run_manual_stage2(self):
        """手动阶段2: 选择已有图片文件夹"""
        if not self.task_file:
            QMessageBox.warning(self, "警告", "请先选择任务文件！")
            return
        
        # 提示用户
        reply = QMessageBox.information(
            self, "手动阶段2",
            "请选择包含已处理图片的文件夹。\n\n"
            "注意：请确保文件夹中存在要处理的图片文件（jpg/png）。\n"
            "程序将使用Excel中的文案配置对图片添加文字标签。",
            QMessageBox.Ok | QMessageBox.Cancel
        )
        
        if reply != QMessageBox.Ok:
            return
        
        # 选择文件夹
        folder_path = QFileDialog.getExistingDirectory(
            self, "选择图片文件夹", ""
        )
        
        if folder_path:
            if not self.check_old_report():
                return
            self.result_table.setRowCount(0)
            self.start_worker('manual_stage2', folder_path)
        
    def start_worker(self, mode, manual_dir=None):
        """启动工作线程"""
        self.set_buttons_enabled(False)
        self.complete_frame.setVisible(False)
        self.progress_bar.setValue(0)
        
        # 显示运行指示器和停止按钮
        self.running_indicator.setVisible(True)
        self.stop_btn.setVisible(True)
        self.indicator_timer.start(300)  # 每300ms切换颜色
        
        # 如果是stage2且有之前的结果，复用它
        if mode == 'stage2' and self.worker and self.worker.stage1_results:
            old_results = self.worker.stage1_results
            old_output_dir = self.worker.stage1_output_dir
            self.worker = WorkerThread(mode, self.task_file)
            self.worker.stage1_results = old_results
            self.worker.stage1_output_dir = old_output_dir
        else:
            self.worker = WorkerThread(mode, self.task_file, manual_dir)
        
        self.worker.progress_updated.connect(self.update_progress)
        self.worker.log_message.connect(self.append_log)
        self.worker.result_added.connect(self.add_result_row)
        self.worker.stage_completed.connect(self.on_stage_completed)
        self.worker.error_occurred.connect(self.on_error)
        self.worker.report_saved.connect(self.on_report_saved)
        self.worker.finished.connect(self.on_worker_finished)
        
        self.worker.start()
        
    def set_buttons_enabled(self, enabled):
        """设置按钮启用状态"""
        self.stage1_btn.setEnabled(enabled and self.task_file is not None)
        self.stage2_btn.setEnabled(enabled and self.worker is not None and bool(self.worker.stage1_results))
        self.auto_btn.setEnabled(enabled and self.task_file is not None)
        self.manual_stage2_btn.setEnabled(enabled and self.task_file is not None)
        
    def update_progress(self, current, total, message):
        """更新进度"""
        self.progress_bar.setMaximum(total)
        self.progress_bar.setValue(current)
        self.status_label.setText(message)
        
    def append_log(self, message):
        """添加日志"""
        print(message)
        
    def add_result_row(self, folder, filename, status, output_path):
        """添加结果行到表格"""
        row = self.result_table.rowCount()
        self.result_table.insertRow(row)
        
        # 序号
        num_item = QTableWidgetItem(str(row + 1))
        num_item.setTextAlignment(Qt.AlignCenter)
        num_item.setForeground(QBrush(QColor("#94a3b8")))
        self.result_table.setItem(row, 0, num_item)
        
        # 文件 (文件夹/文件名)
        file_item = QTableWidgetItem(f"{folder}/{filename}")
        file_item.setForeground(QBrush(QColor("#e2e8f0")))
        self.result_table.setItem(row, 1, file_item)
        
        # 状态
        status_item = QTableWidgetItem(status)
        status_item.setTextAlignment(Qt.AlignCenter)
        if "成功" in status or "完成" in status:
            status_item.setForeground(QBrush(QColor("#4ade80")))
            status_item.setBackground(QBrush(QColor(34, 197, 94, 30)))
        else:
            status_item.setForeground(QBrush(QColor("#f87171")))
            status_item.setBackground(QBrush(QColor(248, 113, 113, 30)))
        self.result_table.setItem(row, 2, status_item)
        
        # 输出/链接
        output_item = QTableWidgetItem(output_path)
        output_item.setForeground(QBrush(QColor("#93c5fd")))
        self.result_table.setItem(row, 3, output_item)
        
        self.result_table.scrollToBottom()
            
    def on_stage_completed(self, stage_name, output_dir, success):
        """阶段完成处理"""
        self.current_output_dir = output_dir
        
        if stage_name == "stage1":
            self.complete_label.setText("✅ 阶段1已完成！请检查输出目录确认图片质量。")
            self.output_path_label.setText(f"输出目录: {output_dir}")
            self.report_label.setText("")
            self.complete_frame.setVisible(True)
            self.stage2_btn.setEnabled(True)
            # 阶段1不显示报告按钮
            self.open_report_btn.setVisible(False)
            self.open_report_folder_btn.setVisible(False)
        elif stage_name in ("stage2", "manual_stage2"):
            self.complete_label.setText("✅ 全部完成！图片已上传到Google Drive。")
            self.output_path_label.setText(f"输出目录: {output_dir}")
            self.complete_frame.setVisible(True)
            # 阶段2显示报告按钮
            self.open_report_btn.setVisible(True)
            self.open_report_folder_btn.setVisible(True)
    
    def on_report_saved(self, report_path):
        """报告保存完成"""
        self.report_file = report_path
        self.report_label.setText(f"报告文件: {os.path.basename(report_path)}")
            
    def on_error(self, error_message):
        """错误处理"""
        QMessageBox.critical(self, "错误", error_message)
        self.status_label.setText(f"错误")
        
    def on_worker_finished(self):
        """工作线程完成"""
        self.set_buttons_enabled(True)
        # 隐藏运行指示器
        self.running_indicator.setVisible(False)
        self.stop_btn.setVisible(False)
        self.indicator_timer.stop()
        self.status_label.setText("完成")
    
    def animate_indicator(self):
        """动画更新运行指示器颜色"""
        self.indicator_index = (self.indicator_index + 1) % len(self.indicator_colors)
        color = self.indicator_colors[self.indicator_index]
        self.running_indicator.setStyleSheet(f"font-size: 18px; color: {color};")
    
    def stop_processing(self):
        """停止处理并提示清理"""
        if not self.worker or not self.worker.isRunning():
            return
        
        # 停止工作线程
        self.worker.stop()
        self.worker.wait(2000)
        
        # 隐藏指示器
        self.running_indicator.setVisible(False)
        self.stop_btn.setVisible(False)
        self.indicator_timer.stop()
        self.status_label.setText("已停止")
        self.progress_bar.setValue(0)  # 重置进度条
        self.set_buttons_enabled(True)
        
        # 获取当前输出目录
        output_dir = self.current_output_dir or "temp_processed"
        output_path = os.path.abspath(output_dir)
        
        # 提示清理对话框
        msg = QMessageBox(self)
        msg.setIcon(QMessageBox.Warning)
        msg.setWindowTitle("任务已停止")
        msg.setText(f"处理已中断！\n\n"
                   f"此任务产生的临时文件可能需要清理：\n"
                   f"📁 {output_path}\n\n"
                   f"建议删除这些文件以避免数据混乱。")
        
        delete_btn = msg.addButton("🗑️ 删除全部", QMessageBox.DestructiveRole)
        open_btn = msg.addButton("📂 打开文件夹", QMessageBox.ActionRole)
        close_btn = msg.addButton("关闭", QMessageBox.RejectRole)
        
        msg.exec_()
        
        clicked = msg.clickedButton()
        if clicked == delete_btn:
            try:
                import shutil
                if os.path.exists(output_path):
                    shutil.rmtree(output_path)
                    QMessageBox.information(self, "成功", f"已删除: {output_path}")
            except Exception as e:
                QMessageBox.warning(self, "删除失败", f"无法删除: {e}")
        elif clicked == open_btn:
            if os.path.exists(output_path):
                subprocess.run(['explorer', output_path])
            else:
                QMessageBox.warning(self, "警告", f"目录不存在: {output_path}")
    
    def open_output_folder(self):
        """打开输出文件夹"""
        if self.current_output_dir:
            path = os.path.abspath(self.current_output_dir)
            if os.path.exists(path):
                subprocess.run(['explorer', path])
            else:
                QMessageBox.warning(self, "警告", f"目录不存在: {path}")
    
    def open_report(self):
        """打开报告Excel"""
        if self.report_file and os.path.exists(self.report_file):
            os.startfile(self.report_file)
        else:
            QMessageBox.warning(self, "警告", "报告文件不存在")
    
    def open_report_folder(self):
        """打开报告所在文件夹"""
        report_path = os.path.abspath("final_report.xlsx")
        folder = os.path.dirname(report_path)
        if os.path.exists(folder):
            subprocess.run(['explorer', folder])
        else:
            QMessageBox.warning(self, "警告", f"目录不存在: {folder}")
    
    def check_old_report(self):
        """检查旧报告文件，提示删除以避免数据混乱"""
        report_path = os.path.abspath("final_report.xlsx")
        if os.path.exists(report_path):
            # 创建自定义对话框
            msg = QMessageBox(self)
            msg.setIcon(QMessageBox.Warning)
            msg.setWindowTitle("检测到旧报告")
            msg.setText("发现已存在的报告文件 final_report.xlsx\n\n"
                       "继续运行可能导致新旧数据混合。\n"
                       "建议删除旧报告后再开始新任务。")
            
            delete_btn = msg.addButton("🗑️ 删除旧报告", QMessageBox.DestructiveRole)
            open_btn = msg.addButton("📂 打开文件夹", QMessageBox.ActionRole)
            cancel_btn = msg.addButton("取消", QMessageBox.RejectRole)
            continue_btn = msg.addButton("继续运行", QMessageBox.AcceptRole)
            
            msg.exec_()
            
            clicked = msg.clickedButton()
            if clicked == delete_btn:
                try:
                    os.remove(report_path)
                    QMessageBox.information(self, "成功", "旧报告已删除！")
                    return True
                except Exception as e:
                    QMessageBox.warning(self, "删除失败", f"无法删除文件: {e}")
                    return False
            elif clicked == open_btn:
                subprocess.run(['explorer', '/select,', report_path])
                return False  # 用户需要手动处理后重新点击
            elif clicked == continue_btn:
                return True  # 用户选择继续
            else:
                return False  # 取消
        return True  # 没有旧报告，可以继续
    
    def closeEvent(self, event):
        """关闭事件"""
        if self.worker and self.worker.isRunning():
            reply = QMessageBox.question(
                self, "确认退出",
                "正在处理中，确定要退出吗？",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No
            )
            
            if reply == QMessageBox.Yes:
                self.worker.stop()
                self.worker.wait(3000)
                event.accept()
            else:
                event.ignore()
        else:
            event.accept()


def main():
    """主函数"""
    app = QApplication(sys.argv)
    
    font = QFont("Microsoft YaHei UI", 10)
    app.setFont(font)
    
    window = MainWindow()
    window.show()
    
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
