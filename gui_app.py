# -*- coding: utf-8 -*-
"""
图片处理工具 GUI - v1.1.3
为客户提供简单易用的图片处理工具
"""

# 版本信息
APP_VERSION = "1.1.3"
GITHUB_REPO = "stokisai/wuli"

import sys
import os
import math
import shutil
import subprocess
import threading
from datetime import datetime
from pathlib import Path

from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QFileDialog, QTableWidget, QTableWidgetItem,
    QProgressBar, QTextEdit, QFrame, QSplitter, QMessageBox,
    QHeaderView, QGroupBox, QSizePolicy, QScrollArea, QCheckBox,
    QStackedWidget, QLineEdit, QFormLayout, QComboBox, QInputDialog,
    QDialog, QGridLayout
)
from PySide6.QtCore import Qt, QThread, Signal, QUrl, QTimer, QObject
from PySide6.QtGui import QFont, QColor, QPalette, QDesktopServices, QIcon, QBrush, QTextCursor, QPixmap

import pandas as pd
import configparser
import logging
import requests

# 导入处理模块
from image_processor import ImageProcessor
from oss_uploader import OSSUploader
from comfyui_client import ComfyUIClient
from utils import setup_logging, ensure_dir
from updater import UpdateCheckWorker, UpdateDialog

# 设置日志
logger = setup_logging()


class WorkerThread(QThread):
    """后台工作线程"""
    progress_updated = Signal(int, int, str)  # current, total, message
    log_message = Signal(str)  # 日志消息
    result_added = Signal(str, str, str, str)  # folder, filename, status, output_path
    stage_completed = Signal(str, str, bool)  # stage_name, output_dir, success
    error_occurred = Signal(str)  # error message
    report_saved = Signal(str)  # report file path
    
    def __init__(self, mode, task_file, manual_stage2_dir=None, comfyui_url=None, source_path=None, stage1_output_dir=None, workflow_path=None, parent=None):
        super().__init__(parent)
        self.mode = mode  # 'stage1', 'stage2', 'full_auto', 'manual_stage2'
        self.task_file = task_file
        self.manual_stage2_dir = manual_stage2_dir  # 手动阶段2的输入目录
        self.comfyui_url = comfyui_url  # 全局 ComfyUI 地址
        self.source_path = source_path  # 全局图片源路径
        self.stage1_results = {}
        self.stage1_output_dir = stage1_output_dir  # Global stage1 output path from config
        self.workflow_path = workflow_path  # 用户选择的工作流路径
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

        # 使用全局 source_path
        source_path = self.source_path
        if not source_path:
            self.error_occurred.emit("未配置图片源路径！请在「配置」页面设置图片源路径")
            return
        if not os.path.exists(source_path):
            self.error_occurred.emit(f"图片源路径不存在: {source_path}")
            return

        grouped = df_tasks.groupby(['Folder Name'], sort=False)

        for folder_name, group_df in grouped:
            folder_images = self._collect_images(source_path)
            task_rows = [row for _, row in group_df.iterrows()]
            
            for idx, (folder_rel, images) in enumerate(folder_images):
                for img_idx, img_path in enumerate(images):
                    row_data = task_rows[min(idx, len(task_rows)-1)] if task_rows else {}
                    
                    excel_stage1_dir = None
                    if 'Processed image 1stage' in row_data and pd.notna(row_data.get('Processed image 1stage')):
                        excel_stage1_dir = str(row_data.get('Processed image 1stage')).strip()
                        if excel_stage1_dir.lower() == 'nan':
                            excel_stage1_dir = None

                    task_info = {
                        'source_path': img_path,
                        'img_name': os.path.basename(img_path),
                        'folder_rel_path': folder_rel,
                        'stage1_dir': excel_stage1_dir,
                        'jp_top': str(row_data.get('Top Text JP', '')) if pd.notna(row_data.get('Top Text JP')) else '',
                        'jp_bottom': str(row_data.get('Bottom Text JP', '')) if pd.notna(row_data.get('Bottom Text JP')) else '',
                        'top_size': int(float(row_data.get('Top Font Size', 0))) if pd.notna(row_data.get('Top Font Size')) else 0,
                        'bottom_size': int(float(row_data.get('Bottom Font Size', 0))) if pd.notna(row_data.get('Bottom Font Size')) else 0,
                        'font_name': str(row_data.get('fonts', '')) if pd.notna(row_data.get('fonts')) else None,
                    }
                    
                    if task_info['jp_top'].lower() == 'nan': task_info['jp_top'] = ''
                    if task_info['jp_bottom'].lower() == 'nan': task_info['jp_bottom'] = ''
                    if task_info['stage1_dir'] and task_info['stage1_dir'].lower() == 'nan': task_info['stage1_dir'] = None
                    
                    all_tasks.append(task_info)
        
        if not all_tasks:
            self.error_occurred.emit("未找到任何有效任务!")
            return
            
        # 使用全局 ComfyUI 地址
        global_comfyui_url = self.comfyui_url
        global_stage1_dir = (self.stage1_output_dir or "").strip()
        stage1_dir_from = "Config"
        if not global_stage1_dir:
            for task in all_tasks:
                if task['stage1_dir']:
                    global_stage1_dir = task['stage1_dir']
                    stage1_dir_from = "Excel"
                    break

        if not global_comfyui_url:
            self.error_occurred.emit("ComfyUI URL is not configured. Please set it in the Config page.")
            return
        if not global_stage1_dir:
            self.error_occurred.emit("Stage1 output directory is missing. Please set it in Config (legacy fallback: Excel column 'Processed image 1stage').")
            return
        if os.path.isfile(global_stage1_dir):
            self.error_occurred.emit(f"Stage1 output path is a file, not a folder: {global_stage1_dir}")
            return
        try:
            ensure_dir(global_stage1_dir)
        except Exception as e:
            self.error_occurred.emit(f"Failed to create stage1 output directory: {e}")
            return

        self.stage1_output_dir = global_stage1_dir
        self.log(f"Output directory ({stage1_dir_from}): {global_stage1_dir}")
        self.log(f"Total tasks: {len(all_tasks)}")

        # Initialize ComfyUI client
        try:
            comfyui_client = ComfyUIClient.from_url(global_comfyui_url)
            if self.workflow_path:
                comfyui_client.load_workflow(self.workflow_path)
                self.log(f"✓ 已加载工作流: {os.path.basename(self.workflow_path)}")
            # 真正检查连接
            if not comfyui_client.check_connection():
                self.error_occurred.emit(f"无法连接ComfyUI服务器: {global_comfyui_url}")
                return
            self.log(f"✓ 已连接ComfyUI: {global_comfyui_url}")
        except Exception as e:
            self.error_occurred.emit(f"无法连接ComfyUI服务器: {e}")
            return
        
        # 处理图片
        success_count = 0
        skipped_a_count = 0
        skipped_b_count = 0
        
        for idx, task in enumerate(all_tasks, 1):
            if self.should_stop:
                self.log("用户取消操作")
                return
            
            img_name_lower = task['img_name'].lower()
            img_stem_lower = os.path.splitext(img_name_lower)[0]

            # 规则A: 文件名(不含扩展名)为 'a' -> 完全跳过
            if img_stem_lower == 'a':
                self.progress_updated.emit(idx, len(all_tasks), f"跳过: {task['img_name']}")
                self.log(f"⏭ ({idx}/{len(all_tasks)}) {task['img_name']} - 跳过(规则A)")
                self.result_added.emit(task['folder_rel_path'], task['img_name'], "跳过A", "")
                skipped_a_count += 1
                continue

            # 规则B: 文件名(不含扩展名)为 'b' -> 跳过ComfyUI，复制原图到Stage1文件夹
            if img_stem_lower == 'b':
                self.progress_updated.emit(idx, len(all_tasks), f"复制: {task['img_name']}")
                # 创建Stage1子文件夹并复制原图
                stage1_subfolder = os.path.join(global_stage1_dir, task['folder_rel_path'])
                ensure_dir(stage1_subfolder)
                stage1_output = os.path.join(stage1_subfolder, task['img_name'])
                
                try:
                    import shutil
                    shutil.copy2(task['source_path'], stage1_output)
                    self.log(f"⏭ ({idx}/{len(all_tasks)}) {task['img_name']} - 跳过ComfyUI(规则B)，原图已复制到Stage1")
                    self.result_added.emit(task['folder_rel_path'], task['img_name'], "跳过ComfyUI", stage1_output)
                    # 使用复制后的路径
                    self.stage1_results[task['source_path']] = {
                        'output': stage1_output,
                        'task': task
                    }
                    skipped_b_count += 1
                    success_count += 1
                except Exception as copy_err:
                    self.log(f"✗ ({idx}/{len(all_tasks)}) {task['img_name']} - 复制失败: {copy_err}")
                    self.result_added.emit(task['folder_rel_path'], task['img_name'], "复制失败", "")
                continue
                
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
        
        self.log(f"阶段1完成: {success_count}/{len(all_tasks)} 成功 (跳过A:{skipped_a_count}, 跳过ComfyUI-B:{skipped_b_count})")
        self.stage_completed.emit("stage1", global_stage1_dir, success_count == len(all_tasks))
    
    def run_stage2(self):
        """执行阶段2: 添加文字标签并上传"""
        self.log("开始 阶段2: 添加文字标签")
        
        if not self.stage1_results:
            self.error_occurred.emit("没有阶段1的处理结果！请先运行阶段1")
            return
        
        processor = ImageProcessor()
        uploader = OSSUploader()
        oss_enabled = uploader.authenticate()
        if oss_enabled:
            self.log("✓ 阿里云 OSS 认证成功")
        else:
            self.log("⚠ 阿里云 OSS 认证失败，将跳过上传")
        
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
                    # 上传到阿里云 OSS
                    if oss_enabled:
                        try:
                            # 使用清理过的文件夹名（替换反斜杠）
                            folder_name = task['folder_rel_path'].replace("\\", "_").replace("/", "_")
                            oss_folder = uploader.create_folder(folder_name)
                            self.log(f"  OSS文件夹: {oss_folder}")
                            
                            if oss_folder:
                                file_obj = uploader.upload_file(processed_path, oss_folder)
                                if file_obj:
                                    result_link = uploader.get_direct_link(file_obj['id'])
                                    self.log(f"  ✓ 已上传: {result_link}")
                                else:
                                    self.log(f"  ⚠ 上传失败")
                            else:
                                self.log(f"  ⚠ 创建OSS文件夹失败")
                        except Exception as upload_err:
                            self.log(f"  ⚠ OSS错误: {str(upload_err)}")
                    
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
        self.stage_completed.emit("stage2", os.path.abspath("temp_processed"), success_count == len(tasks))
    
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
        uploader = OSSUploader()
        oss_enabled = uploader.authenticate()
        if oss_enabled:
            self.log("✓ 阿里云 OSS 认证成功")
        else:
            self.log("⚠ 阿里云 OSS 认证失败，将跳过上传")
        
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
                    if oss_enabled:
                        try:
                            folder_name = task['folder_rel_path'].replace("\\", "_").replace("/", "_")
                            oss_folder = uploader.create_folder(folder_name)
                            if oss_folder:
                                file_obj = uploader.upload_file(processed_path, oss_folder)
                                if file_obj:
                                    result_link = uploader.get_direct_link(file_obj['id'])
                        except Exception as upload_err:
                            self.log(f"  ⚠ OSS错误: {str(upload_err)}")
                    
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
        self.stage_completed.emit("manual_stage2", os.path.abspath("temp_processed"), success_count == len(all_tasks))
    
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


class LogSignalBridge(QObject):
    """Bridge Python logging to Qt UI thread."""
    message = Signal(str)


class GuiLogHandler(logging.Handler):
    """Logging handler forwarding formatted logs to Qt signal."""

    def __init__(self, bridge: LogSignalBridge):
        super().__init__()
        self.bridge = bridge

    def emit(self, record):
        try:
            msg = self.format(record)
        except Exception:
            msg = record.getMessage()
        self.bridge.message.emit(msg)


class ComfyUIConnectionTestWorker(QThread):
    """ComfyUI connection test worker."""
    check_finished = Signal(bool, str, str)  # ok, tested_url, message

    def __init__(self, url, parent=None):
        super().__init__(parent)
        self.url = url

    def run(self):
        try:
            client = ComfyUIClient.from_url(self.url)
            if client.check_connection():
                self.check_finished.emit(True, self.url, "连接成功，ComfyUI 服务可用")
            else:
                self.check_finished.emit(False, self.url, "连接失败，请确认地址、端口和 ComfyUI 服务状态")
        except Exception as e:
            self.check_finished.emit(False, self.url, f"连接异常: {e}")

class ClickableLabel(QLabel):
    """可点击的图片标签"""
    clicked = Signal(str)

    def __init__(self, image_path="", parent=None):
        super().__init__(parent)
        self._image_path = image_path
        self.setCursor(Qt.PointingHandCursor)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton and self._image_path:
            self.clicked.emit(self._image_path)
        super().mousePressEvent(event)


class ThumbnailLoader(QThread):
    """后台加载缩略图线程"""
    thumbnail_ready = Signal(int, object)  # index, QPixmap

    def __init__(self, image_paths, size=280, parent=None):
        super().__init__(parent)
        self.image_paths = image_paths
        self.size = size

    def run(self):
        for idx, path in enumerate(self.image_paths):
            try:
                pixmap = QPixmap(path)
                if not pixmap.isNull():
                    pixmap = pixmap.scaled(
                        self.size, self.size,
                        Qt.KeepAspectRatio,
                        Qt.SmoothTransformation,
                    )
                self.thumbnail_ready.emit(idx, pixmap)
            except Exception:
                self.thumbnail_ready.emit(idx, QPixmap())


class ImagePreviewDialog(QDialog):
    """大图预览窗口"""

    def __init__(self, image_path, parent=None):
        super().__init__(parent)
        self.setObjectName("imagePreviewDialog")
        self.setWindowTitle(os.path.basename(image_path))
        self.setModal(True)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(8, 8, 8, 8)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setAlignment(Qt.AlignCenter)

        img_label = QLabel()
        pixmap = QPixmap(image_path)
        if not pixmap.isNull():
            screen = self.screen()
            if screen:
                avail = screen.availableGeometry()
                max_w = int(avail.width() * 0.85)
                max_h = int(avail.height() * 0.85)
            else:
                max_w, max_h = 1200, 800
            pixmap = pixmap.scaled(max_w, max_h, Qt.KeepAspectRatio, Qt.SmoothTransformation)
        img_label.setPixmap(pixmap)
        img_label.setAlignment(Qt.AlignCenter)
        scroll.setWidget(img_label)
        layout.addWidget(scroll, 1)

        close_btn = QPushButton("关闭")
        close_btn.clicked.connect(self.accept)
        layout.addWidget(close_btn, 0, Qt.AlignRight)

        self.resize(pixmap.width() + 40, pixmap.height() + 80)


class ReprocessWorkerThread(QThread):
    """用新工作流重新处理选中图片"""
    progress_updated = Signal(int, int, str)  # current, total, message
    image_done = Signal(str, bool)  # output_path, success
    all_done = Signal(bool)  # overall success

    def __init__(self, selected_outputs, source_map, comfyui_url, workflow_path, parent=None):
        super().__init__(parent)
        self.selected_outputs = selected_outputs  # list of output paths
        self.source_map = source_map  # {output_path: source_path}
        self.comfyui_url = comfyui_url
        self.workflow_path = workflow_path
        self.should_stop = False

    def run(self):
        try:
            client = ComfyUIClient.from_url(self.comfyui_url)
            if self.workflow_path:
                client.load_workflow(self.workflow_path)
            if not client.check_connection():
                self.all_done.emit(False)
                return

            total = len(self.selected_outputs)
            success_count = 0
            for idx, output_path in enumerate(self.selected_outputs, 1):
                if self.should_stop:
                    break
                source_path = self.source_map.get(output_path, "")
                if not source_path or not os.path.exists(source_path):
                    self.progress_updated.emit(idx, total, f"源文件缺失: {os.path.basename(output_path)}")
                    self.image_done.emit(output_path, False)
                    continue

                self.progress_updated.emit(idx, total, os.path.basename(output_path))
                ok = client.process_image(source_path, output_path)
                self.image_done.emit(output_path, ok)
                if ok:
                    success_count += 1

            self.all_done.emit(success_count == total)
        except Exception as e:
            logger.exception("ReprocessWorkerThread error")
            self.all_done.emit(False)

    def stop(self):
        self.should_stop = True


class ImageGalleryDialog(QDialog):
    """Stage1 图库预览 + 选图重处理"""

    def __init__(self, image_paths, source_map, comfyui_url,
                 current_workflow_name, workflows_dir, parent=None):
        super().__init__(parent)
        self.setObjectName("galleryDialog")
        self.setWindowTitle("图库 - Stage1 输出结果")
        self.setWindowFlags(
            self.windowFlags()
            | Qt.WindowMinimizeButtonHint
            | Qt.WindowMaximizeButtonHint
        )
        self.setModal(False)
        self.resize(1100, 800)

        self._image_paths = list(image_paths)
        self._source_map = dict(source_map)
        self._comfyui_url = comfyui_url
        self._current_workflow_name = current_workflow_name
        self._workflows_dir = Path(workflows_dir)
        self._checkboxes = []
        self._thumb_labels = []
        self._reprocess_worker = None

        self._build_ui()
        self._build_gallery_grid()
        self._start_thumbnail_loader(self._image_paths)

    # ---- UI construction ----

    def _build_ui(self):
        root = QVBoxLayout(self)
        root.setSpacing(10)
        root.setContentsMargins(16, 16, 16, 16)

        # Title
        title = QLabel("图库 - Stage1 输出结果")
        title.setObjectName("pageTitle")
        root.addWidget(title)

        # Stats bar
        stats_bar = QHBoxLayout()
        self._total_label = QLabel(f"共 {len(self._image_paths)} 张")
        self._total_label.setObjectName("sectionLabel")
        stats_bar.addWidget(self._total_label)

        self._selected_label = QLabel("已选择: 0 张")
        self._selected_label.setObjectName("sectionLabel")
        stats_bar.addWidget(self._selected_label)

        stats_bar.addStretch()
        select_all_btn = QPushButton("全选")
        select_all_btn.clicked.connect(self._select_all)
        stats_bar.addWidget(select_all_btn)

        deselect_all_btn = QPushButton("取消全选")
        deselect_all_btn.clicked.connect(self._deselect_all)
        stats_bar.addWidget(deselect_all_btn)
        root.addLayout(stats_bar)

        # Scroll area for gallery grid
        self._scroll = QScrollArea()
        self._scroll.setWidgetResizable(True)
        self._scroll.setObjectName("galleryScroll")
        self._grid_widget = QWidget()
        self._grid_layout = QGridLayout(self._grid_widget)
        self._grid_layout.setSpacing(12)
        self._scroll.setWidget(self._grid_widget)
        root.addWidget(self._scroll, 1)

        # Reprocess frame
        reprocess_frame = QFrame()
        reprocess_frame.setObjectName("reprocessFrame")
        rp_layout = QVBoxLayout(reprocess_frame)
        rp_layout.setContentsMargins(12, 12, 12, 12)
        rp_layout.setSpacing(8)

        wf_row = QHBoxLayout()
        wf_row.addWidget(QLabel("工作流:"))
        self._wf_combo = QComboBox()
        self._wf_combo.setObjectName("configInput")
        self._wf_combo.setMinimumHeight(32)
        self._populate_workflow_combo()
        wf_row.addWidget(self._wf_combo, 1)
        rp_layout.addLayout(wf_row)

        hint = QLabel(f"当前使用: {self._current_workflow_name}  |  建议尝试其他工作流")
        hint.setStyleSheet("color: #94a3b8; font-size: 12px;")
        rp_layout.addWidget(hint)

        action_row = QHBoxLayout()
        self._reprocess_btn = QPushButton("重新处理选中图片 (0 张)")
        self._reprocess_btn.setObjectName("reprocessBtn")
        self._reprocess_btn.setMinimumHeight(38)
        self._reprocess_btn.clicked.connect(self._on_reprocess)
        action_row.addWidget(self._reprocess_btn)

        self._rp_progress = QProgressBar()
        self._rp_progress.setTextVisible(True)
        self._rp_progress.setFormat("%v/%m  %p%")
        self._rp_progress.setVisible(False)
        action_row.addWidget(self._rp_progress, 1)
        rp_layout.addLayout(action_row)

        # 详细进度状态行
        self._rp_status_label = QLabel("")
        self._rp_status_label.setStyleSheet("color: #94a3b8; font-size: 12px; padding: 2px 0;")
        self._rp_status_label.setVisible(False)
        rp_layout.addWidget(self._rp_status_label)

        root.addWidget(reprocess_frame)

        # Bottom buttons
        bottom = QHBoxLayout()
        bottom.addStretch()
        save_btn = QPushButton("保存关闭")
        save_btn.setObjectName("saveConfigBtn")
        save_btn.setMinimumHeight(36)
        save_btn.clicked.connect(self._on_save)
        bottom.addWidget(save_btn)
        root.addLayout(bottom)

    def _populate_workflow_combo(self):
        self._wf_combo.clear()
        names = sorted(p.stem for p in self._workflows_dir.glob("*.json"))
        self._wf_combo.addItems(names)
        # Try to select a different workflow than current
        for i, n in enumerate(names):
            if n != self._current_workflow_name:
                self._wf_combo.setCurrentIndex(i)
                break

    def _build_gallery_grid(self):
        """Build 3-column grid of image cards."""
        # Clear existing
        while self._grid_layout.count():
            item = self._grid_layout.takeAt(0)
            w = item.widget()
            if w:
                w.deleteLater()
        self._checkboxes.clear()
        self._thumb_labels.clear()

        cols = 3
        for idx, path in enumerate(self._image_paths):
            row, col = divmod(idx, cols)

            cell = QFrame()
            cell.setObjectName("galleryCell")
            cell_layout = QVBoxLayout(cell)
            cell_layout.setContentsMargins(8, 8, 8, 8)
            cell_layout.setSpacing(4)

            cb = QCheckBox()
            cb.setObjectName("galleryCheckbox")
            cb.stateChanged.connect(self._on_checkbox_changed)
            self._checkboxes.append(cb)
            cell_layout.addWidget(cb)

            thumb = ClickableLabel(path)
            thumb.setFixedSize(280, 280)
            thumb.setAlignment(Qt.AlignCenter)
            thumb.setText("加载中...")
            thumb.setStyleSheet("color: #94a3b8;")
            thumb.clicked.connect(self._on_image_clicked)
            self._thumb_labels.append(thumb)
            cell_layout.addWidget(thumb, 0, Qt.AlignCenter)

            name_label = QLabel(os.path.basename(path))
            name_label.setAlignment(Qt.AlignCenter)
            name_label.setWordWrap(True)
            name_label.setStyleSheet("color: #cbd5e1; font-size: 12px;")
            cell_layout.addWidget(name_label)

            self._grid_layout.addWidget(cell, row, col)

    def _start_thumbnail_loader(self, paths):
        self._thumb_loader = ThumbnailLoader(paths, size=280, parent=self)
        self._thumb_loader.thumbnail_ready.connect(self._on_thumbnail_ready)
        self._thumb_loader.start()

    # ---- Slots ----

    def _on_thumbnail_ready(self, idx, pixmap):
        if 0 <= idx < len(self._thumb_labels):
            label = self._thumb_labels[idx]
            if pixmap and not pixmap.isNull():
                label.setPixmap(pixmap)
                label.setText("")
            else:
                label.setText("加载失败")

    def _on_image_clicked(self, path):
        dlg = ImagePreviewDialog(path, self)
        dlg.exec()

    def _on_checkbox_changed(self):
        count = sum(1 for cb in self._checkboxes if cb.isChecked())
        self._selected_label.setText(f"已选择: {count} 张")
        self._reprocess_btn.setText(f"重新处理选中图片 ({count} 张)")

    def _select_all(self):
        for cb in self._checkboxes:
            cb.setChecked(True)

    def _deselect_all(self):
        for cb in self._checkboxes:
            cb.setChecked(False)

    def _on_reprocess(self):
        selected = [
            self._image_paths[i]
            for i, cb in enumerate(self._checkboxes)
            if cb.isChecked()
        ]
        if not selected:
            QMessageBox.information(self, "提示", "请先勾选需要重新处理的图片")
            return

        wf_name = self._wf_combo.currentText()
        if not wf_name:
            QMessageBox.warning(self, "警告", "请选择一个工作流")
            return
        wf_path = str(self._workflows_dir / f"{wf_name}.json")

        self._reprocess_btn.setEnabled(False)
        self._rp_progress.setVisible(True)
        self._rp_progress.setValue(0)
        self._rp_progress.setMaximum(len(selected))
        self._rp_status_label.setVisible(True)
        self._rp_status_label.setText("准备中...")
        self.setWindowTitle("图库 - 重处理中...")

        self._reprocess_worker = ReprocessWorkerThread(
            selected, self._source_map, self._comfyui_url, wf_path, self
        )
        self._reprocess_worker.progress_updated.connect(self._on_rp_progress)
        self._reprocess_worker.all_done.connect(self._on_reprocess_complete)
        self._reprocess_worker.start()

    def _on_rp_progress(self, current, total, msg):
        self._rp_progress.setMaximum(total)
        self._rp_progress.setValue(current)
        self._rp_status_label.setText(f"正在处理 ({current}/{total}): {msg}")
        self.setWindowTitle(f"图库 - 重处理中 {current}/{total}")

    def _on_reprocess_complete(self, success):
        self._reprocess_btn.setEnabled(True)
        self._rp_progress.setVisible(False)
        self._rp_status_label.setVisible(False)
        self._reprocess_worker = None
        self.setWindowTitle("图库 - Stage1 输出结果")

        if success:
            QMessageBox.information(self, "完成", "选中图片已全部重新处理！")
        else:
            QMessageBox.warning(self, "提示", "部分图片重新处理失败，请检查日志。")

        # Refresh gallery showing only reprocessed images
        reprocessed = [
            self._image_paths[i]
            for i, cb in enumerate(self._checkboxes)
            if cb.isChecked()
        ]
        self._refresh_gallery(reprocessed)

    def _refresh_gallery(self, paths):
        """Rebuild gallery with a new set of paths."""
        self._image_paths = list(paths)
        self._total_label.setText(f"共 {len(self._image_paths)} 张")
        self._build_gallery_grid()
        self._start_thumbnail_loader(self._image_paths)
        self._on_checkbox_changed()

    def _on_save(self):
        self.accept()

    def closeEvent(self, event):
        """关闭时检查是否正在重处理"""
        if self._reprocess_worker and self._reprocess_worker.isRunning():
            reply = QMessageBox.question(
                self, "正在处理中",
                "图片正在重新处理中，关闭窗口将在后台继续。\n\n"
                "确定要关闭吗？",
                QMessageBox.Yes | QMessageBox.No,
                QMessageBox.No,
            )
            if reply != QMessageBox.Yes:
                event.ignore()
                return
        event.accept()

class MainWindow(QMainWindow):
    """主窗口"""

    def __init__(self):
        super().__init__()
        self.worker = None
        self.task_file = None
        self.current_output_dir = None
        self.report_file = None
        self._update_checker = None
        self._comfyui_test_worker = None
        self._comfyui_test_ok = False
        self._comfyui_tested_url = ""
        self._log_bridge = None
        self._gui_log_handler = None
        self._runtime_log_max_lines = 6000
        self._last_progress_marker = None
        self._stage1_workflow_name = ""
        self._comfyui_glow_timer = None
        self._comfyui_glow_step = 0

        self.init_ui()
        self._init_runtime_log_capture()
        self._load_existing_log_file()
        self._load_saved_task_file()

        # ?????2???????
        self._update_check_silent = True
        QTimer.singleShot(2000, lambda: self._check_for_updates(silent=True))
    def init_ui(self):
        """Initialize UI with left sidebar navigation and right content area."""
        self.setWindowTitle(f"\u56fe\u7247\u5904\u7406\u5de5\u5177 v{APP_VERSION}")
        self.setMinimumSize(1400, 900)
        self.resize(1600, 1000)

        central_widget = QWidget()
        central_widget.setObjectName("appRoot")
        self.setCentralWidget(central_widget)

        root_layout = QHBoxLayout(central_widget)
        root_layout.setSpacing(12)
        root_layout.setContentsMargins(12, 12, 12, 12)

        self.sidebar = QFrame()
        self.sidebar.setObjectName("sidebar")
        self.sidebar.setFixedWidth(220)
        sidebar_layout = QVBoxLayout(self.sidebar)
        sidebar_layout.setSpacing(8)
        sidebar_layout.setContentsMargins(12, 12, 12, 12)

        sidebar_title = QLabel("\u56fe\u7247\u5904\u7406\u5de5\u5177")
        sidebar_title.setObjectName("sidebarTitle")
        sidebar_layout.addWidget(sidebar_title)

        self.nav_tool_btn = QPushButton("\u56fe\u7247\u5904\u7406\u5de5\u5177")
        self.nav_tool_btn.setObjectName("navButton")
        self.nav_tool_btn.setCheckable(True)
        self.nav_tool_btn.clicked.connect(lambda: self.switch_page(0))
        sidebar_layout.addWidget(self.nav_tool_btn)

        self.nav_info_btn = QPushButton("\u914d\u7f6e")
        self.nav_info_btn.setObjectName("navButton")
        self.nav_info_btn.setCheckable(True)
        self.nav_info_btn.clicked.connect(lambda: self.switch_page(1))
        sidebar_layout.addWidget(self.nav_info_btn)

        sidebar_layout.addStretch()

        version_label = QLabel(f"v{APP_VERSION}")
        version_label.setObjectName("sidebarVersion")
        version_label.setAlignment(Qt.AlignCenter)
        sidebar_layout.addWidget(version_label)

        self.update_check_btn = QPushButton("检查更新")
        self.update_check_btn.setObjectName("updateCheckBtn")
        self.update_check_btn.clicked.connect(lambda: self._check_for_updates(silent=False))
        sidebar_layout.addWidget(self.update_check_btn)

        content_frame = QFrame()
        content_frame.setObjectName("contentArea")
        content_layout = QVBoxLayout(content_frame)
        content_layout.setContentsMargins(0, 0, 0, 0)
        content_layout.setSpacing(0)

        self.page_stack = QStackedWidget()
        self.page_stack.setObjectName("contentStack")
        content_layout.addWidget(self.page_stack)

        root_layout.addWidget(self.sidebar)
        root_layout.addWidget(content_frame, 1)

        tool_page = QWidget()
        tool_page.setObjectName("toolPage")
        main_layout = QVBoxLayout(tool_page)
        main_layout.setSpacing(12)
        main_layout.setContentsMargins(16, 16, 16, 16)

        title_label = QLabel("\u56fe\u7247\u5904\u7406\u5de5\u5177")
        title_label.setObjectName("pageTitle")
        title_label.setAlignment(Qt.AlignLeft | Qt.AlignVCenter)
        main_layout.addWidget(title_label)

        file_layout = QHBoxLayout()
        file_layout.setSpacing(10)

        file_icon = QLabel("\u4efb\u52a1\u6587\u4ef6")
        file_icon.setObjectName("fileIcon")
        file_layout.addWidget(file_icon)

        self.file_label = QLabel("\u8bf7\u9009\u62e9 Excel \u4efb\u52a1\u6587\u4ef6...")
        self.file_label.setObjectName("fileLabel")
        self.file_label.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        self.file_label.setMaximumWidth(600)
        file_layout.addWidget(self.file_label)

        browse_btn = QPushButton("\u6d4f\u89c8...")
        browse_btn.setObjectName("browseBtn")
        browse_btn.clicked.connect(self.browse_file)
        file_layout.addWidget(browse_btn)

        self.save_task_btn = QPushButton("保存")
        self.save_task_btn.setObjectName("saveConfigBtn")
        self.save_task_btn.setMinimumHeight(30)
        self.save_task_btn.clicked.connect(self._save_task_file_path)
        file_layout.addWidget(self.save_task_btn)

        main_layout.addLayout(file_layout)

        btn_layout1 = QHBoxLayout()
        btn_layout1.setSpacing(10)

        self.stage1_btn = QPushButton("\u9636\u6bb51\nComfyUI \u5904\u7406")
        self.stage1_btn.setObjectName("stage1Btn")
        self.stage1_btn.setMinimumHeight(54)
        self.stage1_btn.clicked.connect(self.run_stage1)
        self.stage1_btn.setEnabled(False)
        btn_layout1.addWidget(self.stage1_btn)

        self.stage2_btn = QPushButton("\u9636\u6bb52\n\u6dfb\u52a0\u6587\u5b57")
        self.stage2_btn.setObjectName("stage2Btn")
        self.stage2_btn.setMinimumHeight(54)
        self.stage2_btn.clicked.connect(self.run_stage2)
        self.stage2_btn.setEnabled(False)
        btn_layout1.addWidget(self.stage2_btn)

        self.auto_btn = QPushButton("\u5168\u6d41\u7a0b\u81ea\u52a8\n\u65e0\u9700\u786e\u8ba4")
        self.auto_btn.setObjectName("autoBtn")
        self.auto_btn.setMinimumHeight(54)
        self.auto_btn.clicked.connect(self.run_full_auto)
        self.auto_btn.setEnabled(False)
        btn_layout1.addWidget(self.auto_btn)

        main_layout.addLayout(btn_layout1)

        btn_layout2 = QHBoxLayout()
        btn_layout2.addStretch()

        self.manual_stage2_btn = QPushButton("\u624b\u52a8\u9636\u6bb52\n\u9009\u62e9\u5df2\u6709\u56fe\u7247\u6587\u4ef6\u5939")
        self.manual_stage2_btn.setObjectName("manualStage2Btn")
        self.manual_stage2_btn.setMinimumHeight(44)
        self.manual_stage2_btn.setMinimumWidth(220)
        self.manual_stage2_btn.clicked.connect(self.run_manual_stage2)
        self.manual_stage2_btn.setEnabled(False)
        btn_layout2.addWidget(self.manual_stage2_btn)

        btn_layout2.addStretch()
        main_layout.addLayout(btn_layout2)

        progress_layout = QHBoxLayout()

        self.running_indicator = QLabel("o")
        self.running_indicator.setObjectName("runningIndicator")
        self.running_indicator.setFixedWidth(20)
        self.running_indicator.setVisible(False)
        progress_layout.addWidget(self.running_indicator)

        self.indicator_timer = QTimer()
        self.indicator_timer.timeout.connect(self.animate_indicator)
        self._pulse_step = 0
        self._running_btn = None
        self._btn_pulse_on = False
        # 每个按钮的呼吸灯颜色主题: (dim_bg, bright_bg, dim_border, bright_border)
        self._btn_color_themes = {
            'stage1Btn':       ((30, 58, 95),  (79, 140, 255), (45, 90, 142),  (122, 180, 255)),
            'stage2Btn':       ((20, 83, 45),  (34, 197, 94),  (26, 122, 66),  (74, 222, 128)),
            'autoBtn':         ((59, 31, 110), (139, 92, 246), (91, 58, 158),  (167, 139, 250)),
            'manualStage2Btn': ((45, 55, 72),  (100, 130, 170),(74, 85, 104),  (140, 165, 200)),
        }

        self.progress_bar = QProgressBar()
        self.progress_bar.setObjectName("progressBar")
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setFormat("%p%")
        self.progress_bar.setMinimumHeight(22)
        progress_layout.addWidget(self.progress_bar, 1)

        self.status_label = QLabel("\u7b49\u5f85\u5f00\u59cb...")
        self.status_label.setObjectName("statusLabel")
        self.status_label.setMinimumWidth(200)
        progress_layout.addWidget(self.status_label)

        self.stop_btn = QPushButton("\u505c\u6b62")
        self.stop_btn.setObjectName("stopBtn")
        self.stop_btn.setFixedWidth(70)
        self.stop_btn.clicked.connect(self.stop_processing)
        self.stop_btn.setVisible(False)
        progress_layout.addWidget(self.stop_btn)

        main_layout.addLayout(progress_layout)

        result_label = QLabel("\u5904\u7406\u7ed3\u679c")
        result_label.setObjectName("sectionLabel")
        main_layout.addWidget(result_label)

        self.result_table = QTableWidget()
        self.result_table.setObjectName("resultTable")
        self.result_table.setColumnCount(4)
        self.result_table.setHorizontalHeaderLabels(["\u5e8f\u53f7", "\u6587\u4ef6", "\u72b6\u6001", "\u8f93\u51fa/\u94fe\u63a5"])
        self.result_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Fixed)
        self.result_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.result_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Fixed)
        self.result_table.horizontalHeader().setSectionResizeMode(3, QHeaderView.Stretch)
        self.result_table.setColumnWidth(0, 50)
        self.result_table.setColumnWidth(2, 70)
        self.result_table.setSelectionBehavior(QTableWidget.SelectRows)
        self.result_table.verticalHeader().setVisible(False)
        main_layout.addWidget(self.result_table, 1)

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
        self.open_folder_btn = QPushButton("\u6253\u5f00\u8f93\u51fa\u6587\u4ef6\u5939")
        self.open_folder_btn.setObjectName("openFolderBtn")
        self.open_folder_btn.clicked.connect(self.open_output_folder)
        btn_layout.addWidget(self.open_folder_btn)

        self.open_report_btn = QPushButton("\u6253\u5f00\u62a5\u544a Excel")
        self.open_report_btn.setObjectName("openReportBtn")
        self.open_report_btn.clicked.connect(self.open_report)
        btn_layout.addWidget(self.open_report_btn)

        self.open_report_folder_btn = QPushButton("\u6253\u5f00\u62a5\u544a\u76ee\u5f55")
        self.open_report_folder_btn.setObjectName("openReportFolderBtn")
        self.open_report_folder_btn.clicked.connect(self.open_report_folder)
        btn_layout.addWidget(self.open_report_folder_btn)

        self.gallery_btn = QPushButton("查看图库")
        self.gallery_btn.setObjectName("openFolderBtn")
        self.gallery_btn.clicked.connect(self._open_gallery)
        self.gallery_btn.setVisible(False)
        btn_layout.addWidget(self.gallery_btn)

        complete_layout.addLayout(btn_layout)

        main_layout.addWidget(self.complete_frame)

        info_page = QWidget()
        info_page.setObjectName("infoPage")
        info_layout = QVBoxLayout(info_page)
        info_layout.setContentsMargins(24, 24, 24, 24)
        info_layout.setSpacing(10)

        info_title = QLabel("\u914d\u7f6e")
        info_title.setObjectName("pageTitle")
        info_layout.addWidget(info_title)

        # ComfyUI 全局端口设置
        comfyui_group = QGroupBox("ComfyUI 全局端口")
        comfyui_group.setObjectName("configGroup")
        comfyui_form = QHBoxLayout(comfyui_group)
        comfyui_form.setContentsMargins(12, 12, 12, 12)

        comfyui_label = QLabel("ComfyUI 地址:")
        comfyui_label.setObjectName("configLabel")
        comfyui_form.addWidget(comfyui_label)

        self.comfyui_url_input = QLineEdit()
        self.comfyui_url_input.setObjectName("configInput")
        self.comfyui_url_input.setMinimumHeight(36)
        self.comfyui_url_input.setPlaceholderText("例如: http://127.0.0.1:8188")
        self.comfyui_url_input.textChanged.connect(self._on_comfyui_url_changed)

        # 读取 config.ini 已保存的值（无保存则留空）
        _, parser = self._read_runtime_config()
        saved_host = parser.get("ComfyUI", "Host", fallback="")
        saved_port = parser.get("ComfyUI", "DefaultPort", fallback="")
        if saved_host and saved_port:
            scheme = "https" if saved_port in ("443",) else "http"
            self.comfyui_url_input.setText(f"{scheme}://{saved_host}:{saved_port}")
        comfyui_form.addWidget(self.comfyui_url_input, 1)

        self.test_comfyui_btn = QPushButton("测试连接")
        self.test_comfyui_btn.setObjectName("testConfigBtn")
        self.test_comfyui_btn.setMinimumHeight(36)
        self.test_comfyui_btn.setMinimumWidth(92)
        self.test_comfyui_btn.clicked.connect(self._test_comfyui_connection)
        comfyui_form.addWidget(self.test_comfyui_btn)

        self.save_comfyui_btn = QPushButton("保存")
        self.save_comfyui_btn.setObjectName("saveConfigBtn")
        self.save_comfyui_btn.setMinimumHeight(36)
        self.save_comfyui_btn.setMinimumWidth(72)
        self.save_comfyui_btn.setEnabled(False)
        self.save_comfyui_btn.clicked.connect(self._save_comfyui_url)
        comfyui_form.addWidget(self.save_comfyui_btn)

        self.comfyui_status_label = QLabel("")
        self.comfyui_status_label.setObjectName("configStatus")
        self.comfyui_status_label.setWordWrap(True)
        self.comfyui_status_label.setProperty("state", "pending")

        info_layout.addWidget(comfyui_group)
        info_layout.addWidget(self.comfyui_status_label)
        self._on_comfyui_url_changed(self.comfyui_url_input.text())
        
        info_layout.addSpacing(10)

        # ========== 选择工作流 ==========
        workflow_group = QGroupBox("选择工作流")
        workflow_group.setObjectName("configGroup")
        workflow_form = QHBoxLayout(workflow_group)
        workflow_form.setContentsMargins(12, 12, 12, 12)

        wf_label = QLabel("工作流:")
        wf_label.setObjectName("configLabel")
        workflow_form.addWidget(wf_label)

        self.workflow_combo = QComboBox()
        self.workflow_combo.setObjectName("configInput")
        self.workflow_combo.setMinimumHeight(36)
        self._refresh_workflow_combo()
        # 从 config.ini 恢复上次选择
        saved_wf = parser.get("ComfyUI", "SelectedWorkflow", fallback="默认工作流")
        idx = self.workflow_combo.findText(saved_wf)
        if idx >= 0:
            self.workflow_combo.setCurrentIndex(idx)
        workflow_form.addWidget(self.workflow_combo, 1)

        self.upload_workflow_btn = QPushButton("上传")
        self.upload_workflow_btn.setObjectName("testConfigBtn")
        self.upload_workflow_btn.setMinimumHeight(36)
        self.upload_workflow_btn.setMinimumWidth(72)
        self.upload_workflow_btn.clicked.connect(self._upload_workflow)
        workflow_form.addWidget(self.upload_workflow_btn)

        self.delete_workflow_btn = QPushButton("删除")
        self.delete_workflow_btn.setObjectName("testConfigBtn")
        self.delete_workflow_btn.setMinimumHeight(36)
        self.delete_workflow_btn.setMinimumWidth(72)
        self.delete_workflow_btn.clicked.connect(self._delete_workflow)
        workflow_form.addWidget(self.delete_workflow_btn)

        self.save_workflow_btn = QPushButton("保存")
        self.save_workflow_btn.setObjectName("saveConfigBtn")
        self.save_workflow_btn.setMinimumHeight(36)
        self.save_workflow_btn.setMinimumWidth(72)
        self.save_workflow_btn.clicked.connect(self._save_workflow_selection)
        workflow_form.addWidget(self.save_workflow_btn)

        info_layout.addWidget(workflow_group)

        info_layout.addSpacing(10)

        # ========== 图片源路径配置 ==========
        source_group = QGroupBox("图片源路径")
        source_group.setObjectName("configGroup")
        source_form = QHBoxLayout(source_group)
        source_form.setContentsMargins(12, 12, 12, 12)

        source_label = QLabel("源路径:")
        source_label.setObjectName("configLabel")
        source_form.addWidget(source_label)

        self.source_path_input = QLineEdit()
        self.source_path_input.setObjectName("configInput")
        self.source_path_input.setMinimumHeight(36)
        self.source_path_input.setPlaceholderText("选择图片源文件夹路径...")
        saved_source = parser.get("Paths", "SourcePath", fallback="")
        if saved_source:
            self.source_path_input.setText(saved_source)
        source_form.addWidget(self.source_path_input, 1)

        self.browse_source_btn = QPushButton("浏览")
        self.browse_source_btn.setObjectName("testConfigBtn")
        self.browse_source_btn.setMinimumHeight(36)
        self.browse_source_btn.setMinimumWidth(72)
        self.browse_source_btn.clicked.connect(self._browse_source_path)
        source_form.addWidget(self.browse_source_btn)

        self.save_source_btn = QPushButton("保存")
        self.save_source_btn.setObjectName("saveConfigBtn")
        self.save_source_btn.setMinimumHeight(36)
        self.save_source_btn.setMinimumWidth(72)
        self.save_source_btn.clicked.connect(self._save_source_path)
        source_form.addWidget(self.save_source_btn)

        info_layout.addWidget(source_group)
        
        info_layout.addSpacing(10)


        # ========== Stage1 Output Path ==========
        stage1_dir_group = QGroupBox("阶段1 输出路径")
        stage1_dir_group.setObjectName("configGroup")
        stage1_dir_form = QHBoxLayout(stage1_dir_group)
        stage1_dir_form.setContentsMargins(12, 12, 12, 12)

        stage1_dir_label = QLabel("输出路径:")
        stage1_dir_label.setObjectName("configLabel")
        stage1_dir_form.addWidget(stage1_dir_label)

        self.stage1_output_input = QLineEdit()
        self.stage1_output_input.setObjectName("configInput")
        self.stage1_output_input.setMinimumHeight(36)
        self.stage1_output_input.setPlaceholderText("选择阶段1输出文件夹路径...")
        saved_stage1_output = parser.get("Paths", "Stage1OutputPath", fallback="")
        if saved_stage1_output:
            self.stage1_output_input.setText(saved_stage1_output)
        stage1_dir_form.addWidget(self.stage1_output_input, 1)

        self.browse_stage1_output_btn = QPushButton("浏览")
        self.browse_stage1_output_btn.setObjectName("testConfigBtn")
        self.browse_stage1_output_btn.setMinimumHeight(36)
        self.browse_stage1_output_btn.setMinimumWidth(72)
        self.browse_stage1_output_btn.clicked.connect(self._browse_stage1_output_dir)
        stage1_dir_form.addWidget(self.browse_stage1_output_btn)

        self.save_stage1_output_btn = QPushButton("保存")
        self.save_stage1_output_btn.setObjectName("saveConfigBtn")
        self.save_stage1_output_btn.setMinimumHeight(36)
        self.save_stage1_output_btn.setMinimumWidth(72)
        self.save_stage1_output_btn.clicked.connect(self._save_stage1_output_dir)
        stage1_dir_form.addWidget(self.save_stage1_output_btn)

        self.clear_stage1_btn = QPushButton("清空")
        self.clear_stage1_btn.setObjectName("clearDangerBtn")
        self.clear_stage1_btn.setMinimumHeight(36)
        self.clear_stage1_btn.setMinimumWidth(72)
        self.clear_stage1_btn.clicked.connect(self._clear_stage1_output_dir)
        stage1_dir_form.addWidget(self.clear_stage1_btn)

        info_layout.addWidget(stage1_dir_group)
        
        info_layout.addSpacing(10)


        oss_group = QGroupBox("阿里云 OSS 配置 (只读)")
        oss_group.setObjectName("configGroup")
        # 使用 QFormLayout 对齐标签和输入框
        oss_form_layout = QFormLayout(oss_group)
        oss_form_layout.setContentsMargins(20, 20, 20, 20)
        oss_form_layout.setSpacing(15)
        oss_form_layout.setLabelAlignment(Qt.AlignRight)

        # 读取 OSS 配置
        oss_endpoint = parser.get("OSS", "Endpoint", fallback="未配置")
        oss_bucket = parser.get("OSS", "Bucket", fallback="未配置")
        oss_key_id = parser.get("OSS", "AccessKeyId", fallback="未配置")
        oss_key_secret = parser.get("OSS", "AccessKeySecret", fallback="未配置")

        def create_readonly_input(text, is_password=False):
            inp = QLineEdit(text)
            inp.setObjectName("configInput")
            inp.setReadOnly(True)
            inp.setMinimumHeight(32)
            if is_password:
                inp.setEchoMode(QLineEdit.Password)
            # 稍微改变背景色以示只读
            inp.setStyleSheet("QLineEdit { background-color: rgba(30, 41, 59, 0.5); color: #94a3b8; border: 1px solid #334155; }")
            return inp

        self.oss_endpoint_input = create_readonly_input(oss_endpoint)
        oss_form_layout.addRow("Endpoint:", self.oss_endpoint_input)

        self.oss_bucket_input = create_readonly_input(oss_bucket)
        oss_form_layout.addRow("Bucket:", self.oss_bucket_input)

        self.oss_key_input = create_readonly_input(oss_key_id)
        oss_form_layout.addRow("AccessKeyId:", self.oss_key_input)

        self.oss_secret_input = create_readonly_input(oss_key_secret, is_password=True)
        oss_form_layout.addRow("AccessKeySecret:", self.oss_secret_input)
        
        # 调整标签样式
        for i in range(oss_form_layout.rowCount()):
            item = oss_form_layout.itemAt(i, QFormLayout.LabelRole)
            if item and item.widget():
                item.widget().setStyleSheet("font-weight: bold; color: #cbd5e1; font-size: 13px;")

        info_layout.addWidget(oss_group)

        log_group = QGroupBox("\u8fd0\u884c\u65e5\u5fd7")
        log_group.setObjectName("logGroup")
        log_layout = QVBoxLayout(log_group)
        log_layout.setContentsMargins(12, 12, 12, 12)
        log_layout.setSpacing(8)

        self.runtime_log_view = QTextEdit()
        self.runtime_log_view.setObjectName("runtimeLogView")
        self.runtime_log_view.setReadOnly(True)
        self.runtime_log_view.setLineWrapMode(QTextEdit.NoWrap)
        self.runtime_log_view.setMinimumHeight(220)
        log_layout.addWidget(self.runtime_log_view, 1)

        log_btn_layout = QHBoxLayout()
        log_btn_layout.addStretch()
        self.copy_log_btn = QPushButton("\u4e00\u952e\u590d\u5236\u5168\u90e8\u65e5\u5fd7")
        self.copy_log_btn.setObjectName("copyLogBtn")
        self.copy_log_btn.setMinimumHeight(34)
        self.copy_log_btn.clicked.connect(self._copy_runtime_logs)
        log_btn_layout.addWidget(self.copy_log_btn)
        log_layout.addLayout(log_btn_layout)

        info_layout.addWidget(log_group, 1)

        info_layout.addStretch()

        self.page_stack.addWidget(tool_page)
        self.page_stack.addWidget(info_page)

        self.apply_styles()
        self.switch_page(0)

    def apply_styles(self):
        """Load dark theme from QSS file."""
        qss_path = Path(__file__).parent / "styles" / "dark_theme.qss"
        try:
            with open(qss_path, "r", encoding="utf-8") as f:
                qss = f.read()

            app = QApplication.instance()
            if app:
                app.setStyleSheet(qss)
            else:
                self.setStyleSheet(qss)
        except Exception as e:
            logger.warning(f"Failed to load stylesheet: {qss_path} ({e})")

    def switch_page(self, index):
        """Switch content page from left navigation."""
        self.page_stack.setCurrentIndex(index)
        self.nav_tool_btn.setProperty("active", index == 0)
        self.nav_info_btn.setProperty("active", index == 1)
        for btn in (self.nav_tool_btn, self.nav_info_btn):
            btn.style().unpolish(btn)
            btn.style().polish(btn)
            btn.update()

    def _read_runtime_config(self):
        """Read config.ini for info page display only."""
        config_path = Path(__file__).parent / "config.ini"
        parser = configparser.ConfigParser()
        if config_path.exists():
            parser.read(config_path, encoding="utf-8")
        return config_path, parser

    def _init_runtime_log_capture(self):
        """Attach a logging handler to stream logs into config page."""
        self._log_bridge = LogSignalBridge()
        self._log_bridge.message.connect(self._append_runtime_log)

        self._gui_log_handler = GuiLogHandler(self._log_bridge)
        self._gui_log_handler.setLevel(logging.DEBUG)
        self._gui_log_handler.setFormatter(
            logging.Formatter(
                "%(asctime)s | %(levelname)-7s | %(name)s | %(message)s",
                "%Y-%m-%d %H:%M:%S",
            )
        )

        root_logger = logging.getLogger('')
        root_logger.addHandler(self._gui_log_handler)

    def _load_existing_log_file(self):
        """Load existing process.log so users can inspect previous run details."""
        if not hasattr(self, "runtime_log_view"):
            return
        log_path = Path(__file__).parent / "process.log"
        if not log_path.exists():
            return

        try:
            with open(log_path, "r", encoding="utf-8", errors="replace") as f:
                content = f.read().strip()
            if content:
                self.runtime_log_view.setPlainText(content)
                self.runtime_log_view.moveCursor(QTextCursor.End)
        except Exception as e:
            self.runtime_log_view.append(f"[log-load-error] {e}")

    def _load_saved_task_file(self):
        """从 config.ini 加载上次保存的任务文件路径"""
        _, parser = self._read_runtime_config()
        saved_path = parser.get("Paths", "InputTaskFile", fallback="")
        if saved_path and os.path.isfile(saved_path):
            self.task_file = saved_path
            self.file_label.setText(saved_path)
            self.stage1_btn.setEnabled(True)
            self.auto_btn.setEnabled(True)
            self.manual_stage2_btn.setEnabled(True)

    def _append_runtime_log(self, message: str):
        """Append one log line into runtime log panel."""
        if not hasattr(self, "runtime_log_view"):
            return
        if message is None:
            return

        self.runtime_log_view.append(str(message).rstrip())

        # Keep recent logs only to avoid unlimited memory growth.
        content = self.runtime_log_view.toPlainText().splitlines()
        if len(content) > self._runtime_log_max_lines:
            content = content[-self._runtime_log_max_lines:]
            self.runtime_log_view.setPlainText("\n".join(content))

        self.runtime_log_view.moveCursor(QTextCursor.End)

    def _copy_runtime_logs(self):
        """Copy all runtime logs with one click."""
        if not hasattr(self, "runtime_log_view"):
            return

        text = self.runtime_log_view.toPlainText().strip()
        if not text:
            QMessageBox.information(self, "\u63d0\u793a", "\u5f53\u524d\u6ca1\u6709\u53ef\u590d\u5236\u7684\u65e5\u5fd7\u3002")
            return

        QApplication.clipboard().setText(text)
        QMessageBox.information(self, "\u590d\u5236\u6210\u529f", "\u65e5\u5fd7\u5df2\u590d\u5236\u5230\u526a\u8d34\u677f\u3002")

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

    def _save_task_file_path(self):
        """保存任务文件路径到 config.ini"""
        if not self.task_file:
            QMessageBox.warning(self, "警告", "请先选择任务文件")
            return
        config_path = Path(__file__).parent / "config.ini"
        parser = configparser.ConfigParser()
        if config_path.exists():
            parser.read(config_path, encoding="utf-8")
        if not parser.has_section("Paths"):
            parser.add_section("Paths")
        parser.set("Paths", "InputTaskFile", self.task_file)
        with open(config_path, "w", encoding="utf-8") as f:
            parser.write(f)
        QMessageBox.information(self, "保存成功", f"任务文件路径已保存: {self.task_file}")
            
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
        """??????"""
        logger.info(f"Start worker: mode={mode}, manual_dir={manual_dir}")
        self._stage1_workflow_name = self.workflow_combo.currentText()
        self.set_buttons_enabled(False)
        self.complete_frame.setVisible(False)
        self.progress_bar.setValue(0)

        # ????????????
        self.running_indicator.setVisible(True)
        self.stop_btn.setVisible(True)
        self.indicator_timer.start(40)  # 40ms ≈ 25fps 平滑呼吸动画
        self._pulse_step = 0

        # 高亮当前运行的按钮
        mode_btn_map = {
            'stage1': self.stage1_btn,
            'stage2': self.stage2_btn,
            'full_auto': self.auto_btn,
            'manual_stage2': self.manual_stage2_btn,
        }
        self._set_running_btn(mode_btn_map.get(mode))

        # ???stage2???????????
        if mode == 'stage2' and self.worker and self.worker.stage1_results:
            old_results = self.worker.stage1_results
            old_output_dir = self.worker.stage1_output_dir
            self.worker = WorkerThread(
                mode,
                self.task_file,
                comfyui_url=self.get_comfyui_url(),
                source_path=self.get_source_path(),
                stage1_output_dir=self.get_stage1_output_dir(),
                workflow_path=self.get_selected_workflow_path(),
            )
            self.worker.stage1_results = old_results
            self.worker.stage1_output_dir = old_output_dir
        else:
            self.worker = WorkerThread(
                mode,
                self.task_file,
                manual_dir,
                comfyui_url=self.get_comfyui_url(),
                source_path=self.get_source_path(),
                stage1_output_dir=self.get_stage1_output_dir(),
                workflow_path=self.get_selected_workflow_path(),
            )

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
        """????"""
        self.progress_bar.setMaximum(total)
        self.progress_bar.setValue(current)
        self.status_label.setText(message)

        marker = (current, total, message)
        if marker != self._last_progress_marker:
            self._last_progress_marker = marker
            logger.info(f"Progress: {current}/{total} | {message}")
    def append_log(self, message):
        """Worker signal hook; global logger handler already captures details."""
        _ = message
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
            self.gallery_btn.setVisible(True)
        elif stage_name in ("stage2", "manual_stage2"):
            self.complete_label.setText("✅ 全部完成！图片已处理并上传到阿里云 OSS。")
            self.output_path_label.setText(f"输出目录: {os.path.abspath(output_dir)}")
            if self.report_file:
                self.report_label.setText(f"报告文件: {self.report_file}")
            self.complete_frame.setVisible(True)
            # 阶段2显示报告按钮
            self.open_report_btn.setVisible(True)
            self.open_report_folder_btn.setVisible(True)
            self.gallery_btn.setVisible(False)
    
    def on_report_saved(self, report_path):
        """报告保存完成"""
        self.report_file = report_path
        self.report_label.setText(f"报告文件: {report_path}")
            
    def on_error(self, error_message):
        """????"""
        logger.error(f"Worker error: {error_message}")
        QMessageBox.critical(self, "错误", error_message)
        self.status_label.setText("错误")
    def on_worker_finished(self):
        """??????"""
        logger.info("Worker finished")
        # ???????
        self.running_indicator.setVisible(False)
        self.stop_btn.setVisible(False)
        self.indicator_timer.stop()
        self._set_running_btn(None)
        # 先清除运行按钮状态，再启用按钮（避免 _set_running_btn 把按钮又禁用）
        self.set_buttons_enabled(True)
        self.status_label.setText("\u5b8c\u6210")
    def animate_indicator(self):
        """平滑正弦波呼吸灯动画"""
        self._pulse_step += 1
        # 正弦波: 周期约2.5秒 (2.5s / 0.04s = ~63 steps per half cycle)
        t = math.sin(self._pulse_step * 0.05) * 0.5 + 0.5  # 0.0 ~ 1.0

        # 运行指示器小圆点颜色
        indicator_r = int(100 + 55 * t)
        indicator_g = int(160 + 60 * t)
        indicator_b = int(220 + 35 * t)
        self.running_indicator.setStyleSheet(
            f"font-size: 18px; color: rgb({indicator_r},{indicator_g},{indicator_b});"
        )

        # 呼吸脉冲: 平滑渐变按钮背景和边框
        if self._running_btn:
            obj_name = self._running_btn.objectName()
            theme = self._btn_color_themes.get(obj_name)
            if theme:
                dim_bg, bright_bg, dim_bd, bright_bd = theme
                bg = tuple(int(d + (b - d) * t) for d, b in zip(dim_bg, bright_bg))
                bd = tuple(int(d + (b - d) * t) for d, b in zip(dim_bd, bright_bd))
                self._running_btn.setStyleSheet(
                    f"color: #ffffff; font-weight: 700;"
                    f"background: rgb({bg[0]},{bg[1]},{bg[2]});"
                    f"border: 2px solid rgb({bd[0]},{bd[1]},{bd[2]});"
                    f"border-radius: 8px; padding: 8px 12px; min-height: 30px;"
                )

    def _set_running_btn(self, btn):
        """设置/清除当前运行中的按钮高亮"""
        if self._running_btn:
            # 清除内联样式，恢复QSS主题样式
            self._running_btn.setStyleSheet("")
            self._running_btn.setProperty("running", False)
            self._running_btn.style().unpolish(self._running_btn)
            self._running_btn.style().polish(self._running_btn)
            self._running_btn.setEnabled(False)
        self._running_btn = btn
        self._pulse_step = 0
        if btn:
            btn.setProperty("running", True)
            btn.setEnabled(True)
            btn.style().unpolish(btn)
            btn.style().polish(btn)
    
    def _open_gallery(self):
        """打开 Stage1 图库预览（如果已有窗口则激活显示）"""
        # 如果已有图库窗口且未关闭，直接激活显示
        if hasattr(self, '_gallery_dlg') and self._gallery_dlg is not None:
            if self._gallery_dlg.isVisible():
                self._gallery_dlg.showNormal()
                self._gallery_dlg.activateWindow()
                self._gallery_dlg.raise_()
                return

        if not self.worker or not self.worker.stage1_results:
            QMessageBox.warning(self, "警告", "没有阶段1的处理结果！")
            return

        image_paths = []
        source_map = {}  # output_path -> source_path
        for src_path, info in self.worker.stage1_results.items():
            out = info.get("output", "")
            if out and os.path.isfile(out):
                image_paths.append(out)
                source_map[out] = src_path

        if not image_paths:
            QMessageBox.warning(self, "警告", "未找到有效的输出图片！")
            return

        self._gallery_dlg = ImageGalleryDialog(
            image_paths=image_paths,
            source_map=source_map,
            comfyui_url=self.get_comfyui_url(),
            current_workflow_name=self._stage1_workflow_name,
            workflows_dir=str(self._get_workflows_dir()),
            parent=self,
        )
        self._gallery_dlg.show()

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
        self._set_running_btn(None)
        self.status_label.setText("已停止")
        self.progress_bar.setValue(0)  # 重置进度条
        self.set_buttons_enabled(True)
        
        # 获取当前输出目录（优先从worker获取实际路径）
        output_dir = self.current_output_dir
        if not output_dir and self.worker and hasattr(self.worker, 'stage1_output_dir'):
            output_dir = self.worker.stage1_output_dir
        if not output_dir:
            output_dir = self.get_stage1_output_dir() or self.get_source_path()
        if not output_dir:
            self.set_buttons_enabled(True)
            return
        self.current_output_dir = output_dir
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
        
        msg.exec()
        
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

        # 停止后显示complete_frame，方便用户打开输出目录
        if self.current_output_dir and os.path.exists(os.path.abspath(self.current_output_dir)):
            self.complete_label.setText("⚠️ 任务已停止")
            self.output_path_label.setText(f"输出目录: {output_path}")
            self.report_label.setText("")
            self.complete_frame.setVisible(True)
            self.open_report_btn.setVisible(False)
            self.open_report_folder_btn.setVisible(False)
            self.gallery_btn.setVisible(False)

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
            
            msg.exec()
            
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

    # ---- ComfyUI 全局配置 ----

    def get_comfyui_url(self) -> str:
        """??????? ComfyUI ???"""
        return self.comfyui_url_input.text().strip()

    def _normalize_comfyui_url(self, url: str) -> str:
        """Normalize user input URL for test/save."""
        clean = (url or "").strip()
        if not clean:
            return ""
        if "://" not in clean:
            clean = f"http://{clean}"
        return clean.rstrip("/")

    def _set_comfyui_status(self, state: str, message: str):
        """Update ComfyUI config status text and style state."""
        if not hasattr(self, "comfyui_status_label"):
            return
        self.comfyui_status_label.setProperty("state", state)
        self.comfyui_status_label.setText(message)
        self.comfyui_status_label.style().unpolish(self.comfyui_status_label)
        self.comfyui_status_label.style().polish(self.comfyui_status_label)
        self.comfyui_status_label.update()

    def _start_comfyui_glow(self):
        """启动 ComfyUI 输入框绿色流动边框动画"""
        if self._comfyui_glow_timer is None:
            self._comfyui_glow_timer = QTimer(self)
            self._comfyui_glow_timer.timeout.connect(self._animate_comfyui_glow)
        self._comfyui_glow_step = 0
        self._comfyui_glow_timer.start(40)

    def _stop_comfyui_glow(self):
        """停止绿色流动边框动画，恢复默认样式"""
        if self._comfyui_glow_timer:
            self._comfyui_glow_timer.stop()
        if hasattr(self, "comfyui_url_input"):
            self.comfyui_url_input.setStyleSheet("")

    def _animate_comfyui_glow(self):
        """绿色流动电流动画 — 双正弦波叠加产生流动感"""
        self._comfyui_glow_step += 1
        s = self._comfyui_glow_step
        # 主波: 慢速呼吸 (周期~3s)
        t1 = math.sin(s * 0.04) * 0.5 + 0.5
        # 副波: 快速闪烁 (周期~0.8s), 幅度较小
        t2 = math.sin(s * 0.16) * 0.15
        t = max(0.0, min(1.0, t1 + t2))

        # 边框颜色: 从暗绿到亮绿
        br = int(20 + 14 * t)
        bg = int(120 + 137 * t)  # 120 → 257 clamped to 255
        bb = int(60 + 68 * t)
        bg = min(bg, 255)
        # 背景微微泛绿光
        bkg_a = int(8 + 12 * t)

        self.comfyui_url_input.setStyleSheet(
            f"QLineEdit {{ "
            f"background: rgba(10, {30 + int(20*t)}, 15, 240); "
            f"color: #a7f3d0; "
            f"border: 2px solid rgb({br},{bg},{bb}); "
            f"border-radius: 8px; padding: 8px 10px; "
            f"font-size: 14px; font-weight: 600; }}"
        )

    def _on_comfyui_url_changed(self, _text: str):
        """URL changed: require re-test before save."""
        self._comfyui_test_ok = False
        self._comfyui_tested_url = ""
        self._stop_comfyui_glow()

        if not hasattr(self, "save_comfyui_btn") or not hasattr(self, "test_comfyui_btn"):
            return

        current_url = self._normalize_comfyui_url(self.get_comfyui_url())
        self.save_comfyui_btn.setEnabled(False)
        self.test_comfyui_btn.setEnabled(bool(current_url))
        self._set_comfyui_status(
            "pending",
            "请先点击“测试连接”，连接成功后再保存为全局配置。"
        )

    def _test_comfyui_connection(self):
        """Test current ComfyUI URL without saving."""
        raw_url = self.get_comfyui_url()
        url = self._normalize_comfyui_url(raw_url)
        if not url:
            QMessageBox.warning(self, "警告", "请输入 ComfyUI 地址")
            return

        if url != raw_url:
            self.comfyui_url_input.setText(url)

        if self._comfyui_test_worker and self._comfyui_test_worker.isRunning():
            return

        self._comfyui_test_ok = False
        self._comfyui_tested_url = ""
        self.save_comfyui_btn.setEnabled(False)
        self.test_comfyui_btn.setEnabled(False)
        self.test_comfyui_btn.setText("测试中...")
        self._set_comfyui_status("testing", f"正在测试连接: {url}")

        self._comfyui_test_worker = ComfyUIConnectionTestWorker(url, self)
        self._comfyui_test_worker.check_finished.connect(self._on_comfyui_test_finished)
        self._comfyui_test_worker.start()

    def _on_comfyui_test_finished(self, ok: bool, tested_url: str, message: str):
        """Handle async test result."""
        current_url = self._normalize_comfyui_url(self.get_comfyui_url())
        self.test_comfyui_btn.setEnabled(True)
        self.test_comfyui_btn.setText("测试连接")
        self._comfyui_test_worker = None

        # Ignore stale result when user changed URL during test
        if tested_url != current_url:
            self.save_comfyui_btn.setEnabled(False)
            self._set_comfyui_status("pending", "地址已修改，请重新测试连接。")
            return

        if ok:
            self._comfyui_test_ok = True
            self._comfyui_tested_url = tested_url
            self.save_comfyui_btn.setEnabled(True)
            self._set_comfyui_status("ok", f"{message}，可点击“保存”写入全局配置。")
            self._start_comfyui_glow()
        else:
            self._comfyui_test_ok = False
            self._comfyui_tested_url = ""
            self.save_comfyui_btn.setEnabled(False)
            self._set_comfyui_status("error", message)
            self._stop_comfyui_glow()

    def _save_comfyui_url(self):
        """?? ComfyUI ??? config.ini?????????????"""
        url = self._normalize_comfyui_url(self.get_comfyui_url())
        if not url:
            QMessageBox.warning(self, "警告", "请输入 ComfyUI 地址")
            return

        if not self._comfyui_test_ok or self._comfyui_tested_url != url:
            QMessageBox.warning(
                self,
                "未通过测试",
                "请先点击“测试连接”，并在连接成功后再保存当前地址。"
            )
            return

        config_path = Path(__file__).parent / "config.ini"
        parser = configparser.ConfigParser()
        if config_path.exists():
            parser.read(config_path, encoding="utf-8")

        from urllib.parse import urlparse
        parsed = urlparse(url)
        host = parsed.hostname or "127.0.0.1"
        port = str(parsed.port or (443 if parsed.scheme == "https" else 8188))

        if not parser.has_section("ComfyUI"):
            parser.add_section("ComfyUI")
        parser.set("ComfyUI", "Host", host)
        parser.set("ComfyUI", "DefaultPort", port)

        with open(config_path, "w", encoding="utf-8") as f:
            parser.write(f)

        QMessageBox.information(self, "保存成功", f"ComfyUI 地址已保存: {host}:{port}")
        self._set_comfyui_status("ok", f"已保存全局配置: {host}:{port}")
        self.save_comfyui_btn.setEnabled(False)

    # ---- 图片源路径配置 ----

    def get_source_path(self) -> str:
        """返回当前配置的图片源路径"""
        return self.source_path_input.text().strip()

    def _browse_source_path(self):
        """打开文件夹选择对话框"""
        folder = QFileDialog.getExistingDirectory(self, "选择图片源文件夹", self.get_source_path() or "")
        if folder:
            self.source_path_input.setText(folder)

    def _save_source_path(self):
        """保存图片源路径到 config.ini"""
        path = self.get_source_path()
        if not path:
            QMessageBox.warning(self, "警告", "请输入或选择图片源路径")
            return
        if not os.path.isdir(path):
            QMessageBox.warning(self, "警告", f"路径不存在: {path}")
            return

        config_path = Path(__file__).parent / "config.ini"
        parser = configparser.ConfigParser()
        if config_path.exists():
            parser.read(config_path, encoding="utf-8")
        if not parser.has_section("Paths"):
            parser.add_section("Paths")
        parser.set("Paths", "SourcePath", path)
        with open(config_path, "w", encoding="utf-8") as f:
            parser.write(f)
        QMessageBox.information(self, "保存成功", f"图片源路径已保存: {path}")

    # ---- Stage1 Output Path Config ----

    def get_stage1_output_dir(self) -> str:
        """Return configured stage1 output directory."""
        if not hasattr(self, "stage1_output_input"):
            return ""
        return self.stage1_output_input.text().strip()

    def _browse_stage1_output_dir(self):
        """Open folder chooser for stage1 output directory."""
        default_dir = self.get_stage1_output_dir() or self.get_source_path() or ""
        folder = QFileDialog.getExistingDirectory(self, "Select stage1 output folder", default_dir)
        if folder:
            self.stage1_output_input.setText(folder)

    def _save_stage1_output_dir(self):
        """Save stage1 output directory into config.ini."""
        path = self.get_stage1_output_dir()
        if not path:
            QMessageBox.warning(self, "Warning", "Please input or choose stage1 output directory")
            return
        if os.path.isfile(path):
            QMessageBox.warning(self, "Warning", f"Path is a file, not a folder: {path}")
            return
        try:
            os.makedirs(path, exist_ok=True)
        except Exception as e:
            QMessageBox.warning(self, "Warning", f"Failed to create directory: {e}")
            return

        config_path = Path(__file__).parent / "config.ini"
        parser = configparser.ConfigParser()
        if config_path.exists():
            parser.read(config_path, encoding="utf-8")
        if not parser.has_section("Paths"):
            parser.add_section("Paths")
        parser.set("Paths", "Stage1OutputPath", path)
        with open(config_path, "w", encoding="utf-8") as f:
            parser.write(f)
        QMessageBox.information(self, "Saved", f"Stage1 output path saved: {path}")

    def _clear_stage1_output_dir(self):
        """清空阶段1输出目录下的所有内容（保留根文件夹）"""
        path = self.get_stage1_output_dir()
        if not path:
            QMessageBox.warning(self, "警告", "请先配置阶段1输出路径")
            return
        if not os.path.isdir(path):
            QMessageBox.warning(self, "警告", f"目录不存在: {path}")
            return

        # 统计内容
        items = os.listdir(path)
        if not items:
            QMessageBox.information(self, "提示", "目录已经是空的。")
            return

        reply = QMessageBox.warning(
            self, "确认清空",
            f"即将删除以下目录中的所有内容:\n\n"
            f"📁 {path}\n\n"
            f"共 {len(items)} 个文件/文件夹，此操作不可撤销！",
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No,
        )
        if reply != QMessageBox.Yes:
            return

        try:
            for item in items:
                item_path = os.path.join(path, item)
                if os.path.isdir(item_path):
                    shutil.rmtree(item_path)
                else:
                    os.remove(item_path)
            QMessageBox.information(self, "完成", f"已清空: {path}")
        except Exception as e:
            QMessageBox.warning(self, "清空失败", f"部分内容无法删除: {e}")

    # ---- Workflow Selection Config ----

    def _get_workflows_dir(self) -> Path:
        """Return the workflows/ directory path, creating it if needed."""
        wf_dir = Path(__file__).parent / "workflows"
        wf_dir.mkdir(exist_ok=True)
        return wf_dir

    def _refresh_workflow_combo(self):
        """Scan workflows/ directory and repopulate the combo box."""
        self.workflow_combo.blockSignals(True)
        current = self.workflow_combo.currentText()
        self.workflow_combo.clear()
        wf_dir = self._get_workflows_dir()
        names = sorted(
            p.stem for p in wf_dir.glob("*.json")
        )
        self.workflow_combo.addItems(names)
        # restore previous selection if still present
        idx = self.workflow_combo.findText(current)
        if idx >= 0:
            self.workflow_combo.setCurrentIndex(idx)
        self.workflow_combo.blockSignals(False)

    def get_selected_workflow_path(self) -> str:
        """Return full path of the currently selected workflow JSON."""
        name = self.workflow_combo.currentText()
        if not name:
            return ""
        return str(self._get_workflows_dir() / f"{name}.json")

    def _upload_workflow(self):
        """Let user pick a JSON file, name it, and copy into workflows/."""
        src, _ = QFileDialog.getOpenFileName(
            self, "选择工作流文件", "", "JSON文件 (*.json)"
        )
        if not src:
            return
        name, ok = QInputDialog.getText(
            self, "命名工作流", "请输入工作流名称:"
        )
        if not ok or not name.strip():
            return
        name = name.strip()
        dest = self._get_workflows_dir() / f"{name}.json"
        if dest.exists():
            ret = QMessageBox.question(
                self, "确认覆盖",
                f"工作流 \"{name}\" 已存在，是否覆盖？",
            )
            if ret != QMessageBox.Yes:
                return
        shutil.copy2(src, dest)
        self._refresh_workflow_combo()
        idx = self.workflow_combo.findText(name)
        if idx >= 0:
            self.workflow_combo.setCurrentIndex(idx)
        QMessageBox.information(self, "上传成功", f"工作流 \"{name}\" 已添加")

    def _delete_workflow(self):
        """Delete the currently selected workflow file."""
        name = self.workflow_combo.currentText()
        if not name:
            return
        ret = QMessageBox.question(
            self, "确认删除",
            f"确定要删除工作流 \"{name}\" 吗？",
        )
        if ret != QMessageBox.Yes:
            return
        path = self._get_workflows_dir() / f"{name}.json"
        if path.exists():
            path.unlink()
        self._refresh_workflow_combo()
        QMessageBox.information(self, "删除成功", f"工作流 \"{name}\" 已删除")

    def _save_workflow_selection(self):
        """Save the current workflow selection to config.ini."""
        name = self.workflow_combo.currentText()
        if not name:
            QMessageBox.warning(self, "警告", "请先选择一个工作流")
            return
        config_path = Path(__file__).parent / "config.ini"
        parser = configparser.ConfigParser()
        if config_path.exists():
            parser.read(config_path, encoding="utf-8")
        if not parser.has_section("ComfyUI"):
            parser.add_section("ComfyUI")
        parser.set("ComfyUI", "SelectedWorkflow", name)
        with open(config_path, "w", encoding="utf-8") as f:
            parser.write(f)
        QMessageBox.information(self, "保存成功", f"已选择工作流: {name}")

    def _check_for_updates(self, silent=True):
        """Check updates in background; show dialogs only when silent=False."""
        if self._update_checker and self._update_checker.isRunning():
            return

        self._update_check_silent = silent
        self.update_check_btn.setEnabled(False)
        self.update_check_btn.setText("检查中...")
        logger.info(f"Start update check: version={APP_VERSION}, repo={GITHUB_REPO}, silent={silent}")

        self._update_checker = UpdateCheckWorker(APP_VERSION, GITHUB_REPO)
        self._update_checker.update_available.connect(self._on_update_available)
        self._update_checker.no_update.connect(self._on_no_update)
        self._update_checker.check_failed.connect(self._on_update_check_failed)
        self._update_checker.start()

    def _on_update_available(self, release_info):
        """New version found."""
        self.update_check_btn.setEnabled(True)
        self.update_check_btn.setText("检查更新")
        logger.info(f"Update available: v{release_info.version}")
        dialog = UpdateDialog(release_info, APP_VERSION, parent=self)
        dialog.exec()

    def _on_no_update(self):
        """Already latest version."""
        self.update_check_btn.setEnabled(True)
        self.update_check_btn.setText("检查更新")
        logger.info("Already latest version")
        if not self._update_check_silent:
            QMessageBox.information(
                self,
                "检查更新",
                f"当前已是最新版本 v{APP_VERSION}"
            )

    def _on_update_check_failed(self, error):
        """Update check failed."""
        self.update_check_btn.setEnabled(True)
        self.update_check_btn.setText("检查更新")
        logger.error(f"Update check failed: {error}")

        if not self._update_check_silent:
            hint = (
                "可能原因：\n"
                "1. 仓库为私有仓库，客户端未授权访问；\n"
                "2. 尚未创建 GitHub Release；\n"
                "3. Release 未上传 .zip 更新包资产。"
            )
            QMessageBox.warning(self, "检查更新失败", f"{error}\n\n{hint}")

