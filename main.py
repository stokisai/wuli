import pandas as pd
import os
import configparser
import logging
import time
from datetime import datetime
from image_processor import ImageProcessor
from drive_uploader import DriveUploader
from comfyui_client import ComfyUIClient
from utils import setup_logging, ensure_dir, check_and_download_font

logger = setup_logging()

def load_processed_log(log_path="processed_log.txt"):
    if os.path.exists(log_path):
        with open(log_path, "r", encoding="utf-8") as f:
            return set(line.strip() for line in f)
    return set()

def mark_processed(file_name, log_path="processed_log.txt"):
    with open(log_path, "a", encoding="utf-8") as f:
        f.write(file_name + "\n")

def main():
    config = configparser.ConfigParser()
    config.read("config.ini", encoding='utf-8')
    
    input_task_file = config["Paths"]["InputTaskFile"]
    output_report_file = config["Paths"]["OutputReportFile"]
    
    if not os.path.exists(input_task_file):
        logger.error(f"Task file not found: {input_task_file}")
        return

    # Initialize Modules
    processor = ImageProcessor()
    uploader = DriveUploader()
    
    # Try auth
    if not uploader.authenticate():
        logger.warning("Google Drive authentication failed or skipped. Uploads will fail.")
    
    try:
        df_tasks = pd.read_excel(input_task_file)
    except Exception as e:
        logger.error(f"Failed to read task file: {e}")
        return

    # 1. GROUP BY 'Folder Name' + 'Source Path' to handle multi-row logic
    # We maintain order of appearance in Excel
    grouped = df_tasks.groupby(['Folder Name', 'Source Path'], sort=False)
    
    report_aggregator = {}
    all_tasks = []  # 在循环外初始化任务列表
    
    for (folder_name, source_path), group_df in grouped:
        logger.info(f"Processing Task Group: {folder_name} (Rows: {len(group_df)})")
        
        if folder_name not in report_aggregator:
            report_aggregator[folder_name] = {} # Dict to hold Image 1, Image 2...

        # A. Create Drive Folder (once per group)
        drive_folder_id = uploader.create_folder(folder_name)
        if not drive_folder_id:
             logger.warning(f"Drive folder creation failed for {folder_name}. Continuing with local processing only.")
        
        # B. Get Images - 递归遍历所有子文件夹，按文件夹分组
        if not os.path.exists(source_path):
            logger.error(f"Source path not found: {source_path}")
            continue
        
        def collect_images_by_folder(root_path):
            """递归收集所有子文件夹及其图片
            
            返回: [(folder_rel_path, [image_full_paths]), ...]
            按深度优先顺序，每个子文件夹单独作为一个处理单元
            """
            folder_images = []  # list of (folder_rel_path, [images])
            valid_exts = ('.jpg', '.jpeg', '.png')
            
            def process_folder(folder_path):
                logger.info(f"Scanning folder: {folder_path}")
                images_in_folder = []
                subdirs = []
                
                try:
                    items = sorted(os.listdir(folder_path))
                except Exception as e:
                    logger.error(f"Error listing folder {folder_path}: {e}")
                    return
                
                for item in items:
                    full_path = os.path.join(folder_path, item)
                    
                    if os.path.isfile(full_path):
                        lower_item = item.lower()
                        if lower_item.endswith(valid_exts):
                            # 过滤无效文件
                            if "副本" in item: continue
                            if "copy" in lower_item: continue
                            if "._" in item: continue
                            if item.startswith("$"): continue
                            images_in_folder.append(full_path)
                    elif os.path.isdir(full_path):
                        subdirs.append(full_path)
                
                # 如果当前文件夹有图片，添加为一个处理单元
                if images_in_folder:
                    # 计算相对于root的路径作为文件夹名
                    rel_folder = os.path.relpath(folder_path, root_path)
                    if rel_folder == ".":
                        rel_folder = os.path.basename(root_path)  # 根目录用其名称
                    folder_images.append((rel_folder, sorted(images_in_folder)))
                    logger.info(f"  → Folder '{rel_folder}' has {len(images_in_folder)} images")
                
                # 递归处理子文件夹（深度优先）
                for subdir in subdirs:
                    process_folder(subdir)
            
            process_folder(root_path)
            return folder_images
        
        try:
            folder_groups = collect_images_by_folder(source_path)
            
            if not folder_groups:
                logger.warning(f"No valid images found in {source_path} (including subfolders)")
                continue
            
            total_images = sum(len(imgs) for _, imgs in folder_groups)
            logger.info(f"Total: {len(folder_groups)} folders, {total_images} images from {source_path}")
                
        except Exception as e:
             logger.error(f"Error collecting images in {source_path}: {e}")
             continue
        
        # C. 获取Excel中的任务行配置（用于文案等参数）
        task_rows = [row for _, row in group_df.iterrows()]
        global_row_idx = 0  # 全局图片计数，用于对应Excel行
        
        # D. 按文件夹处理
        for folder_rel_path, images_in_folder in folder_groups:
            # 使用子文件夹的相对路径作为文件夹名（不加Excel前缀）
            drive_subfolder_name = folder_rel_path.replace(os.sep, "_")
            drive_subfolder_id = uploader.create_folder(drive_subfolder_name)
            if not drive_subfolder_id:
                logger.warning(f"Drive folder creation failed for {drive_subfolder_name}")
            
            # 初始化该文件夹的报告记录
            if drive_subfolder_name not in report_aggregator:
                report_aggregator[drive_subfolder_name] = {}
            
            logger.info(f"Processing folder: {folder_rel_path} ({len(images_in_folder)} images)")
            
            for local_idx, full_img_path in enumerate(images_in_folder):
                img_name = os.path.basename(full_img_path)
                
                # 获取对应的Excel行数据（如果还有的话）
                if global_row_idx < len(task_rows):
                    row_data = task_rows[global_row_idx]
                else:
                    # 没有更多Excel行，使用最后一行的配置
                    row_data = task_rows[-1] if task_rows else {}
                    logger.warning(f"No more Excel rows, using last row config for {img_name}")
                
                # Extract Row Data
                jp_top = str(row_data.get("Top Text JP", "")) if pd.notna(row_data.get("Top Text JP")) else ""
                jp_bottom = str(row_data.get("Bottom Text JP", "")) if pd.notna(row_data.get("Bottom Text JP")) else ""
                
                # Font Sizes
                top_size = 0
                if "Top Font Size" in row_data and pd.notna(row_data["Top Font Size"]):
                    try: top_size = int(float(row_data["Top Font Size"]))
                    except: pass
                    
                bottom_size = 0
                if "Bottom Font Size" in row_data and pd.notna(row_data["Bottom Font Size"]):
                    try: bottom_size = int(float(row_data["Bottom Font Size"]))
                    except: pass
                
                # Font Name
                font_name = None
                if "fonts" in row_data and pd.notna(row_data["fonts"]):
                    font_name = str(row_data["fonts"]).strip()
                    if font_name.lower() == 'nan':
                        font_name = None

                if jp_top.lower() == "nan": jp_top = ""
                if jp_bottom.lower() == "nan": jp_bottom = ""

                # ========== 读取ComfyUI配置 ==========
                comfyui_url = None
                if "comfyui" in row_data and pd.notna(row_data["comfyui"]):
                    comfyui_url = str(row_data["comfyui"]).strip()
                    if comfyui_url.lower() == 'nan':
                        comfyui_url = None
                
                # 读取1stage输出目录
                processed_1stage_dir = None
                if "Processed image 1stage" in row_data and pd.notna(row_data["Processed image 1stage"]):
                    processed_1stage_dir = str(row_data["Processed image 1stage"]).strip()
                    if processed_1stage_dir.lower() == 'nan':
                        processed_1stage_dir = None

                # 收集任务信息用于两阶段处理
                task_info = {
                    'source_path': full_img_path,
                    'img_name': img_name,
                    'folder_rel_path': folder_rel_path,
                    'comfyui_url': comfyui_url,
                    'stage1_dir': processed_1stage_dir,
                    'jp_top': jp_top,
                    'jp_bottom': jp_bottom,
                    'top_size': top_size,
                    'bottom_size': bottom_size,
                    'font_name': font_name,
                    'drive_subfolder_id': drive_subfolder_id,
                    'drive_subfolder_name': drive_subfolder_name,
                    'local_idx': local_idx,
                }
                
                # 添加到任务列表
                all_tasks.append(task_info)
                
                global_row_idx += 1
    
    # ========== 阶段1: ComfyUI 图生图处理 ==========
    logger.info("="*60)
    logger.info("阶段1: ComfyUI 图生图处理")
    logger.info("="*60)
    
    # 从第一个任务获取全局ComfyUI配置
    global_comfyui_url = None
    global_stage1_dir = None
    for task in all_tasks:
        if task['comfyui_url'] and task['stage1_dir']:
            global_comfyui_url = task['comfyui_url']
            global_stage1_dir = task['stage1_dir']
            break
    
    if not global_comfyui_url or not global_stage1_dir:
        logger.error("未找到ComfyUI配置！请确保Excel第一行配置了 comfyui 和 Processed image 1stage 列")
        return
    
    logger.info(f"ComfyUI服务器: {global_comfyui_url}")
    logger.info(f"Stage1输出目录: {global_stage1_dir}")
    logger.info(f"需要处理的图片总数: {len(all_tasks)}")
    
    stage1_results = {}  # {source_path: stage1_output_path}
    stage1_failures = []  # 记录失败的任务
    
    # 初始化ComfyUI客户端
    comfyui_client = ComfyUIClient.from_url(global_comfyui_url)
    logger.info(f"已连接到ComfyUI服务器: {comfyui_client.base_url}")
    
    # 处理所有图片
    for idx, task in enumerate(all_tasks, 1):
        # 构建stage1输出路径 - 保持源文件夹结构
        stage1_subfolder = os.path.join(global_stage1_dir, task['folder_rel_path'])
        ensure_dir(stage1_subfolder)
        stage1_output = os.path.join(stage1_subfolder, task['img_name'])
        
        logger.info(f"[Stage1] ({idx}/{len(all_tasks)}) 处理: {task['folder_rel_path']}/{task['img_name']}")
        
        try:
            if comfyui_client.process_image(task['source_path'], stage1_output):
                logger.info(f"  ✓ 成功: {stage1_output}")
                stage1_results[task['source_path']] = stage1_output
            else:
                logger.error(f"  ✗ 失败: {task['img_name']}")
                stage1_failures.append(f"{task['folder_rel_path']}/{task['img_name']}")
        except Exception as e:
            logger.error(f"  ✗ 异常: {e}")
            stage1_failures.append(f"{task['folder_rel_path']}/{task['img_name']}")
    
    logger.info("="*60)
    logger.info(f"阶段1完成: {len(stage1_results)}/{len(all_tasks)} 成功")
    logger.info("="*60)
    
    # 如果有ComfyUI任务失败，停止处理
    if stage1_failures:
        logger.error("ComfyUI处理失败，程序终止！")
        logger.error(f"失败的图片 ({len(stage1_failures)}张):")
        for f in stage1_failures[:10]:  # 只显示前10个
            logger.error(f"  - {f}")
        if len(stage1_failures) > 10:
            logger.error(f"  ... 还有 {len(stage1_failures) - 10} 张")
        return
    
    # ========== 阶段1验证 ==========
    logger.info("="*60)
    logger.info("阶段1验证: 检查输出是否完整")
    logger.info("="*60)
    
    # 统计源文件数量和输出文件数量
    source_count = len(all_tasks)
    output_count = len(stage1_results)
    
    # 检查输出目录中的实际文件数量
    actual_output_files = 0
    for root, dirs, files in os.walk(global_stage1_dir):
        for f in files:
            if f.lower().endswith(('.jpg', '.jpeg', '.png')):
                actual_output_files += 1
    
    logger.info(f"源图片数量: {source_count}")
    logger.info(f"成功处理数量: {output_count}")
    logger.info(f"输出目录实际文件数: {actual_output_files}")
    
    if output_count != source_count:
        logger.error(f"警告: 输出数量不匹配！ 源图: {source_count}, 输出: {output_count}")
        return
    
    logger.info("✓ 验证通过: 所有图片已成功处理")
    
    # ========== 等待用户确认 ==========
    logger.info("="*60)
    logger.info("阶段1已完成！请检查输出目录确认图片质量。")
    logger.info(f"输出目录: {global_stage1_dir}")
    logger.info("="*60)
    
    print("\n" + "="*60)
    print("请确认是否继续执行阶段2（添加文字标签）？")
    print("输入 'y' 或 'yes' 继续，其他任何输入将取消操作")
    print("="*60)
    
    user_input = input(">>> ").strip().lower()
    if user_input not in ('y', 'yes'):
        logger.info("用户取消操作，程序终止。")
        print("已取消，程序退出。")
        return
    
    logger.info("用户确认继续，开始阶段2...")
    
    # ========== 阶段2: 添加文字标签 ==========
    logger.info("="*60)
    logger.info("阶段2: 添加文字标签")
    logger.info("="*60)
    
    for idx, task in enumerate(all_tasks, 1):
        # 必须使用ComfyUI处理后的图片
        if task['source_path'] not in stage1_results:
            logger.error(f"[Stage2] 错误: 未找到处理后的图片: {task['folder_rel_path']}/{task['img_name']}")
            continue
        
        current_img_path = stage1_results[task['source_path']]
        logger.info(f"[Stage2] ({idx}/{len(all_tasks)}) 添加文案: {task['folder_rel_path']}/{task['img_name']}")
        
        # 输出文件名
        output_filename = f"{task['folder_rel_path']}_{task['img_name']}".replace(os.sep, "_")
        
        temp_output_dir = "temp_processed"
        ensure_dir(temp_output_dir)
        processed_path = os.path.join(temp_output_dir, output_filename)
        
        logger.info(f"[Stage2] 添加文案: {task['folder_rel_path']}/{task['img_name']}")
        
        # 处理图片
        result_link = "Processing Failed"
        try:
            success = processor.process_image(
                current_img_path, processed_path, 
                task['jp_top'], task['jp_bottom'],
                top_size=task['top_size'], 
                bottom_size=task['bottom_size'], 
                font_name=task['font_name']
            )
            
            if success:
                result_link = "Upload Skipped/Failed"
                if task['drive_subfolder_id']:
                    file_obj = uploader.upload_file(processed_path, task['drive_subfolder_id'])
                    if file_obj:
                        result_link = uploader.get_direct_link(file_obj['id'])
            else:
                logger.error(f"图片处理失败: {task['img_name']}")
        except Exception as e:
            logger.error(f"处理异常 {task['img_name']}: {e}")
        
        # 记录到报告
        if task['drive_subfolder_name'] not in report_aggregator:
            report_aggregator[task['drive_subfolder_name']] = {}
        report_aggregator[task['drive_subfolder_name']][f"Image {task['local_idx']+1}"] = result_link
        
    # Save Report
    final_rows = []
    # Convert aggregator dict to List of Dicts
    for f_name, links_dict in report_aggregator.items():
        row_dict = {"Folder Name": f_name}
        row_dict.update(links_dict) # Merges "Image 1": "http..."
        final_rows.append(row_dict)

    if final_rows:
        new_df = pd.DataFrame(final_rows)
        # Combine
        existing_report_df = pd.DataFrame()
        if os.path.exists(output_report_file):
            try:
                existing_report_df = pd.read_excel(output_report_file)
            except: pass

        if not existing_report_df.empty:
            # We need to handle column mismatch if new report has fewer/more image columns
            final_df = pd.concat([existing_report_df, new_df], ignore_index=True)
            try:
               final_df.drop_duplicates(subset=['Folder Name'], keep='last', inplace=True)
            except: pass
        else:
            final_df = new_df
            
        final_df.to_excel(output_report_file, index=False)
        logger.info(f"Report saved to {output_report_file}")
    else:
         logger.info("No report to save.")

if __name__ == "__main__":
    main()
