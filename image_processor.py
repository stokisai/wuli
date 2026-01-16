"""
Image Processor - Simplified Version
只保留日文标签添加功能，移除所有文字检测和inpainting代码
"""
import os
import cv2
import re
import numpy as np
import logging
from PIL import Image, ImageDraw, ImageFont
import configparser

logger = logging.getLogger(__name__)


class ImageProcessor:
    def __init__(self, config_path="config.ini"):
        self.debug_mode = False
        
        self.config = configparser.ConfigParser()
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                self.config.read_file(f)
        except FileNotFoundError:
            logger.warning(f"Config file not found at {config_path}")
        except Exception as e:
            logger.error(f"Error reading config: {e}")
        
        self.font_path = self.resolve_font()
        if self.font_path:
            logger.info(f"Font loaded: {self.font_path}")
        else:
            logger.warning("No font found, will use system default")
    
    def resolve_font(self):
        """Find a valid font path"""
        cfg_font = self.config.get("Paths", "FontPath", fallback="")
        if cfg_font and os.path.exists(cfg_font):
            return cfg_font
            
        candidates = [
            r"C:\Windows\Fonts\Alibaba-PuHuiTi-Medium.ttf",
            r"C:\Windows\Fonts\Alibaba-PuHuiTi-Regular.ttf",
            r"C:\Windows\Fonts\msgothic.ttc",
            r"C:\Windows\Fonts\meiryo.ttc",
            r"C:\Windows\Fonts\msyh.ttc",
            r"C:\Windows\Fonts\simsun.ttc",
        ]
        
        for f in candidates:
            if os.path.exists(f):
                logger.info(f"Using Font: {f}")
                return f
        return None
    
    def resolve_font_by_name(self, font_name, text_to_render=None):
        """
        根据字体名称查找对应的字体文件
        
        Args:
            font_name: 字体名称，如"阿里巴巴普惠体H"或"标小智无界黑"
            text_to_render: 要渲染的文本（用于检测是否包含日文假名）
        
        Returns:
            str: 字体文件的完整路径
        """
        if not font_name:
            return self.font_path
        
        def contains_japanese_kana(text):
            """检测文本是否包含日文假名"""
            if not text:
                return False
            hiragana = re.search(r'[\u3040-\u309F]', text)
            katakana = re.search(r'[\u30A0-\u30FF]', text)
            return bool(hiragana or katakana)
        
        # 中文字体列表（不支持日文假名）
        chinese_only_fonts = ['阿里巴巴普惠体', '标小智无界黑', '思源黑体', '微软雅黑']
        
        # 日文兼容字体（Windows系统）
        japanese_fonts = [
            r"C:\Windows\Fonts\meiryo.ttc",
            r"C:\Windows\Fonts\msgothic.ttc",
            r"C:\Windows\Fonts\YuGothR.ttc",
            r"C:\Windows\Fonts\msmincho.ttc",
        ]
        
        # 1. 检查是否需要日文字体
        needs_japanese = text_to_render and contains_japanese_kana(text_to_render)
        is_chinese_font = any(cf in font_name for cf in chinese_only_fonts)
        
        if needs_japanese and is_chinese_font:
            logger.info(f"Text contains Japanese kana but font '{font_name}' is Chinese-only. Switching to Japanese font.")
            for jp_font in japanese_fonts:
                if os.path.exists(jp_font):
                    logger.info(f"Using Japanese font: {jp_font}")
                    return jp_font
        
        # 2. 在项目fonts目录查找
        fonts_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "fonts")
        if os.path.exists(fonts_dir):
            for file in os.listdir(fonts_dir):
                if file.endswith(('.ttf', '.ttc', '.otf')):
                    file_base = os.path.splitext(file)[0]
                    if font_name.lower() in file_base.lower() or file_base.lower() in font_name.lower():
                        font_path = os.path.join(fonts_dir, file)
                        logger.info(f"Found font in project dir: {font_path}")
                        return font_path
        
        # 3. 在Windows字体目录查找
        windows_fonts = r"C:\Windows\Fonts"
        if os.path.exists(windows_fonts):
            for file in os.listdir(windows_fonts):
                if file.endswith(('.ttf', '.ttc', '.otf')):
                    file_base = os.path.splitext(file)[0]
                    if font_name.lower() in file_base.lower() or file_base.lower() in font_name.lower():
                        font_path = os.path.join(windows_fonts, file)
                        logger.info(f"Found font in Windows: {font_path}")
                        return font_path
        
        # 4. 回退到默认字体
        logger.warning(f"Font '{font_name}' not found, using default")
        return self.font_path
    
    def fit_text(self, draw, text, max_w, max_h, font_path, start_size=100):
        """Fit text into a box by reducing font size"""
        if not font_path or not os.path.exists(font_path):
            return None, [text], 0
        
        size = start_size
        while size > 10:
            try:
                font = ImageFont.truetype(font_path, size)
            except:
                size -= 5
                continue
            
            lines = []
            current_line = ""
            for char in text:
                test_line = current_line + char
                bbox = draw.textbbox((0, 0), test_line, font=font)
                if bbox[2] - bbox[0] > max_w:
                    if current_line:
                        lines.append(current_line)
                    current_line = char
                else:
                    current_line = test_line
            if current_line:
                lines.append(current_line)
            
            total_h = 0
            for line in lines:
                bbox = draw.textbbox((0, 0), line, font=font)
                total_h += (bbox[3] - bbox[1]) * 1.3
            
            if total_h <= max_h:
                return font, lines, total_h
            
            size -= 5
            
        return None, [text], 0
    
    def draw_outlined_text(self, draw, x, y, text, font, text_color=(255, 255, 255), outline_color=(0, 0, 0), outline_width=5):
        """绘制带描边的文字"""
        for dx in range(-outline_width, outline_width + 1):
            for dy in range(-outline_width, outline_width + 1):
                if dx != 0 or dy != 0:
                    draw.text((x + dx, y + dy), text, font=font, fill=outline_color)
        draw.text((x, y), text, font=font, fill=text_color)
    
    def process_image(self, image_path, output_path, top_text, bottom_text, top_size=0, bottom_size=0, font_name=None, skip_resize=True):
        """
        处理图片：添加日文标签
        
        流程:
        1. 读取图片
        2. 添加顶部标签（粉色多边形）
        3. 添加底部文字（带描边）
        4. 保存
        
        Args:
            image_path: 输入图片路径
            output_path: 输出图片路径
            top_text: 顶部标签文字
            bottom_text: 底部文字
            top_size: 顶部文字大小（0=自动）
            bottom_size: 底部文字大小（0=自动）
            font_name: 字体名称
            skip_resize: 是否跳过resize（默认True，因为ComfyUI已处理）
        
        Returns:
            bool: 成功返回True
        """
        try:
            # 1. 读取图片
            try:
                img_array = np.fromfile(image_path, dtype=np.uint8)
                img = cv2.imdecode(img_array, cv2.IMREAD_COLOR)
            except Exception as e:
                logger.error(f"Failed to read image {image_path}: {e}")
                return False
            
            if img is None:
                logger.error(f"cv2.imdecode returned None for {image_path}")
                return False
            
            logger.info(f"Processing: {os.path.basename(image_path)}")
            
            current_h, current_w = img.shape[:2]
            logger.info(f"Image size: {current_w}x{current_h}")
            
            # 转换为PIL
            img_pil = Image.fromarray(cv2.cvtColor(img, cv2.COLOR_BGR2RGB))
            draw = ImageDraw.Draw(img_pil)
            W, H = img_pil.size
            
            # 解析字体
            current_font_path = self.resolve_font_by_name(font_name, (top_text or "") + (bottom_text or ""))
            if not current_font_path:
                current_font_path = self.font_path
            
            # 3. 添加顶部标签（粉色多边形）
            if top_text:
                label_color = (255, 128, 160)  # Pink
                font_size = int(top_size) if top_size > 0 else int(H * 0.05)
                
                try:
                    font_top = ImageFont.truetype(current_font_path, font_size)
                except:
                    font_top = ImageFont.load_default()
                
                bbox = draw.textbbox((0, 0), top_text, font=font_top)
                tw = bbox[2] - bbox[0]
                th = bbox[3] - bbox[1]
                
                lbl_h = int(th * 1.8)
                lbl_w = int(tw * 1.4) + 40
                y_offset = int(H * 0.03)
                
                # 绘制多边形背景
                points = [
                    (0, y_offset),
                    (lbl_w, y_offset),
                    (lbl_w - 40, y_offset + lbl_h),
                    (0, y_offset + lbl_h)
                ]
                draw.polygon(points, fill=label_color)
                
                # 绘制文字
                text_x = (lbl_w - 40) / 2 - tw / 2
                text_y = y_offset + (lbl_h / 2) - (th / 2) - 5
                draw.text((text_x, text_y), top_text, font=font_top, fill=(0, 0, 0))
                
                logger.info(f"Added top label: '{top_text}'")
            
            # 4. 添加底部文字（带描边）
            if bottom_text:
                area_h = int(H * 0.10)
                y_start = H - area_h
                start_size = int(bottom_size) if bottom_size > 0 else int(area_h * 0.6)
                
                font_bot, lines, text_h_final = self.fit_text(
                    draw, bottom_text, W * 0.9, area_h, current_font_path, start_size=start_size
                )
                
                if font_bot:
                    current_y = y_start + (area_h - text_h_final) / 2
                    line_spacing = 1.5
                    for line in lines:
                        bbox = draw.textbbox((0, 0), line, font=font_bot)
                        lh = bbox[3] - bbox[1]
                        self.draw_outlined_text(draw, 50, current_y, line, font_bot)
                        current_y += lh * line_spacing
                    
                    logger.info(f"Added bottom text: '{bottom_text}'")
            
            # 5. 保存
            final_img_cv2 = cv2.cvtColor(np.array(img_pil), cv2.COLOR_RGB2BGR)
            success, encoded_img = cv2.imencode(os.path.splitext(output_path)[1], final_img_cv2)
            if success:
                encoded_img.tofile(output_path)
                logger.info(f"Saved: {output_path}")
                return True
            else:
                logger.error(f"Failed to encode image")
                return False
        
        except Exception as e:
            logger.error(f"Error processing {image_path}: {e}")
            import traceback
            logger.error(traceback.format_exc())
            return False