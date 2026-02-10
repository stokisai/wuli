import os
from deep_translator import GoogleTranslator
import logging
import requests

logger = logging.getLogger(__name__)

def setup_logging(log_file="process.log"):
    logging.basicConfig(
        filename=log_file,
        filemode='w',
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        encoding='utf-8',
        force=True
    )
    # Also print to console
    console = logging.StreamHandler()
    console.setLevel(logging.INFO)
    logging.getLogger('').addHandler(console)
    return logging.getLogger('')

def ensure_dir(path):
    if not os.path.exists(path):
        os.makedirs(path)

def check_and_download_font(font_path):
    # This is a placeholder. In production, we'd download a Noto Sans JP font.
    # For now, we assume the user puts a valid font there, or we instruct them.
    # If file doesn't exist, we might try to download or warn.
    if not os.path.exists(font_path):
        logger.warning(f"Font file not found at {font_path}. Text rendering might fail or look bad.")
        # Ensure dir exists at least
        ensure_dir(os.path.dirname(font_path))

def translate_text(text, target='ja'):
    """
    Translate text to target language using deep-translator (Google Translate free).
    """
    try:
        if not text or len(text.strip()) == 0:
            return ""
        
        # Simple cache or check? For now just call directly.
        # deep-translator handles chunks automatically for extremely long text, 
        # but for OCR snippets it's fine.
        translated = GoogleTranslator(source='auto', target=target).translate(text)
        return translated
    except Exception as e:
        logger.error(f"Translation failed for '{text}': {e}")
        return text # Return original on failure (fuzzy match might still catch English if pool has English)