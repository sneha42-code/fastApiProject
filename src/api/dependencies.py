from src.core.config import settings
from src.core.logging import setup_logging

def get_logger():
    return setup_logging()

def get_upload_dir():
    return settings.UPLOAD_DIR

def get_output_dir():
    return settings.OUTPUT_DIR

def get_images_dir():
    return settings.IMAGES_DIR
