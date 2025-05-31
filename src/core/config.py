from pydantic_settings import BaseSettings
from pathlib import Path
import os

class Settings(BaseSettings):
    # App settings
    APP_TITLE: str = "Attrition Analysis API"
    APP_DESCRIPTION: str = "API for generating automated attrition analysis reports"
    APP_VERSION: str = "1.0.0"
    
    # Server settings
    HOST: str = "0.0.0.0"
    PORT: int = 8000
    WORKERS: int = 4
    
    # CORS settings
    CORS_ORIGINS: list = ["http://localhost:3000", "https://automateroperting", 
                          "https://sneha42-code.github.io", "*"]
    
    # Directories
    BASE_DIR: Path = Path(__file__).parent.parent.parent
    UPLOAD_DIR: str = str(BASE_DIR / "file_uploads")
    OUTPUT_DIR: str = str(BASE_DIR / "attrition_reports")
    IMAGES_DIR : str =str(BASE_DIR / "images_reports")
    
    class Config:
        env_file = ".env"

settings = Settings()

# Create directories if they don't exist
os.makedirs(settings.UPLOAD_DIR, exist_ok=True)
os.makedirs(settings.OUTPUT_DIR, exist_ok=True)
