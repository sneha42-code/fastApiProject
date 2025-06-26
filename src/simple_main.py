import os
import json
from datetime import datetime
import pandas as pd
import logging
import shutil
import uuid
from fastapi import FastAPI, File, UploadFile, HTTPException
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
import uvicorn

# Setup logging
def setup_logging():
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s'
    )
    return logging.getLogger(__name__)

logger = setup_logging()

app = FastAPI(title="Attrition Analysis Dashboard API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

UPLOAD_DIR = "uploads"
REPORT_DIR = "reports"
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(REPORT_DIR, exist_ok=True)

class UploadResponse(BaseModel):
    file_id: str
    filename: str
    message: str

class ReportResponse(BaseModel):
    status: str
    message: str
    file_id: str
    report_file: str
    download_url: str

# Dummy analysis/report function for demonstration
from src.services.slicer_html import generate_interactive_html_report

def load_data(file_path):
    try:
        df = pd.read_excel(file_path)
        df['Action Type'] = df['Action Type'].fillna('')
        logger.info(f"Successfully loaded data from {file_path}")
        return df
    except Exception as e:
        logger.error(f"Error loading data: {e}")
        return None

@app.get("/")
async def root():
    return {"message": "Attrition Analysis Dashboard API", "version": "1.0.0"}

@app.post("/upload", response_model=UploadResponse)
async def upload_file(file: UploadFile = File(...)):
    try:
        if not file.filename.endswith((".xlsx", ".xls")):
            raise HTTPException(status_code=400, detail="Only Excel files are allowed")
        file_id = str(uuid.uuid4())
        file_path = os.path.join(UPLOAD_DIR, f"{file_id}_{file.filename}")
        with open(file_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)
        logger.info(f"File uploaded successfully: {file_path}")
        return UploadResponse(
            file_id=file_id,
            filename=file.filename,
            message="File uploaded successfully"
        )
    except Exception as e:
        logger.error(f"Error uploading file: {e}")
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/generate-report/{file_id}", response_model=ReportResponse)
async def generate_report(file_id: str):
    try:
        upload_files = os.listdir(UPLOAD_DIR)
        file_path = None
        for f in upload_files:
            if f.startswith(file_id):
                file_path = os.path.join(UPLOAD_DIR, f)
                break
        if not file_path:
            raise HTTPException(status_code=404, detail="File not found")
        df = load_data(file_path)
        if df is None:
            raise HTTPException(status_code=400, detail="Failed to load data from file")
        success, report_path, report_filename = generate_interactive_html_report(df, REPORT_DIR, file_id)
        if not success:
            raise HTTPException(status_code=500, detail="Failed to generate report")
        return ReportResponse(
            status="success",
            message="Report generated successfully",
            file_id=file_id,
            report_file=report_filename,
            download_url=f"/download/{file_id}/{report_filename}"
        )
    except Exception as e:
        logger.error(f"Error generating report: {e}")
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/download/{file_id}/{filename}")
async def download_report(file_id: str, filename: str):
    try:
        file_path = os.path.join(REPORT_DIR, file_id, filename)
        if not os.path.exists(file_path):
            raise HTTPException(status_code=404, detail="Report not found")
        return FileResponse(
            path=file_path,
            filename=filename,
            media_type='text/html'
        )
    except Exception as e:
        logger.error(f"Error downloading report: {e}")
        raise HTTPException(status_code=500, detail=str(e))

@app.delete("/cleanup/{file_id}")
async def cleanup_files(file_id: str):
    try:
        upload_files = os.listdir(UPLOAD_DIR)
        for f in upload_files:
            if f.startswith(file_id):
                os.remove(os.path.join(UPLOAD_DIR, f))
        report_dir = os.path.join(REPORT_DIR, file_id)
        if os.path.exists(report_dir):
            shutil.rmtree(report_dir)
        return {"message": "Files cleaned up successfully"}
    except Exception as e:
        logger.error(f"Error cleaning up files: {e}")
        raise HTTPException(status_code=500, detail=str(e))

if __name__ == "__main__":
    uvicorn.run("simple_main:app", host="0.0.0.0", port=8000, reload=True)
