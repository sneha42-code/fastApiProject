from fastapi import APIRouter, UploadFile, File, HTTPException, Depends, BackgroundTasks
from fastapi.responses import FileResponse
from src.api.dependencies import get_upload_dir, get_output_dir, get_logger
import os
import uuid
import shutil

from src.services.excel_creator import create_excel_attrition_report, load_data



router = APIRouter()
UPLOAD_DIR = get_upload_dir()
OUTPUT_DIR = get_output_dir()
logger = get_logger()

@router.post("/upload-forExcel/")
async def upload_file(file: UploadFile = File(...)):
    """Upload an Excel file containing HRIS data for attrition analysis"""
    try:
        # Generate a unique ID for the uploaded file
        file_id = str(uuid.uuid4())
        file_location = f"{UPLOAD_DIR}/{file_id}_{file.filename}"
        
        # Save the uploaded file
        with open(file_location, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)
        
        return {
            "file_id": file_id,
            "filename": file.filename,
            "message": "File uploaded successfully. You can now generate a report."
        }
    except Exception as e:
        logger.error(f"File upload error: {e}")
        raise HTTPException(status_code=500, detail=f"File upload failed: {str(e)}")

@router.post("/generate-report-forExcel/")
async def generate_report(file_id: str, background_tasks: BackgroundTasks):
    """Generate an attrition report for the uploaded file"""
    try:
        # Find the file with the given ID
        files = [f for f in os.listdir(UPLOAD_DIR) if f.startswith(f"{file_id}_")]
        
        if not files:
            raise HTTPException(status_code=404, detail="File not found. Please upload the file first.")
        
        file_path = f"{UPLOAD_DIR}/{files[0]}"
        report_dir = f"{OUTPUT_DIR}/{file_id}"
        os.makedirs(report_dir, exist_ok=True)
        
        # Load data
        df = load_data(file_path)
        if df is None:
            raise HTTPException(status_code=500, detail="Failed to load data from the uploaded file.")
            
        # Generate report
        success, report_path, report_time = create_excel_attrition_report(df, report_dir)
        
        if success and report_path:
            return {
                "status": "success",
                "message": "Report generated successfully",
                "file_id": file_id,
                "report_file": os.path.basename(report_path),
                "download_url": f"/download/{file_id}/{os.path.basename(report_path)}"
            }
        else:
            raise HTTPException(status_code=500, detail="Failed to generate report.")
            
    except Exception as e:
        logger.error(f"Report generation error: {e}")
        raise HTTPException(status_code=500, detail=f"Report generation failed: {str(e)}")

@router.get("/download-forExcel/{file_id}/{filename}")
async def download_report(file_id: str, filename: str):
    """Download a generated report"""
    try:
        report_path = f"{OUTPUT_DIR}/{file_id}/{filename}"
        
        if not os.path.exists(report_path):
            raise HTTPException(status_code=404, detail="Report not found")
        
        return FileResponse(
            path=report_path,
            filename=filename,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        logger.error(f"Download error: {e}")
        raise HTTPException(status_code=500, detail=f"Download failed: {str(e)}")
