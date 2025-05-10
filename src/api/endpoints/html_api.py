
from fastapi import APIRouter, UploadFile, File, HTTPException, Depends, BackgroundTasks
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from src.api.dependencies import get_upload_dir, get_output_dir, get_logger
import os
import uuid
import shutil

from src.html import ReportResponse, UploadResponse, generate_html_report,load_data

router = APIRouter()
UPLOAD_DIR = get_upload_dir()
OUTPUT_DIR = get_output_dir()
logger = get_logger()


@router.post("/upload/", response_model=UploadResponse)
async def upload_file(file: UploadFile = File(...)):
    """
    Upload an Excel file containing HRIS data for attrition analysis
    """
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

@router.post("/generate-report/", response_model=ReportResponse)
async def generate_report(file_id: str, background_tasks: BackgroundTasks):
    """
    Generate an interactive HTML attrition report for the uploaded file
    """
    try:
        # Find the file with the given ID
        files = [f for f in os.listdir(UPLOAD_DIR) if f.startswith(f"{file_id}_")]
        
        if not files:
            raise HTTPException(status_code=404, detail="File not found. Please upload the file first.")
        
        file_path = f"{UPLOAD_DIR}/{files[0]}"
        
        # Load data
        df = load_data(file_path)
        if df is None:
            raise HTTPException(status_code=500, detail="Failed to load data from the uploaded file.")
            
        # Generate report
        success, report_path, report_filename = generate_html_report(df, OUTPUT_DIR, file_id)
        
        if success and report_path:
            return {
                "status": "success",
                "message": "Report generated successfully",
                "file_id": file_id,
                "report_file": report_filename,
                "download_url": f"/download/{file_id}/{report_filename}"
            }
        else:
            raise HTTPException(status_code=500, detail="Failed to generate report.")
            
    except Exception as e:
        logger.error(f"Report generation error: {e}")
        raise HTTPException(status_code=500, detail=f"Report generation failed: {str(e)}")

@router.get("/download/{file_id}/{filename}")
async def download_report(file_id: str, filename: str):
    """
    Download a generated HTML report
    """
    try:
        report_path = f"{OUTPUT_DIR}/{file_id}/{filename}"
        
        if not os.path.exists(report_path):
            raise HTTPException(status_code=404, detail="Report not found")
        
        return FileResponse(
            path=report_path,
            filename=filename,
            media_type="text/html"
        )
    except Exception as e:
        logger.error(f"Download error: {e}")
        raise HTTPException(status_code=500, detail=f"Download failed: {str(e)}")

@router.get("/view/{file_id}/{filename}", response_class=HTMLResponse)
async def view_report(file_id: str, filename: str):
    """
    View the HTML report directly in the browser
    """
    try:
        report_path = f"{OUTPUT_DIR}/{file_id}/{filename}"
        
        if not os.path.exists(report_path):
            raise HTTPException(status_code=404, detail="Report not found")
        
        with open(report_path, 'r', encoding='utf-8') as f:
            html_content = f.read()
            
        return HTMLResponse(content=html_content, status_code=200)
    except Exception as e:
        logger.error(f"View report error: {e}")
        raise HTTPException(status_code=500, detail=f"Failed to view report: {str(e)}")
    
    
    # Serve static files (reports directory)
router.mount("/reports", StaticFiles(directory=OUTPUT_DIR), name="reports")