from fastapi import APIRouter, UploadFile, File, HTTPException, Depends, BackgroundTasks
from fastapi.responses import FileResponse, HTMLResponse
from src.services.predictive_report_generator import generate_predictive_attrition_report , load_data
from src.services.predictive_generate_html import generate_predictive_html_report
from src.utils.file_handlers import save_upload_file
from src.api.dependencies import get_upload_dir, get_output_dir, get_logger
import os
import uuid

router = APIRouter()

@router.post("/predictive/upload/")
async def upload_predictive_file(
    file: UploadFile = File(...),
    upload_dir: str = Depends(get_upload_dir),
    logger = Depends(get_logger)
):
    """
    Upload an Excel file containing HRIS data for predictive attrition analysis
    """
    try:
        file_id = str(uuid.uuid4())
        file_location = f"{upload_dir}/{file_id}_{file.filename}"
        
        save_upload_file(file, file_location)
        
        return {
            "file_id": file_id,
            "filename": file.filename,
            "message": "File uploaded successfully. You can now generate predictive reports."
        }
    except Exception as e:
        logger.error(f"File upload error: {e}")
        raise HTTPException(status_code=500, detail=f"File upload failed: {str(e)}")

@router.post("/predictive/generate-report/")
def generate_predictive_report(
    file_id: str,
    background_tasks: BackgroundTasks,
    upload_dir: str = Depends(get_upload_dir),
    output_dir: str = Depends(get_output_dir),
    logger = Depends(get_logger)
):
    """
    Generate a predictive attrition report (Word document) for the uploaded file
    """
    try:
        # Find the file with the given ID
        files = [f for f in os.listdir(upload_dir) if f.startswith(f"{file_id}_")]
        
        if not files:
            raise HTTPException(status_code=404, detail="File not found. Please upload the file first.")
        
        file_path = f"{upload_dir}/{files[0]}"
        report_dir = f"{output_dir}/{file_id}"
        os.makedirs(report_dir, exist_ok=True)
        
        # Load data
        df = load_data(file_path, logger)
        if df is None:
            raise HTTPException(status_code=500, detail="Failed to load data from the uploaded file.")
            
        # Generate predictive report
        success, report_path, report_time = generate_predictive_attrition_report(df, report_dir, logger)
        
        if success and report_path:
            return {
                "status": "success",
                "message": "Predictive report generated successfully",
                "file_id": file_id,
                "report_file": os.path.basename(report_path),
                "download_url": f"/predictive/download/?file_id={file_id}&filename={os.path.basename(report_path)}"
            }
        else:
            raise HTTPException(status_code=500, detail="Failed to generate predictive report.")
            
    except Exception as e:
        logger.error(f"Predictive report generation error: {e}")
        raise HTTPException(status_code=500, detail=f"Predictive report generation failed: {str(e)}")

@router.post("/predictive/generate-html/")
def generate_predictive_html(
    file_id: str,
    background_tasks: BackgroundTasks,
    upload_dir: str = Depends(get_upload_dir),
    output_dir: str = Depends(get_output_dir),
    logger = Depends(get_logger)
):
    """
    Generate a predictive attrition HTML report for the uploaded file
    """
    try:
        # Find the file with the given ID
        files = [f for f in os.listdir(upload_dir) if f.startswith(f"{file_id}_")]
        
        if not files:
            raise HTTPException(status_code=404, detail="File not found. Please upload the file first.")
        
        file_path = f"{upload_dir}/{files[0]}"
        report_dir = f"{output_dir}/{file_id}"
        os.makedirs(report_dir, exist_ok=True)
        
        # Load data
        df = load_data(file_path, logger)
        if df is None:
            raise HTTPException(status_code=500, detail="Failed to load data from the uploaded file.")
            
        # Generate HTML predictive report
        success, report_path, report_filename = generate_predictive_html_report(df, output_dir, file_id, logger)
        
        if success and report_path:
            return {
                "status": "success",
                "message": "Predictive HTML report generated successfully",
                "file_id": file_id,
                "report_file": report_filename,
                "download_url": f"/predictive/download-html/{file_id}/{report_filename}",
                "view_url": f"/predictive/view/{file_id}/{report_filename}"
            }
        else:
            raise HTTPException(status_code=500, detail="Failed to generate predictive HTML report.")
            
    except Exception as e:
        logger.error(f"Predictive HTML report generation error: {e}")
        raise HTTPException(status_code=500, detail=f"Predictive HTML report generation failed: {str(e)}")

@router.get("/predictive/download/")
async def download_predictive_report(
    file_id: str,
    filename: str,
    output_dir: str = Depends(get_output_dir),
    logger = Depends(get_logger)
):
    """
    Download a generated predictive report (Word document)
    """
    try:
        report_path = f"{output_dir}/{file_id}/{filename}"
        
        if not os.path.exists(report_path):
            raise HTTPException(status_code=404, detail="Predictive report not found")
        
        return FileResponse(
            path=report_path,
            filename=filename,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    except Exception as e:
        logger.error(f"Download error: {e}")
        raise HTTPException(status_code=500, detail=f"Download failed: {str(e)}")

@router.get("/predictive/download-html/{file_id}/{filename}")
async def download_predictive_html_report(
    file_id: str, 
    filename: str,
    output_dir: str = Depends(get_output_dir),
    logger = Depends(get_logger)
):
    """
    Download a generated predictive HTML report
    """
    try:
        report_path = f"{output_dir}/{file_id}/{filename}"
        
        if not os.path.exists(report_path):
            raise HTTPException(status_code=404, detail="Predictive HTML report not found")
        
        return FileResponse(
            path=report_path,
            filename=filename,
            media_type="text/html"
        )
    except Exception as e:
        logger.error(f"Download error: {e}")
        raise HTTPException(status_code=500, detail=f"Download failed: {str(e)}")

@router.get("/predictive/view/{file_id}/{filename}", response_class=HTMLResponse)
async def view_predictive_report(
    file_id: str, 
    filename: str,
    output_dir: str = Depends(get_output_dir),
    logger = Depends(get_logger)
):
    """
    View the predictive HTML report directly in the browser
    """
    try:
        report_path = f"{output_dir}/{file_id}/{filename}"
        
        if not os.path.exists(report_path):
            raise HTTPException(status_code=404, detail="Predictive HTML report not found")
        
        with open(report_path, 'r', encoding='utf-8') as f:
            html_content = f.read()
            
        return HTMLResponse(content=html_content, status_code=200)
    except Exception as e:
        logger.error(f"View predictive report error: {e}")
        raise HTTPException(status_code=500, detail=f"Failed to view predictive report: {str(e)}")

@router.get("/predictive/health")
async def health_check_predictive():
    """Health check for predictive analytics service"""
    return {"status": "healthy", "service": "predictive-analytics"}