from fastapi import APIRouter, BackgroundTasks, HTTPException, Depends
from src.services.report_generator import generate_attrition_report
from src.services.data_processor import load_data
from src.api.dependencies import get_upload_dir, get_output_dir, get_logger
import os

router = APIRouter()

@router.post("/generate-report/")
def generate_report(
    file_id: str,
    background_tasks: BackgroundTasks,
    upload_dir: str = Depends(get_upload_dir),
    output_dir: str = Depends(get_output_dir),
    logger = Depends(get_logger)
):
    """
    Generate an attrition report for the uploaded file
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
            
        # Generate report
        success, report_path, report_time = generate_attrition_report(df, report_dir, logger)
        
        if success and report_path:
            return {
                "status": "success",
                "message": "Report generated successfully",
                "file_id": file_id,
                "report_file": os.path.basename(report_path),
                "download_url": f"/api/download/?file_id={file_id}&filename={os.path.basename(report_path)}"
            }
        else:
            raise HTTPException(status_code=500, detail="Failed to generate report.")
            
    except Exception as e:
        logger.error(f"Report generation error: {e}")
        raise HTTPException(status_code=500, detail=f"Report generation failed: {str(e)}")
