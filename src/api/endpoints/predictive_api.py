from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import FileResponse
import os
from src.api.dependencies import get_upload_dir, get_output_dir, get_logger
import src.services.predictive_report_generator as predictive_api
import tempfile
from fastapi import APIRouter
from fastapi.responses import FileResponse

router = APIRouter()
UPLOAD_DIR = get_upload_dir()
OUTPUT_DIR = get_output_dir()
logger = get_logger()

@router.post("/generate-report/")
async def generate_report(file: UploadFile = File(...)):
    """Generate predictive attrition report from uploaded Excel file."""
    if not file.filename.endswith('.xlsx'):
        raise HTTPException(status_code=400, detail="Only Excel (.xlsx) files are supported")

    with tempfile.TemporaryDirectory() as temp_dir:
        try:
            content = await file.read()
            df = predictive_api.load_data(content, file.filename)
            report_path = predictive_api.create_predictive_report(df, OUTPUT_DIR)
            return {
                "message": f"Report generated successfully for {file.filename}",
                "download_url": f"/download/{os.path.basename(report_path)}"
            }
        except Exception as e:
            logger.error(f"Failed to process file {file.filename}: {e}")
            raise HTTPException(status_code=400, detail=f"Failed to process file: {str(e)}")

@router.get("/download/{filename}")
async def download_report(filename: str):
    """Download generated report."""
    file_path = f"{OUTPUT_DIR}/{filename}"
    if not os.path.exists(file_path):
        raise HTTPException(status_code=404, detail="Report not found")
    return FileResponse(file_path, media_type='routerlication/vnd.openxmlformats-officedocument.wordprocessingml.document', filename=filename)
