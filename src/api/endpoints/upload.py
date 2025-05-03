from fastapi import APIRouter, UploadFile, File, HTTPException, Depends
from src.utils.file_handlers import save_upload_file
from src.api.dependencies import get_upload_dir, get_logger
import uuid

router = APIRouter()

@router.post("/upload/")
async def upload_file(
    file: UploadFile = File(...),
    upload_dir: str = Depends(get_upload_dir),
    logger = Depends(get_logger)
):
    """
    Upload an Excel file containing HRIS data for attrition analysis
    """
    try:
        file_id = str(uuid.uuid4())
        file_location = f"{upload_dir}/{file_id}_{file.filename}"
        
        save_upload_file(file, file_location)
        
        return {
            "file_id": file_id,
            "filename": file.filename,
            "message": "File uploaded successfully. You can now generate a report."
        }
    except Exception as e:
        logger.error(f"File upload error: {e}")
        raise HTTPException(status_code=500, detail=f"File upload failed: {str(e)}")
