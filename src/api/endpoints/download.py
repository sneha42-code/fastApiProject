from fastapi import APIRouter, HTTPException, Depends
from fastapi.responses import FileResponse
from src.api.dependencies import get_output_dir, get_logger
import os

router = APIRouter()

@router.get("/download/")
async def download_report(
    file_id: str,
    filename: str,
    output_dir: str = Depends(get_output_dir),
    logger = Depends(get_logger)
):
    """
    Download a generated report
    """
    try:
        report_path = f"{output_dir}/{file_id}/{filename}"
        
        if not os.path.exists(report_path):
            raise HTTPException(status_code=404, detail="Report not found")
        
        return FileResponse(
            path=report_path,
            filename=filename,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    except Exception as e:
        logger.error(f"Download error: {e}")
        raise HTTPException(status_code=500, detail=f"Download failed: {str(e)}")
