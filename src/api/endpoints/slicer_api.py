from fastapi import APIRouter, UploadFile, File, HTTPException, Depends, BackgroundTasks, Form
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from src.api.dependencies import get_upload_dir, get_output_dir, get_logger
import os
import uuid
import shutil
from pathlib import Path

from src.services.slicer_html import ReportResponse, UploadResponse,load_data,generate_interactive_html_report

router = APIRouter()
UPLOAD_DIR = Path(get_upload_dir())
OUTPUT_DIR = Path(get_output_dir())
logger = get_logger()


# FastAPI endpoints
@router.post("/slicer-upload/", response_model=UploadResponse)
async def upload_file(file: UploadFile = File(...)):
    """
    Upload an Excel file for processing
    """
    try:
        if not file.filename.endswith(('.xlsx', '.xls')):
            raise HTTPException(status_code=400, detail="Only Excel files are supported")
        
        file_id = str(uuid.uuid4())
        file_path = UPLOAD_DIR / f"{file_id}_{file.filename}"
        
        with open(file_path, "wb") as buffer:
            shutil.copyfileobj(file.file, buffer)
        
        return UploadResponse(
            file_id=file_id,
            filename=file.filename,
            message="File uploaded successfully"
        )
    except Exception as e:
        logger.error(f"Error uploading file: {e}")
        raise HTTPException(status_code=500, detail=f"Error uploading file: {str(e)}")

@router.post("/slicer-generate-report", response_model=ReportResponse)
async def generate_report(file_id: str = Form(...), background_tasks: BackgroundTasks = None):
    """
    Generate attrition analysis report from uploaded file
    """
    try:
        # Find the uploaded file
        file_path = None
        for path in UPLOAD_DIR.glob(f"{file_id}_*"):
            file_path = path
            break
        
        if not file_path or not file_path.exists():
            raise HTTPException(status_code=404, detail="File not found")
        
        # Load and process data
        df = load_data(file_path)
        if df is None:
            raise HTTPException(status_code=400, detail="Failed to load data from file")
        
        # Generate report
        try:
            success, html_path, report_filename = generate_interactive_html_report(df, OUTPUT_DIR, file_id)
        except Exception as e:
            logger.error(f"Exception in generate_interactive_html_report: {e}")
            raise HTTPException(status_code=500, detail=f"Exception in report generation: {str(e)}")
        
        if not success:
            logger.error(f"generate_interactive_html_report failed for file_id={file_id}")
            raise HTTPException(status_code=500, detail="Failed to generate report (see server logs for details)")
        
        # Clean up uploaded file in background
        background_tasks.add_task(os.remove, file_path)
        
        return ReportResponse(
            status="success",
            message="Report generated successfully",
            file_id=file_id,
            report_file=report_filename,
            download_url=f"/slicer-download/{file_id}/{report_filename}"
        )
    except Exception as e:
        logger.error(f"Error generating report: {e}")
        raise HTTPException(status_code=500, detail=f"Error generating report: {str(e)}")

@router.get("/slicer-download/{file_id}/{filename}", response_class=FileResponse)
async def download_report(file_id: str, filename: str):
    """
    Download the generated report
    """
    report_path = OUTPUT_DIR / file_id / filename
    if not report_path.exists():
        raise HTTPException(status_code=404, detail="Report not found")
    
    return FileResponse(
        path=report_path,
        filename=filename,
        media_type='text/html'
    )

@router.get("/", response_class=HTMLResponse)
async def get_upload_form():
    """
    Serve the upload form
    """
    html_content = """
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Attrition Analysis Dashboard</title>
        <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
        <style>
            body {
                font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
                background-color: #f8f9fa;
            }
            .container {
                max-width: 600px;
                margin-top: 50px;
            }
            .card {
                box-shadow: 0 2px 5px rgba(0,0,0,0.1);
                border: none;
                border-radius: 8px;
            }
            .btn-primary {
                background-color: #004C99;
                border-color: #004C99;
            }
            .btn-primary:hover {
                background-color: #4F81BD;
                border-color: #4F81BD;
            }
        </style>
    </head>
    <body>
        <div class="container">
            <div class="card">
                <div class="card-header bg-primary text-white">
                    <h3 class="mb-0">Upload Excel File for Attrition Analysis</h3>
                </div>
                <div class="card-body">
                    <form id="uploadForm" enctype="multipart/form-data">
                        <div class="mb-3">
                            <label for="file" class="form-label">Select Excel File (.xlsx, .xls)</label>
                            <input class="form-control" type="file" id="file" name="file" accept=".xlsx,.xls" required>
                        </div>
                        <button type="submit" class="btn btn-primary w-100">Upload and Generate Report</button>
                    </form>
                    <div id="status" class="mt-3"></div>
                </div>
            </div>
        </div>
        <script>
            document.getElementById('uploadForm').addEventListener('submit', async (e) => {
                e.preventDefault();
                const formData = new FormData(e.target);
                const statusDiv = document.getElementById('status');
                
                statusDiv.innerHTML = '<div class="alert alert-info">Uploading file...</div>';
                
                try {
                    const uploadResponse = await fetch('/upload', {
                        method: 'POST',
                        body: formData
                    });
                    
                    if (!uploadResponse.ok) {
                        throw new Error('File upload failed');
                    }
                    
                    const uploadResult = await uploadResponse.json();
                    statusDiv.innerHTML = '<div class="alert alert-success">File uploaded successfully. Generating report...</div>';
                    
                    const reportResponse = await fetch('/slicer-generate-report', {
                        method: 'POST',
                        headers: { 'Content-Type': 'routerlication/x-www-form-urlencoded' },
                        body: `file_id=${uploadResult.file_id}`
                    });
                    
                    if (!reportResponse.ok) {
                        throw new Error('Report generation failed');
                    }
                    
                    const reportResult = await reportResponse.json();
                    statusDiv.innerHTML = `
                        <div class="alert alert-success">
                            Report generated successfully!
                            <a href="${reportResult.download_url}" class="btn btn-primary btn-sm mt-2">Download Report</a>
                        </div>
                    `;
                } catch (error) {
                    statusDiv.innerHTML = `<div class="alert alert-danger">Error: ${error.message}</div>`;
                }
            });
        </script>
    </body>
    </html>
    """
    return HTMLResponse(content=html_content)

