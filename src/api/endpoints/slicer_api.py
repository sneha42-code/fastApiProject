from fastapi import APIRouter, UploadFile, File, HTTPException, Depends, BackgroundTasks, Form
from fastapi.responses import FileResponse, HTMLResponse
from src.api.dependencies import get_upload_dir, get_output_dir, get_logger
from src.services.slicer_html import (
    ReportResponse, 
    UploadResponse, 
    load_data, 
    generate_interactive_html_report
)
import os
import uuid
import shutil
from pathlib import Path

router = APIRouter()

@router.post("/slicer/upload/", response_model=UploadResponse)
async def upload_file_interactive(
    file: UploadFile = File(...),
    upload_dir: str = Depends(get_upload_dir),
    logger = Depends(get_logger)
):
    """
    Upload an Excel file for interactive dashboard processing
    """
    try:
        if not file.filename.endswith(('.xlsx', '.xls')):
            raise HTTPException(status_code=400, detail="Only Excel files are supported")
        
        file_id = str(uuid.uuid4())
        file_path = Path(upload_dir) / f"{file_id}_{file.filename}"
        
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

@router.post("/slicer/generate-report/", response_model=ReportResponse)
async def generate_interactive_report(
    file_id: str = Form(...),
    background_tasks: BackgroundTasks = None,
    upload_dir: str = Depends(get_upload_dir),
    output_dir: str = Depends(get_output_dir),
    logger = Depends(get_logger)
):
    """
    Generate interactive attrition analysis dashboard from uploaded file
    """
    try:
        # Find the uploaded file
        upload_path = Path(upload_dir)
        file_path = None
        for path in upload_path.glob(f"{file_id}_*"):
            file_path = path
            break
        
        if not file_path or not file_path.exists():
            raise HTTPException(status_code=404, detail="File not found")
        
        # Load and process data
        df = load_data(file_path, logger)
        if df is None:
            raise HTTPException(status_code=400, detail="Failed to load data from file")
        
        # Generate interactive dashboard
        success, html_path, report_filename = generate_interactive_html_report(df, output_dir, file_id, logger)
        
        if not success:
            raise HTTPException(status_code=500, detail="Failed to generate interactive dashboard")
        
        # Clean up uploaded file in background
        if background_tasks:
            background_tasks.add_task(os.remove, file_path)
        
        return ReportResponse(
            status="success",
            message="Interactive dashboard generated successfully",
            file_id=file_id,
            report_file=report_filename,
            download_url=f"/api/v1/slicer/download/{file_id}/{report_filename}"
        )
    except Exception as e:
        logger.error(f"Error generating interactive dashboard: {e}")
        raise HTTPException(status_code=500, detail=f"Error generating dashboard: {str(e)}")

@router.get("/slicer/download/{file_id}/{filename}")
async def download_interactive_report(
    file_id: str, 
    filename: str,
    output_dir: str = Depends(get_output_dir),
    logger = Depends(get_logger)
):
    """
    Download the generated interactive dashboard
    """
    try:
        # Clean the file_id to remove any URL encoding issues
        clean_file_id = file_id.strip().replace('"', '').replace('%22', '')
        clean_filename = filename.strip()
        
        logger.info(f"Download request: file_id={clean_file_id}, filename={clean_filename}")
        
        # Try multiple possible paths
        possible_paths = [
            Path(output_dir) / clean_file_id / clean_filename,
            Path(output_dir) / file_id / filename,  # Original path
            Path(output_dir) / clean_filename,  # Direct file path
        ]
        
        report_path = None
        for path in possible_paths:
            if path.exists():
                report_path = path
                break
        
        if not report_path:
            raise HTTPException(status_code=404, detail="Interactive dashboard not found")
        
        return FileResponse(
            path=report_path,
            filename=clean_filename,
            media_type='text/html'
        )
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Download error: {e}")
        raise HTTPException(status_code=500, detail=f"Download failed: {str(e)}")

@router.get("/slicer/view/{file_id}/{filename}", response_class=HTMLResponse)
async def view_interactive_dashboard(
    file_id: str, 
    filename: str,
    output_dir: str = Depends(get_output_dir),
    logger = Depends(get_logger)
):
    """
    View the interactive dashboard directly in the browser
    """
    try:
        # Clean the file_id to remove any URL encoding issues
        clean_file_id = file_id.strip().replace('"', '').replace('%22', '')
        clean_filename = filename.strip()
        
        logger.info(f"Looking for dashboard: file_id={clean_file_id}, filename={clean_filename}")
        
        # Try multiple possible paths
        possible_paths = [
            Path(output_dir) / clean_file_id / clean_filename,
            Path(output_dir) / file_id / filename,  # Original path
            Path(output_dir) / clean_filename,  # Direct file path
        ]
        
        report_path = None
        for path in possible_paths:
            logger.info(f"Checking path: {path}")
            if path.exists():
                report_path = path
                logger.info(f"Found dashboard at: {path}")
                break
        
        if not report_path:
            # List available files for debugging
            try:
                output_path = Path(output_dir)
                if output_path.exists():
                    available_dirs = [d.name for d in output_path.iterdir() if d.is_dir()]
                    logger.error(f"Available directories in {output_dir}: {available_dirs}")
                    
                    # Check specific directory
                    file_dir = output_path / clean_file_id
                    if file_dir.exists():
                        available_files = [f.name for f in file_dir.iterdir() if f.is_file()]
                        logger.error(f"Available files in {file_dir}: {available_files}")
                else:
                    logger.error(f"Output directory does not exist: {output_dir}")
            except Exception as debug_e:
                logger.error(f"Error during debugging: {debug_e}")
                
            raise HTTPException(status_code=404, detail=f"Interactive dashboard not found. Checked paths: {[str(p) for p in possible_paths]}")
        
        with open(report_path, 'r', encoding='utf-8') as f:
            html_content = f.read()
            
        return HTMLResponse(content=html_content, status_code=200)
    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"View dashboard error: {e}")
        raise HTTPException(status_code=500, detail=f"Failed to view dashboard: {str(e)}")

@router.get("/slicer/upload-form/", response_class=HTMLResponse)
async def get_interactive_upload_form():
    """
    Serve the upload form for interactive dashboard
    """
    html_content = """
    <!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Interactive Attrition Analysis Dashboard</title>
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
            #statusMessage {
                display: none;
                margin-top: 20px;
            }
        </style>
    </head>
    <body>
        <div class="container">
            <div class="card p-4">
                <h2 class="text-center mb-4">Upload Excel File for Interactive Attrition Dashboard</h2>
                <form id="uploadForm">
                    <div class="mb-3">
                        <label for="fileInput" class="form-label">Select Excel File (.xlsx, .xls)</label>
                        <input class="form-control" type="file" id="fileInput" name="file" accept=".xlsx,.xls" required>
                    </div>
                    <button type="submit" class="btn btn-primary w-100">Upload and Generate Interactive Dashboard</button>
                </form>
                <div id="statusMessage" class="alert"></div>
            </div>
        </div>
        <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
        <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
        <script>
            $(document).ready(function() {
                $('#uploadForm').on('submit', function(e) {
                    e.preventDefault();
                    const formData = new FormData();
                    const fileInput = $('#fileInput')[0].files[0];
                    formData.append('file', fileInput);
                    
                    const statusMessage = $('#statusMessage');
                    statusMessage.removeClass('alert-success alert-danger').hide();
                    
                    $.ajax({
                        url: '/api/v1/slicer/upload/',
                        type: 'POST',
                        data: formData,
                        processData: false,
                        contentType: false,
                        success: function(response) {
                            const fileId = response.file_id;
                            $.ajax({
                                url: '/api/v1/slicer/generate-report/',
                                type: 'POST',
                                data: { file_id: fileId },
                                success: function(reportResponse) {
                                    statusMessage.addClass('alert-success')
                                        .html(`Dashboard generated! <a href="${reportResponse.download_url}" class="alert-link">Download</a> | <a href="/api/v1/slicer/view/${fileId}/${reportResponse.report_file}" class="alert-link" target="_blank">View Dashboard</a>`)
                                        .show();
                                },
                                error: function(xhr) {
                                    statusMessage.addClass('alert-danger')
                                        .text('Error generating dashboard: ' + (xhr.responseJSON?.detail || 'Unknown error'))
                                        .show();
                                }
                            });
                        },
                        error: function(xhr) {
                            statusMessage.addClass('alert-danger')
                                .text('Error uploading file: ' + (xhr.responseJSON?.detail || 'Unknown error'))
                                .show();
                        }
                    });
                });
            });
        </script>
    </body>
    </html>
    """
    return HTMLResponse(content=html_content)

# @router.get("/slicer/health")
# async def health_check_interactive():
#     """Health check for interactive dashboard service"""
#     return {"status": "healthy", "service": "interactive-dashboard"}