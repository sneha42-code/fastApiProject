from fastapi import APIRouter, UploadFile, File, HTTPException, Depends, BackgroundTasks
from fastapi.responses import FileResponse, HTMLResponse
from src.services.predictive_report_generator import generate_enhanced_predictive_attrition_report ,load_data
from src.services.predictive_generate_html import generate_predictive_html_report 
from src.utils.file_handlers import save_upload_file
from src.api.dependencies import get_upload_dir, get_output_dir, get_logger
import os
import uuid
import shutil
import tempfile
from pathlib import Path
import pandas as pd
import xlsxwriter

router = APIRouter()

@router.post("/predictive/upload/")
async def upload_file_enhanced_predictive(
    file: UploadFile = File(...),
    upload_dir: str = Depends(get_upload_dir),
    logger = Depends(get_logger)
):
    """
    Upload an Excel file containing HRIS data for enhanced predictive attrition analysis
    """
    try:
        # Validate file type
        if not file.filename.endswith(('.xlsx', '.xls')):
            raise HTTPException(status_code=400, detail="Only Excel files (.xlsx, .xls) are supported")
        
        file_id = str(uuid.uuid4())
        file_location = f"{upload_dir}/{file_id}_{file.filename}"
        
        save_upload_file(file, file_location)
        
        return {
            "file_id": file_id,
            "filename": file.filename,
            "message": "File uploaded successfully. You can now generate enhanced predictive reports."
        }
    except Exception as e:
        logger.error(f"File upload error: {e}")
        raise HTTPException(status_code=500, detail=f"File upload failed: {str(e)}")

@router.post("/predictive/generate-report/")
def generate_enhanced_predictive_report(
    file_id: str,
    background_tasks: BackgroundTasks,
    upload_dir: str = Depends(get_upload_dir),
    output_dir: str = Depends(get_output_dir),
    logger = Depends(get_logger)
):
    """
    Generate an enhanced predictive attrition report (Word document) for the uploaded file
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
        
        # Validate required columns for predictive analysis
        required_columns = ['Employee Name', 'Gender', 'Function', 'Job Location', 'Grade', 'Action Date', 'Date of Joining', 'Action Type']
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            raise HTTPException(
                status_code=400, 
                detail=f"Missing required columns for predictive analysis: {', '.join(missing_columns)}"
            )
            
        # Generate enhanced predictive report
        success, report_path, report_time = generate_enhanced_predictive_attrition_report(df, report_dir, logger)
        
        if success and report_path:
            return {
                "status": "success",
                "message": "Enhanced predictive report generated successfully",
                "file_id": file_id,
                "report_file": os.path.basename(report_path),
                "download_url": f"/predictive/api/v1/download/?file_id={file_id}&filename={os.path.basename(report_path)}"
            }
        else:
            raise HTTPException(status_code=500, detail="Failed to generate enhanced predictive report.")
            
    except Exception as e:
        logger.error(f"Enhanced predictive report generation error: {e}")
        raise HTTPException(status_code=500, detail=f"Enhanced predictive report generation failed: {str(e)}")

@router.post("/predictive/generate-html/")
def generate_enhanced_predictive_html(
    file_id: str,
    background_tasks: BackgroundTasks,
    upload_dir: str = Depends(get_upload_dir),
    output_dir: str = Depends(get_output_dir),
    logger = Depends(get_logger)
):
    """
    Generate an enhanced predictive attrition HTML report for the uploaded file
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
        
        # Validate required columns
        required_columns = ['Employee Name', 'Gender', 'Function', 'Job Location', 'Grade', 'Action Date', 'Date of Joining', 'Action Type']
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            raise HTTPException(
                status_code=400, 
                detail=f"Missing required columns for predictive analysis: {', '.join(missing_columns)}"
            )
            
        # Generate HTML predictive report
        success, report_path, report_filename = generate_predictive_html_report(df, output_dir, file_id, logger)
        
        if success and report_path:
            return {
                "status": "success",
                "message": "Enhanced predictive HTML report generated successfully",
                "file_id": file_id,
                "report_file": report_filename,
                "download_url": f"/api/v1/predictive/download-html/{file_id}/{report_filename}",
                "view_url": f"/api/v1/predictive/view/{file_id}/{report_filename}"
            }
        else:
            raise HTTPException(status_code=500, detail="Failed to generate enhanced predictive HTML report.")
            
    except Exception as e:
        logger.error(f"Enhanced predictive HTML report generation error: {e}")
        raise HTTPException(status_code=500, detail=f"Enhanced predictive HTML report generation failed: {str(e)}")

@router.post("/predictive/generate-excel/")
def generate_enhanced_predictive_excel(
    file_id: str,
    background_tasks: BackgroundTasks,
    upload_dir: str = Depends(get_upload_dir),
    output_dir: str = Depends(get_output_dir),
    logger = Depends(get_logger)
):
    """
    Generate an enhanced predictive attrition Excel report for the uploaded file
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
        
        # Validate required columns
        required_columns = ['Employee Name', 'Gender', 'Function', 'Job Location', 'Grade', 'Action Date', 'Date of Joining', 'Action Type']
        missing_columns = [col for col in required_columns if col not in df.columns]
        if missing_columns:
            raise HTTPException(
                status_code=400, 
                detail=f"Missing required columns for predictive analysis: {', '.join(missing_columns)}"
            )
        
        # Generate Excel report with predictive data
        excel_path = create_enhanced_predictive_excel_report(df, report_dir, logger)
        
        if excel_path:
            return {
                "status": "success",
                "message": "Enhanced predictive Excel report generated successfully",
                "file_id": file_id,
                "report_file": os.path.basename(excel_path),
                "download_url": f"/api/v1/predictive/download-excel/{file_id}/{os.path.basename(excel_path)}"
            }
        else:
            raise HTTPException(status_code=500, detail="Failed to generate enhanced predictive Excel report.")
            
    except Exception as e:
        logger.error(f"Enhanced predictive Excel report generation error: {e}")
        raise HTTPException(status_code=500, detail=f"Enhanced predictive Excel report generation failed: {str(e)}")

@router.get("/predictive/download/")
async def download_enhanced_predictive_report(
    file_id: str,
    filename: str,
    output_dir: str = Depends(get_output_dir),
    logger = Depends(get_logger)
):
    """
    Download a generated enhanced predictive report (Word document)
    """
    try:
        report_path = f"{output_dir}/{file_id}/{filename}"
        
        if not os.path.exists(report_path):
            raise HTTPException(status_code=404, detail="Enhanced predictive report not found")
        
        return FileResponse(
            path=report_path,
            filename=filename,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    except Exception as e:
        logger.error(f"Download error: {e}")
        raise HTTPException(status_code=500, detail=f"Download failed: {str(e)}")

@router.get("/predictive/download-html/{file_id}/{filename}")
async def download_enhanced_predictive_html_report(
    file_id: str, 
    filename: str,
    output_dir: str = Depends(get_output_dir),
    logger = Depends(get_logger)
):
    """
    Download a generated enhanced predictive HTML report
    """
    try:
        report_path = f"{output_dir}/{file_id}/{filename}"
        
        if not os.path.exists(report_path):
            raise HTTPException(status_code=404, detail="Enhanced predictive HTML report not found")
        
        return FileResponse(
            path=report_path,
            filename=filename,
            media_type="text/html"
        )
    except Exception as e:
        logger.error(f"Download error: {e}")
        raise HTTPException(status_code=500, detail=f"Download failed: {str(e)}")

@router.get("/predictive/download-excel/{file_id}/{filename}")
async def download_enhanced_predictive_excel_report(
    file_id: str, 
    filename: str,
    output_dir: str = Depends(get_output_dir),
    logger = Depends(get_logger)
):
    """
    Download a generated enhanced predictive Excel report
    """
    try:
        report_path = f"{output_dir}/{file_id}/{filename}"
        
        if not os.path.exists(report_path):
            raise HTTPException(status_code=404, detail="Enhanced predictive Excel report not found")
        
        return FileResponse(
            path=report_path,
            filename=filename,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        logger.error(f"Download error: {e}")
        raise HTTPException(status_code=500, detail=f"Download failed: {str(e)}")

@router.get("/predictive/view/{file_id}/{filename}", response_class=HTMLResponse)
async def view_enhanced_predictive_report(
    file_id: str, 
    filename: str,
    output_dir: str = Depends(get_output_dir),
    logger = Depends(get_logger)
):
    """
    View the enhanced predictive HTML report directly in the browser
    """
    try:
        report_path = f"{output_dir}/{file_id}/{filename}"
        
        if not os.path.exists(report_path):
            raise HTTPException(status_code=404, detail="Enhanced predictive HTML report not found")
        
        with open(report_path, 'r', encoding='utf-8') as f:
            html_content = f.read()
            
        return HTMLResponse(content=html_content, status_code=200)
    except Exception as e:
        logger.error(f"View enhanced predictive report error: {e}")
        raise HTTPException(status_code=500, detail=f"Failed to view enhanced predictive report: {str(e)}")

def create_enhanced_predictive_excel_report(df, output_dir, logger):
    """Create an enhanced Excel report with predictive analytics data"""
    try:
        from datetime import datetime
        from sklearn.model_selection import train_test_split
        from sklearn.ensemble import RandomForestClassifier
        from sklearn.preprocessing import LabelEncoder
        
        report_time = datetime.now().strftime("%Y%m%d_%H%M%S")
        excel_path = f"{output_dir}/Enhanced_Predictive_Attrition_Report_{report_time}.xlsx"
        
        # Prepare data for prediction
        df['Is_Separation'] = df['Action Type'].str.contains('Exit|Resignation|Termination|Separation', na=False).astype(int)
        df['Tenure in Years'] = (pd.to_datetime(df['Action Date']) - pd.to_datetime(df['Date of Joining'])).dt.days / 365.25
        df['Tenure in Years'] = df['Tenure in Years'].fillna(0)
        
        features = ['Gender', 'Tenure in Years', 'Function', 'Grade']
        X = df[features].copy()
        y = df['Is_Separation']
        
        # Encode categorical variables
        label_encoders = {}
        for col in X.select_dtypes(include=['object']).columns:
            X[col] = X[col].astype(str).fillna('Unknown')
            le = LabelEncoder()
            X[col] = le.fit_transform(X[col])
            label_encoders[col] = le
        
        X = X.fillna(0)
        
        # Train model
        X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2, random_state=42)
        model = RandomForestClassifier(n_estimators=100, random_state=42)
        model.fit(X_train, y_train)
        
        # Predict for active employees
        active_df = df[~df['Action Type'].str.contains('Exit|Resignation|Termination|Separation', na=False)].copy()
        
        if not active_df.empty:
            active_df['Tenure in Years'] = (pd.to_datetime(active_df['Action Date']) - pd.to_datetime(active_df['Date of Joining'])).dt.days / 365.25
            active_df['Tenure in Years'] = active_df['Tenure in Years'].fillna(0)
            
            X_active = active_df[features].copy()
            for col in X_active.select_dtypes(include=['object']).columns:
                X_active[col] = X_active[col].astype(str).fillna('Unknown')
                if col in label_encoders:
                    le = label_encoders[col]
                    X_active[col] = X_active[col].apply(lambda x: le.transform([x])[0] if x in le.classes_ else 0)
            
            X_active = X_active.fillna(0)
            active_df['Attrition_Probability'] = model.predict_proba(X_active)[:, 1]
            
            # Filter high-risk employees
            high_risk_employees = active_df[active_df['Attrition_Probability'] > 0.7].copy()
            high_risk_employees = high_risk_employees.sort_values(by='Attrition_Probability', ascending=False)
        else:
            high_risk_employees = pd.DataFrame()
        
        # Create Excel file with multiple sheets
        with pd.ExcelWriter(excel_path, engine='xlsxwriter') as writer:
            workbook = writer.book
            
            # Define formats
            header_format = workbook.add_format({
                'bold': True,
                'bg_color': '#4F81BD',
                'font_color': 'white',
                'border': 1,
                'align': 'center',
                'valign': 'vcenter'
            })
            
            high_risk_format = workbook.add_format({
                'bg_color': '#FFE6E6',
                'border': 1,
                'align': 'center'
            })
            
            medium_risk_format = workbook.add_format({
                'bg_color': '#FFF2CC',
                'border': 1,
                'align': 'center'
            })
            
            # Sheet 1: High-Risk Employees
            if not high_risk_employees.empty:
                display_columns = ['Employee Name', 'Gender', 'Function', 'Job Location', 'Grade', 'Attrition_Probability']
                high_risk_display = high_risk_employees[display_columns].copy()
                high_risk_display['Attrition_Probability'] = high_risk_display['Attrition_Probability'].round(3)
                
                high_risk_display.to_excel(writer, sheet_name='High Risk Employees', index=False, startrow=1)
                worksheet = writer.sheets['High Risk Employees']
                
                # Add title
                worksheet.write(0, 0, 'High-Risk Employees (Probability > 0.7)', header_format)
                worksheet.merge_range(0, 0, 0, len(display_columns)-1, 'High-Risk Employees (Probability > 0.7)', header_format)
                
                # Format headers
                for col_num, value in enumerate(high_risk_display.columns):
                    worksheet.write(1, col_num, value, header_format)
                
                # Format data rows based on risk level
                for row_num in range(len(high_risk_display)):
                    probability = high_risk_display.iloc[row_num]['Attrition_Probability']
                    row_format = high_risk_format if probability >= 0.8 else medium_risk_format
                    
                    for col_num in range(len(display_columns)):
                        worksheet.write(row_num + 2, col_num, high_risk_display.iloc[row_num, col_num], row_format)
                
                # Auto-adjust column widths
                for i, col in enumerate(display_columns):
                    max_len = max(len(str(col)), high_risk_display[col].astype(str).str.len().max())
                    worksheet.set_column(i, i, min(max_len + 2, 50))
                
                # Add autofilter
                worksheet.autofilter(1, 0, len(high_risk_display) + 1, len(display_columns) - 1)
            
            # Sheet 2: Summary Statistics
            summary_data = []
            if not high_risk_employees.empty:
                # Gender summary
                gender_summary = high_risk_employees.groupby('Gender').agg({
                    'Employee Name': 'count',
                    'Attrition_Probability': ['mean', 'max']
                }).round(3)
                gender_summary.columns = ['Count', 'Avg_Probability', 'Max_Probability']
                
                for gender, row in gender_summary.iterrows():
                    summary_data.append(['Gender', gender, row['Count'], row['Avg_Probability'], row['Max_Probability']])
                
                # Function summary
                function_summary = high_risk_employees.groupby('Function').agg({
                    'Employee Name': 'count',
                    'Attrition_Probability': ['mean', 'max']
                }).round(3)
                function_summary.columns = ['Count', 'Avg_Probability', 'Max_Probability']
                
                for function, row in function_summary.iterrows():
                    summary_data.append(['Function', function, row['Count'], row['Avg_Probability'], row['Max_Probability']])
                
                # Location summary
                location_summary = high_risk_employees.groupby('Job Location').agg({
                    'Employee Name': 'count',
                    'Attrition_Probability': ['mean', 'max']
                }).round(3)
                location_summary.columns = ['Count', 'Avg_Probability', 'Max_Probability']
                
                for location, row in location_summary.iterrows():
                    summary_data.append(['Location', location, row['Count'], row['Avg_Probability'], row['Max_Probability']])
            
            # Create summary sheet
            summary_df = pd.DataFrame(summary_data, columns=['Category', 'Value', 'Count', 'Avg_Probability', 'Max_Probability'])
            if not summary_df.empty:
                summary_df.to_excel(writer, sheet_name='Summary Statistics', index=False, startrow=1)
                worksheet = writer.sheets['Summary Statistics']
                
                # Add title
                worksheet.write(0, 0, 'High-Risk Employee Summary Statistics', header_format)
                worksheet.merge_range(0, 0, 0, 4, 'High-Risk Employee Summary Statistics', header_format)
                
                # Format headers
                for col_num, value in enumerate(summary_df.columns):
                    worksheet.write(1, col_num, value, header_format)
                
                # Auto-adjust column widths
                for i, col in enumerate(summary_df.columns):
                    max_len = max(len(str(col)), summary_df[col].astype(str).str.len().max() if not summary_df.empty else 10)
                    worksheet.set_column(i, i, min(max_len + 2, 50))
        
        logger.info(f"Enhanced predictive Excel report saved to {excel_path}")
        return excel_path
        
    except Exception as e:
        logger.error(f"Failed to create enhanced predictive Excel report: {e}")
        return None

# @router.get("/health")
# async def health_check_enhanced_predictive():
#     """Health check for enhanced predictive analytics service"""
#     return {"status": "healthy", "service": "enhanced-analytics"}