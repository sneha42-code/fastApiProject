import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from datetime import datetime
import os
import json
import logging
import numpy as np
from pathlib import Path
import uuid
import shutil
from typing import List, Optional
from fastapi import FastAPI, File, UploadFile, HTTPException, BackgroundTasks, Form, Response
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel

# Set up logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("attrition_report.log"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# Create directories
UPLOAD_DIR = "uploads"
OUTPUT_DIR = "attrition_reports"
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

# Create FastAPI app
app = FastAPI(
    title="Attrition Analysis API",
    description="API for generating automated interactive attrition analysis reports",
    version="1.0.0"
)

# Add CORS middleware
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # Allow all origins
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Response models
class UploadResponse(BaseModel):
    file_id: str
    filename: str
    message: str

class ReportResponse(BaseModel):
    status: str
    message: str
    file_id: str
    report_file: str
    download_url: str

# Data analysis functions
def load_data(file_path):
    """
    Load and prepare the data
    """
    try:
        df = pd.read_excel(file_path)
        
        # Handle common column naming variations
        if 'Employee Name' not in df.columns:
            # Look for alternative ID columns
            id_columns = ['Employee ID', 'EmployeeID', 'Emp ID', 'EmpID', 'ID', 'Name']
            for col in id_columns:
                if col in df.columns:
                    df['Employee Name'] = df[col]
                    break
            
        # Ensure Action Type column exists
        if 'Action Type' not in df.columns:
            status_columns = ['Status', 'Employee Status', 'Employment Status', 
                             'Job Status', 'Action', 'Termination Status', 'Exit Status']
            for col in status_columns:
                if col in df.columns:
                    df['Action Type'] = df[col]
                    break
                    
        # Ensure Gender column exists
        if 'Gender' not in df.columns and 'Sex' in df.columns:
            df['Gender'] = df['Sex']
            
        # Ensure Function column exists
        if 'Function' not in df.columns:
            if 'Department' in df.columns:
                df['Function'] = df['Department']
            elif 'Business Unit' in df.columns:
                df['Function'] = df['Business Unit']
                
        # Ensure Grade column exists
        if 'Grade' not in df.columns:
            grade_alternatives = [
                "Job Grade", "Pay Grade", "Grade Level", "Salary Grade", 
                "Grade Code", "Band", "Position Grade"
            ]
            for alt_column in grade_alternatives:
                if alt_column in df.columns:
                    df['Grade'] = df[alt_column]
                    break
                    
        df['Action Type'] = df['Action Type'].fillna('')
        logger.info(f"Successfully loaded data from {file_path}")
        return df
    except Exception as e:
        logger.error(f"Error loading data: {e}")
        return None

def calculate_overall_statistics(df):
    """
    Calculate overall attrition statistics
    """
    try:
        total_employees = len(df)
        exits = df[df['Action Type'].str.contains('Exit|Resignation|Termination', na=False)]
        total_exits = len(exits)
        attrition_rate = round((total_exits / total_employees * 100), 2)
        
        return {
            'totalEmployees': total_employees,
            'totalExits': total_exits,
            'attritionRate': attrition_rate
        }
    except Exception as e:
        logger.error(f"Error calculating overall statistics: {e}")
        return None

def calculate_gender_analysis(df):
    """
    Calculate gender-wise attrition analysis
    """
    try:
        attrition_by_gender = df[df['Action Type'].str.contains('Exit|Resignation|Termination', na=False)].groupby('Gender').agg({
            'Employee Name': 'count'
        }).rename(columns={'Employee Name': 'Attrition Count'})
        
        total_by_gender = df.groupby('Gender')['Employee Name'].count()
        attrition_by_gender['Total Employees'] = total_by_gender
        attrition_by_gender['Attrition Rate %'] = round((attrition_by_gender['Attrition Count'] / attrition_by_gender['Total Employees'] * 100), 2)
        
        # Convert to list format for HTML
        result = []
        for gender, row in attrition_by_gender.iterrows():
            result.append([
                str(gender),
                int(row['Attrition Count']),
                int(row['Total Employees']),
                float(row['Attrition Rate %'])
            ])
        
        return result
    except Exception as e:
        logger.error(f"Error in gender analysis: {e}")
        return []

def calculate_location_analysis(df):
    """
    Calculate location-wise attrition analysis
    """
    try:
        total_by_location = df.groupby('Job Location')['Employee Name'].count()
        valid_locations = total_by_location[total_by_location >= 50].index
        
        attrition_by_location = df[
            (df['Action Type'].str.contains('Exit|Resignation|Termination', na=False)) & 
            (df['Job Location'].isin(valid_locations))
        ].groupby('Job Location').agg({
            'Employee Name': 'count'
        }).rename(columns={'Employee Name': 'Attrition Count'})
        
        total_by_location = total_by_location[valid_locations]
        attrition_by_location['Total Employees'] = total_by_location
        attrition_by_location['Attrition Rate %'] = round((attrition_by_location['Attrition Count'] / attrition_by_location['Total Employees'] * 100), 2)
        
        # Convert to list format for HTML
        result = []
        for location, row in attrition_by_location.iterrows():
            result.append([
                str(location),
                int(row['Attrition Count']),
                int(row['Total Employees']),
                float(row['Attrition Rate %'])
            ])
        
        return result
    except Exception as e:
        logger.error(f"Error in location analysis: {e}")
        return []

def calculate_function_analysis(df):
    """
    Calculate function-wise attrition analysis
    """
    try:
        total_by_function = df.groupby('Function')['Employee Name'].count()
        valid_functions = total_by_function[total_by_function >= 20].index
        
        attrition_by_function = df[
            (df['Action Type'].str.contains('Exit|Resignation|Termination', na=False)) & 
            (df['Function'].isin(valid_functions))
        ].groupby('Function').agg({
            'Employee Name': 'count'
        }).rename(columns={'Employee Name': 'Attrition Count'})
        
        total_by_function = total_by_function[valid_functions]
        attrition_by_function['Total Employees'] = total_by_function
        attrition_by_function['Attrition Rate %'] = round((attrition_by_function['Attrition Count'] / attrition_by_function['Total Employees'] * 100), 2)
        
        # Convert to list format for HTML
        result = []
        for function, row in attrition_by_function.iterrows():
            result.append([
                str(function),
                int(row['Attrition Count']),
                int(row['Total Employees']),
                float(row['Attrition Rate %'])
            ])
        
        return result
    except Exception as e:
        logger.error(f"Error in function analysis: {e}")
        return []

def calculate_tenure_analysis(df):
    """
    Calculate tenure analysis of exited employees
    """
    try:
        # Filter to only include employees who have exited
        exits_df = df[df['Action Type'].str.contains('Exit|Resignation|Termination', na=False)].copy()

        # Calculate tenure in years
        exits_df['Tenure in Days'] = (pd.to_datetime(exits_df['Action Date']) - pd.to_datetime(exits_df['Date of Joining'])).dt.days
        exits_df['Tenure in Years'] = exits_df['Tenure in Days'] / 365.25

        # Create tenure bins
        tenure_bins = [0, 1, 2, 3, 5, 10, float('inf')]
        tenure_labels = ['<1 year', '1-2 years', '2-3 years', '3-5 years', '5-10 years', '>10 years']
        exits_df['Tenure Band'] = pd.cut(exits_df['Tenure in Years'], bins=tenure_bins, labels=tenure_labels, right=False)

        # Group by tenure band
        tenure_analysis = exits_df.groupby('Tenure Band', observed=True).size().reset_index(name='Count')
        tenure_analysis['Percentage'] = round((tenure_analysis['Count'] / tenure_analysis['Count'].sum() * 100), 2)
        
        # Convert to list format for HTML
        result = []
        for _, row in tenure_analysis.iterrows():
            result.append([
                str(row['Tenure Band']),
                int(row['Count']),
                float(row['Percentage'])
            ])
        
        return result
    except Exception as e:
        logger.error(f"Error in tenure analysis: {e}")
        return []

def calculate_grade_analysis(df):
    """
    Calculate grade-wise attrition analysis
    """
    try:
        total_by_grade = df.groupby('Grade')['Employee Name'].count()
        valid_grades = total_by_grade[total_by_grade >= 10].index
        
        attrition_by_grade = df[
            (df['Action Type'].str.contains('Exit|Resignation|Termination', na=False)) & 
            (df['Grade'].isin(valid_grades))
        ].groupby('Grade').agg({
            'Employee Name': 'count'
        }).rename(columns={'Employee Name': 'Attrition Count'})
        
        total_by_grade = total_by_grade[valid_grades]
        attrition_by_grade['Total Employees'] = total_by_grade
        attrition_by_grade['Attrition Rate %'] = round((attrition_by_grade['Attrition Count'] / attrition_by_grade['Total Employees'] * 100), 2)
        
        # Convert to list format for HTML
        result = []
        for grade, row in attrition_by_grade.iterrows():
            result.append([
                str(grade),
                int(row['Attrition Count']),
                int(row['Total Employees']),
                float(row['Attrition Rate %'])
            ])
        
        return result
    except Exception as e:
        logger.error(f"Error in grade analysis: {e}")
        return []

def calculate_trend_analysis(df):
    """
    Calculate quarterly and monthly attrition trends
    """
    try:
        # Filter to include only exit actions
        exits_df = df[df['Action Type'].str.contains('Exit|Resignation|Termination', na=False)].copy()

        # Ensure dates are in datetime format
        exits_df['Action Date'] = pd.to_datetime(exits_df['Action Date'])

        # Add month and quarter columns
        exits_df['Month'] = exits_df['Action Date'].dt.strftime('%Y-%m')
        exits_df['Quarter'] = exits_df['Action Date'].dt.to_period('Q').astype(str)

        # Calculate quarterly trends
        quarterly_trends = exits_df.groupby('Quarter').size().reset_index(name='Exit Count')
        quarterly_trends = quarterly_trends.sort_values('Quarter')
        
        # Calculate monthly trends
        monthly_trends = exits_df.groupby('Month').size().reset_index(name='Exit Count')
        monthly_trends = monthly_trends.sort_values('Month')
        
        # Convert to list format for HTML
        quarterly_result = []
        for _, row in quarterly_trends.iterrows():
            quarterly_result.append([
                str(row['Quarter']),
                int(row['Exit Count'])
            ])
        
        monthly_result = []
        for _, row in monthly_trends.iterrows():
            monthly_result.append([
                str(row['Month']),
                int(row['Exit Count'])
            ])
        
        return quarterly_result, monthly_result
    except Exception as e:
        logger.error(f"Error in trend analysis: {e}")
        return [], []

def generate_html_report(df, output_dir, file_id):
    """
    Generate HTML report from the analysis data
    """
    try:
        # Create timestamp for the report
        report_time = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        # Create output directory for this report
        report_dir = f"{output_dir}/{file_id}"
        os.makedirs(report_dir, exist_ok=True)
        
        # Calculate all analyses
        overall_stats = calculate_overall_statistics(df)
        gender_data = calculate_gender_analysis(df)
        location_data = calculate_location_analysis(df)
        function_data = calculate_function_analysis(df)
        tenure_data = calculate_tenure_analysis(df)
        grade_data = calculate_grade_analysis(df)
        quarterly_data, monthly_data = calculate_trend_analysis(df)
        
        # Prepare data for JavaScript
        report_data = {
            'overall': overall_stats,
            'gender': gender_data,
            'location': location_data,
            'function': function_data,
            'tenure': tenure_data,
            'grade': grade_data,
            'quarterly': quarterly_data,
            'monthly': monthly_data
        }
        
        # Read HTML template
        html_template = '''<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Attrition Analysis Report</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <style>
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background-color: #f8f9fa;
        }
        .dashboard-header {
            background-color: #004C99;
            color: white;
            padding: 20px;
            margin-bottom: 30px;
        }
        .section-title {
            color: #004C99;
            font-weight: bold;
            margin-bottom: 20px;
        }
        .card {
            margin-bottom: 20px;
            box-shadow: 0 2px 5px rgba(0,0,0,0.1);
        }
        .stat-card {
            background-color: #f8f9fa;
            border-left: 4px solid #004C99;
        }
        .chart-container {
            position: relative;
            height: 400px;
        }
        .table th {
            background-color: #4F81BD;
            color: white;
        }
        .table-striped tbody tr:nth-of-type(odd) {
            background-color: #EDF3FE;
        }
    </style>
</head>
<body>
    <div class="dashboard-header text-center">
        <h1>AUTOMATED ATTRITION ANALYSIS REPORT</h1>
        <h3>BY VIMAL SINGH</h3>
        <p id="reportDate">Report Generated on: </p>
    </div>

    <div class="container-fluid">
        <div class="row">
            <!-- Overall Statistics Section -->
            <div class="col-12">
                <h2 class="section-title">1. Overall Attrition Statistics</h2>
                <div class="row">
                    <div class="col-md-6">
                        <div class="card stat-card p-3">
                            <h5>Total Employees: <span id="totalEmployees">-</span></h5>
                            <h5>Total Exits: <span id="totalExits">-</span></h5>
                            <h5>Overall Attrition Rate: <span id="attritionRate">-</span>%</h5>
                        </div>
                    </div>
                    <div class="col-md-6">
                        <div class="chart-container">
                            <canvas id="overallChart"></canvas>
                        </div>
                    </div>
                </div>
            </div>

            <!-- Gender Analysis Section -->
            <div class="col-12 mt-5">
                <h2 class="section-title">2. Gender-wise Attrition</h2>
                <div class="row">
                    <div class="col-md-4">
                        <table class="table table-striped" id="genderTable">
                            <thead>
                                <tr>
                                    <th>Gender</th>
                                    <th>Attrition Count</th>
                                    <th>Total Employees</th>
                                    <th>Attrition Rate %</th>
                                </tr>
                            </thead>
                            <tbody></tbody>
                        </table>
                    </div>
                    <div class="col-md-4">
                        <div class="chart-container">
                            <canvas id="genderBarChart"></canvas>
                        </div>
                    </div>
                    <div class="col-md-4">
                        <div class="chart-container">
                            <canvas id="genderPieChart"></canvas>
                        </div>
                    </div>
                </div>
            </div>

            <!-- Location Analysis Section -->
            <div class="col-12 mt-5">
                <h2 class="section-title">3. Location-wise Attrition (≥50 Employees)</h2>
                <div class="row">
                    <div class="col-md-6">
                        <table class="table table-striped" id="locationTable">
                            <thead>
                                <tr>
                                    <th>Location</th>
                                    <th>Attrition Count</th>
                                    <th>Total Employees</th>
                                    <th>Attrition Rate %</th>
                                </tr>
                            </thead>
                            <tbody></tbody>
                        </table>
                    </div>
                    <div class="col-md-6">
                        <div class="chart-container">
                            <canvas id="locationChart"></canvas>
                        </div>
                    </div>
                </div>
            </div>

            <!-- Function Analysis Section -->
            <div class="col-12 mt-5">
                <h2 class="section-title">4. Function-wise Attrition (≥20 Employees)</h2>
                <div class="row">
                    <div class="col-md-6">
                        <table class="table table-striped" id="functionTable">
                            <thead>
                                <tr>
                                    <th>Function</th>
                                    <th>Attrition Count</th>
                                    <th>Total Employees</th>
                                    <th>Attrition Rate %</th>
                                </tr>
                            </thead>
                            <tbody></tbody>
                        </table>
                    </div>
                    <div class="col-md-6">
                        <div class="chart-container">
                            <canvas id="functionChart"></canvas>
                        </div>
                    </div>
                </div>
            </div>

            <!-- Tenure Analysis Section -->
            <div class="col-12 mt-5">
                <h2 class="section-title">5. Tenure Analysis of Exited Employees</h2>
                <div class="row">
                    <div class="col-md-6">
                        <table class="table table-striped" id="tenureTable">
                            <thead>
                                <tr>
                                    <th>Tenure Band</th>
                                    <th>Number of Exits</th>
                                    <th>Percentage</th>
                                </tr>
                            </thead>
                            <tbody></tbody>
                        </table>
                    </div>
                    <div class="col-md-6">
                        <div class="chart-container">
                            <canvas id="tenureChart"></canvas>
                        </div>
                    </div>
                </div>
            </div>

            <!-- Grade Analysis Section -->
            <div class="col-12 mt-5">
                <h2 class="section-title">6. Grade-wise Attrition</h2>
                <div class="row">
                    <div class="col-md-6">
                        <table class="table table-striped" id="gradeTable">
                            <thead>
                                <tr>
                                    <th>Grade</th>
                                    <th>Attrition Count</th>
                                    <th>Total Employees</th>
                                    <th>Attrition Rate %</th>
                                </tr>
                            </thead>
                            <tbody></tbody>
                        </table>
                    </div>
                    <div class="col-md-6">
                        <div class="chart-container">
                            <canvas id="gradeChart"></canvas>
                        </div>
                    </div>
                </div>
            </div>

            <!-- Trend Analysis Section -->
            <div class="col-12 mt-5">
                <h2 class="section-title">7. Quarterly and Monthly Attrition Trends</h2>
                <div class="row">
                    <div class="col-md-6">
                        <h4>Quarterly Trend</h4>
                        <div class="chart-container">
                            <canvas id="quarterlyChart"></canvas>
                        </div>
                    </div>
                    <div class="col-md-6">
                        <h4>Monthly Trend</h4>
                        <div class="chart-container">
                            <canvas id="monthlyChart"></canvas>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </div>

    <script>
        // Color scheme
        const colors = {
            primary: '#004C99',
            secondary: '#4F81BD',
            accent: '#F79646',
            background: '#EDF3FE'
        };

        // Initialize report with data (this will be populated by Python)
        function initializeReport(data) {
            // Set report date
            document.getElementById('reportDate').textContent += new Date().toLocaleString();

            // Overall Statistics
            document.getElementById('totalEmployees').textContent = data.overall.totalEmployees;
            document.getElementById('totalExits').textContent = data.overall.totalExits;
            document.getElementById('attritionRate').textContent = data.overall.attritionRate;

            // Create overall chart
            createOverallChart(data.overall);

            // Gender Analysis
            populateTable('genderTable', data.gender);
            createGenderCharts(data.gender);

            // Location Analysis
            populateTable('locationTable', data.location);
            createBarChart('locationChart', data.location, 'Location', 'Attrition Rate Distribution by Location');

            // Function Analysis
            populateTable('functionTable', data.function);
            createBarChart('functionChart', data.function, 'Function', 'Attrition Rate Distribution by Function');

            // Tenure Analysis
            populateTable('tenureTable', data.tenure);
            createBarChart('tenureChart', data.tenure, 'Tenure Band', 'Attrition Distribution by Tenure');

            // Grade Analysis
            populateTable('gradeTable', data.grade);
            createBarChart('gradeChart', data.grade, 'Grade', 'Attrition Rate Distribution by Grade');

            // Trend Analysis
            createQuarterlyChart(data.quarterly);
            createMonthlyChart(data.monthly);
        }

        function createOverallChart(data) {
            const ctx = document.getElementById('overallChart').getContext('2d');
            new Chart(ctx, {
                type: 'pie',
                data: {
                    labels: ['Active Employees', 'Exited Employees'],
                    datasets: [{
                        data: [data.totalEmployees - data.totalExits, data.totalExits],
                        backgroundColor: [colors.secondary, colors.accent]
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {
                        title: {
                            display: true,
                            text: 'Overall Attrition'
                        },
                        legend: {
                            position: 'bottom'
                        }
                    }
                }
            });
        }

        function populateTable(tableId, data) {
            const tbody = document.querySelector(`#${tableId} tbody`);
            tbody.innerHTML = '';
            data.forEach(row => {
                const tr = document.createElement('tr');
                Object.values(row).forEach(value => {
                    const td = document.createElement('td');
                    td.textContent = value;
                    tr.appendChild(td);
                });
                tbody.appendChild(tr);
            });
        }

        function createGenderCharts(data) {
            // Bar chart
            const barCtx = document.getElementById('genderBarChart').getContext('2d');
            new Chart(barCtx, {
                type: 'bar',
                data: {
                    labels: data.map(row => row[0]),
                    datasets: [{
                        label: 'Attrition Count',
                        data: data.map(row => row[1]),
                        backgroundColor: colors.secondary
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {
                        title: {
                            display: true,
                            text: 'Attrition Count by Gender'
                        }
                    }
                }
            });

            // Pie chart
            const pieCtx = document.getElementById('genderPieChart').getContext('2d');
            new Chart(pieCtx, {
                type: 'pie',
                data: {
                    labels: data.map(row => row[0]),
                    datasets: [{
                        data: data.map(row => row[3]),
                        backgroundColor: [colors.secondary, colors.accent]
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {
                        title: {
                            display: true,
                            text: 'Attrition Rate Distribution by Gender'
                        }
                    }
                }
            });
        }

        function createBarChart(chartId, data, label, title) {
            const ctx = document.getElementById(chartId).getContext('2d');
            new Chart(ctx, {
                type: 'bar',
                data: {
                    labels: data.map(row => row[0]),
                    datasets: [{
                        label: 'Attrition Rate %',
                        data: data.map(row => row[row.length - 1]),
                        backgroundColor: colors.secondary
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {
                        title: {
                            display: true,
                            text: title
                        }
                    },
                    scales: {
                        y: {
                            beginAtZero: true
                        }
                    }
                }
            });
        }

        function createQuarterlyChart(data) {
            const ctx = document.getElementById('quarterlyChart').getContext('2d');
            new Chart(ctx, {
                type: 'bar',
                data: {
                    labels: data.map(row => row[0]),
                    datasets: [{
                        label: 'Exit Count',
                        data: data.map(row => row[1]),
                        backgroundColor: colors.secondary
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {
                        title: {
                            display: true,
                            text: 'Quarterly Attrition Trend'
                        }
                    }
                }
            });
        }

        function createMonthlyChart(data) {
            const ctx = document.getElementById('monthlyChart').getContext('2d');
            new Chart(ctx, {
                type: 'line',
                data: {
                    labels: data.map(row => row[0]),
                    datasets: [{
                        label: 'Exit Count',
                        data: data.map(row => row[1]),
                        borderColor: colors.secondary,
                        tension: 0.4
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {
                        title: {
                            display: true,
                            text: 'Monthly Attrition Trend'
                        }
                    }
                }
            });
        }

        // Data will be injected here by Python
        const reportData = REPLACE_WITH_DATA;
        initializeReport(reportData);
    </script>
</body>
</html>'''
        
        # Replace data placeholder with actual data
        final_html = html_template.replace('REPLACE_WITH_DATA', json.dumps(report_data))
        
        # Save HTML file
        report_filename = f"Attrition_Report_{report_time}.html"
        html_path = f"{report_dir}/{report_filename}"
        with open(html_path, 'w', encoding='utf-8') as f:
            f.write(final_html)
        
        logger.info(f"HTML report saved to {html_path}")
        return True, html_path, report_filename
        
    except Exception as e:
        logger.error(f"Failed to generate HTML report: {e}")
        return False, None, None

# API Endpoints
@app.get("/")
async def read_root():
    """
    Root endpoint to check if the API is working
    """
    return {
        "message": "Welcome to the Attrition Analysis API",
        "version": "1.0.0",
        "endpoints": {
            "/upload": "Upload HRIS Excel file for analysis",
            "/generate-report/{file_id}": "Generate HTML attrition report from uploaded file",
            "/download/{file_id}/{filename}": "Download generated report",
            "/view/{file_id}/{filename}": "View interactive HTML report in browser"
        }
    }

@app.post("/upload/", response_model=UploadResponse)
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

@app.post("/generate-report/{file_id}", response_model=ReportResponse)
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

@app.get("/download/{file_id}/{filename}")
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

@app.get("/view/{file_id}/{filename}", response_class=HTMLResponse)
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
app.mount("/reports", StaticFiles(directory=OUTPUT_DIR), name="reports")

# Main function to run the application
if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)