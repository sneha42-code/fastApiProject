import pandas as pd
from datetime import datetime
import os
import json
import numpy as np
from pathlib import Path
from pydantic import BaseModel

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

def load_data(file_path, logger):
    """
    Load and prepare the data
    """
    try:
        df = pd.read_excel(file_path)
        
        # Handle common column naming variations
        if 'Employee Name' not in df.columns:
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
            result.append({
                'category': str(gender),
                'attritionCount': int(row['Attrition Count']),
                'totalEmployees': int(row['Total Employees']),
                'attritionRate': float(row['Attrition Rate %'])
            })
        
        return result
    except Exception as e:
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
            result.append({
                'category': str(location),
                'attritionCount': int(row['Attrition Count']),
                'totalEmployees': int(row['Total Employees']),
                'attritionRate': float(row['Attrition Rate %'])
            })
        
        return result
    except Exception as e:
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
            result.append({
                'category': str(function),
                'attritionCount': int(row['Attrition Count']),
                'totalEmployees': int(row['Total Employees']),
                'attritionRate': float(row['Attrition Rate %'])
            })
        
        return result
    except Exception as e:
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
            result.append({
                'category': str(row['Tenure Band']),
                'count': int(row['Count']),
                'percentage': float(row['Percentage'])
            })
        
        return result
    except Exception as e:
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
            result.append({
                'category': str(grade),
                'attritionCount': int(row['Attrition Count']),
                'totalEmployees': int(row['Total Employees']),
                'attritionRate': float(row['Attrition Rate %'])
            })
        
        return result
    except Exception as e:
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
            quarterly_result.append({
                'period': str(row['Quarter']),
                'exitCount': int(row['Exit Count'])
            })
        
        monthly_result = []
        for _, row in monthly_trends.iterrows():
            monthly_result.append({
                'period': str(row['Month']),
                'exitCount': int(row['Exit Count'])
            })
        
        return quarterly_result, monthly_result
    except Exception as e:
        return [], []

def extract_slicer_dimensions(df):
    """
    Extract dimensions that can be used as slicers (filters)
    """
    try:
        slicer_data = {}
        
        # Add gender slicer
        if 'Gender' in df.columns:
            gender_values = df['Gender'].dropna().unique().tolist()
            slicer_data['gender'] = [{'value': str(g), 'label': str(g)} for g in gender_values]
        
        # Add location slicer (limit to locations with at least 10 employees)
        if 'Job Location' in df.columns:
            location_counts = df['Job Location'].value_counts()
            valid_locations = location_counts[location_counts >= 10].index.tolist()
            slicer_data['location'] = [{'value': str(loc), 'label': str(loc)} for loc in valid_locations]
        
        # Add function/department slicer
        if 'Function' in df.columns:
            function_counts = df['Function'].value_counts()
            valid_functions = function_counts[function_counts >= 5].index.tolist()
            slicer_data['function'] = [{'value': str(func), 'label': str(func)} for func in valid_functions]
        
        # Add grade slicer
        if 'Grade' in df.columns:
            grade_counts = df['Grade'].value_counts()
            valid_grades = grade_counts[grade_counts >= 5].index.tolist()
            slicer_data['grade'] = [{'value': str(grade), 'label': str(grade)} for grade in valid_grades]
            
        # Add time period slicer if Action Date is available
        if 'Action Date' in df.columns:
            df['Action Date'] = pd.to_datetime(df['Action Date'])
            df['Year'] = df['Action Date'].dt.year
            years = sorted(df['Year'].dropna().unique().tolist())
            slicer_data['year'] = [{'value': str(year), 'label': str(year)} for year in years]
        
        return slicer_data
    except Exception as e:
        return {}

def generate_interactive_html_report(df, output_dir, file_id, logger):
    """
    Generate interactive HTML report with slicers from the analysis data
    """
    try:
        # Create timestamp for the report
        report_time = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        # Clean file_id to avoid any issues
        clean_file_id = str(file_id).strip().replace('"', '').replace("'", '')
        
        # Create output directory for this report
        report_dir = os.path.join(output_dir, clean_file_id)
        os.makedirs(report_dir, exist_ok=True)
        
        logger.info(f"Creating report in directory: {report_dir}")
        
        # Calculate all analyses
        overall_stats = calculate_overall_statistics(df)
        gender_data = calculate_gender_analysis(df)
        location_data = calculate_location_analysis(df)
        function_data = calculate_function_analysis(df)
        tenure_data = calculate_tenure_analysis(df)
        grade_data = calculate_grade_analysis(df)
        quarterly_data, monthly_data = calculate_trend_analysis(df)
        
        # Extract slicer dimensions
        slicer_data = extract_slicer_dimensions(df)
        
        # Prepare data for JavaScript
        report_data = {
            'overall': overall_stats,
            'gender': gender_data,
            'location': location_data,
            'function': function_data,
            'tenure': tenure_data,
            'grade': grade_data,
            'quarterly': quarterly_data,
            'monthly': monthly_data,
            'slicers': slicer_data
        }
        
        # Get HTML template
        html_template = get_interactive_dashboard_template()
        
        # Replace data placeholder with actual data
        final_html = html_template.replace('REPLACE_WITH_DATA', json.dumps(report_data))
        
        # Save HTML file
        report_filename = f"Interactive_Attrition_Dashboard_{report_time}.html"
        html_path = os.path.join(report_dir, report_filename)
        
        with open(html_path, 'w', encoding='utf-8') as f:
            f.write(final_html)
        
        logger.info(f"Interactive HTML dashboard saved to {html_path}")
        
        # Verify file was created
        if os.path.exists(html_path):
            logger.info(f"File verified to exist at: {html_path}")
            return True, html_path, report_filename
        else:
            logger.error(f"File was not created at expected path: {html_path}")
            return False, None, None
            
    except Exception as e:
        logger.error(f"Failed to generate interactive HTML dashboard: {e}")
        return False, None, None

def get_interactive_dashboard_template():
    """
    Return the HTML template for the interactive dashboard
    Note: In production, this should be loaded from a template file
    """
    return '''<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Interactive Attrition Analysis Dashboard</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <link href="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/css/select2.min.css" rel="stylesheet" />
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/select2@4.1.0-rc.0/dist/js/select2.min.js"></script>
    <style>
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background-color: #f8f9fa;
        }
        .dashboard-header {
            background-color: #004C99;
            color: white;
            padding: 20px;
            margin-bottom: 20px;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        }
        .filter-panel {
            background-color: white;
            padding: 15px;
            border-radius: 5px;
            box-shadow: 0 2px 5px rgba(0,0,0,0.1);
            margin-bottom: 20px;
        }
        .section-title {
            color: #004C99;
            font-weight: bold;
            margin-bottom: 20px;
            padding-bottom: 10px;
            border-bottom: 2px solid #4F81BD;
        }
        .card {
            margin-bottom: 20px;
            box-shadow: 0 2px 5px rgba(0,0,0,0.1);
            border: none;
            border-radius: 8px;
        }
        .stat-card {
            background-color: #f8f9fa;
            border-left: 4px solid #004C99;
            transition: transform 0.3s ease;
        }
        .stat-value {
            font-size: 1.8rem;
            font-weight: bold;
            color: #004C99;
        }
        .chart-container {
            position: relative;
            height: 350px;
            margin-bottom: 15px;
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
        <h1>INTERACTIVE ATTRITION ANALYSIS DASHBOARD</h1>
        <h3>BY AUTOMATE REPORTING</h3>
        <p id="reportDate">Report Generated on: </p>
    </div>

    <div class="container-fluid">
        <!-- Filters Panel -->
        <div class="row mb-4">
            <div class="col-12">
                <div class="filter-panel">
                    <h4 class="mb-3">Dashboard Filters</h4>
                    <div class="row">
                        <div class="col-md-3">
                            <label>Gender</label>
                            <select id="genderFilter" class="form-control" multiple="multiple">
                            </select>
                        </div>
                        <div class="col-md-3">
                            <label>Location</label>
                            <select id="locationFilter" class="form-control" multiple="multiple">
                            </select>
                        </div>
                        <div class="col-md-3">
                            <label>Function</label>
                            <select id="functionFilter" class="form-control" multiple="multiple">
                            </select>
                        </div>
                        <div class="col-md-3">
                            <label>Grade</label>
                            <select id="gradeFilter" class="form-control" multiple="multiple">
                            </select>
                        </div>
                    </div>
                    <div class="row mt-2">
                        <div class="col-md-6">
                            <button id="clearAllFilters" class="btn btn-outline-primary btn-sm">Clear All Filters</button>
                        </div>
                        <div class="col-md-6">
                            <p id="filterSummary" class="text-muted mb-0">No filters applied</p>
                        </div>
                    </div>
                </div>
            </div>
        </div>

        <!-- Overall Statistics Section -->
        <div class="row">
            <div class="col-12">
                <h2 class="section-title">Overall Attrition Statistics</h2>
                <div class="row">
                    <div class="col-md-4">
                        <div class="card stat-card p-3">
                            <div>Total Employees</div>
                            <div id="totalEmployees" class="stat-value">-</div>
                        </div>
                    </div>
                    <div class="col-md-4">
                        <div class="card stat-card p-3">
                            <div>Total Exits</div>
                            <div id="totalExits" class="stat-value">-</div>
                        </div>
                    </div>
                    <div class="col-md-4">
                        <div class="card stat-card p-3">
                            <div>Attrition Rate</div>
                            <div id="attritionRate" class="stat-value">-</div>
                        </div>
                    </div>
                </div>
                <div class="row mt-3">
                    <div class="col-md-12">
                        <div class="card">
                            <div class="card-header">Overall Attrition Breakdown</div>
                            <div class="card-body">
                                <div class="chart-container">
                                    <canvas id="overallChart"></canvas>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>

        <!-- Gender Analysis Section -->
        <div class="row mt-4">
            <div class="col-12">
                <h2 class="section-title">Gender-wise Attrition</h2>
                <div class="row">
                    <div class="col-md-5">
                        <div class="card">
                            <div class="card-header">Gender Statistics</div>
                            <div class="card-body">
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
                        </div>
                    </div>
                    <div class="col-md-7">
                        <div class="row">
                            <div class="col-md-6">
                                <div class="card">
                                    <div class="card-header">Attrition Count by Gender</div>
                                    <div class="card-body">
                                        <div class="chart-container">
                                            <canvas id="genderBarChart"></canvas>
                                        </div>
                                    </div>
                                </div>
                            </div>
                            <div class="col-md-6">
                                <div class="card">
                                    <div class="card-header">Attrition Rate by Gender</div>
                                    <div class="card-body">
                                        <div class="chart-container">
                                            <canvas id="genderPieChart"></canvas>
                                        </div>
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>

        <!-- Location Analysis Section -->
        <div class="row mt-4">
            <div class="col-12">
                <h2 class="section-title">Location-wise Attrition</h2>
                <div class="row">
                    <div class="col-md-6">
                        <div class="card">
                            <div class="card-header">Location Statistics</div>
                            <div class="card-body">
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
                        </div>
                    </div>
                    <div class="col-md-6">
                        <div class="card">
                            <div class="card-header">Location Analysis Chart</div>
                            <div class="card-body">
                                <div class="chart-container">
                                    <canvas id="locationChart"></canvas>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>

        <!-- Function Analysis Section -->
        <div class="row mt-4">
            <div class="col-12">
                <h2 class="section-title">Function-wise Attrition</h2>
                <div class="row">
                    <div class="col-md-6">
                        <div class="card">
                            <div class="card-header">Function Statistics</div>
                            <div class="card-body">
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
                        </div>
                    </div>
                    <div class="col-md-6">
                        <div class="card">
                            <div class="card-header">Function Analysis Chart</div>
                            <div class="card-body">
                                <div class="chart-container">
                                    <canvas id="functionChart"></canvas>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>

        <!-- Tenure Analysis Section -->
        <div class="row mt-4">
            <div class="col-12">
                <h2 class="section-title">Tenure Analysis</h2>
                <div class="row">
                    <div class="col-md-6">
                        <div class="card">
                            <div class="card-header">Tenure Statistics</div>
                            <div class="card-body">
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
                        </div>
                    </div>
                    <div class="col-md-6">
                        <div class="card">
                            <div class="card-header">Tenure Analysis Chart</div>
                            <div class="card-body">
                                <div class="chart-container">
                                    <canvas id="tenureChart"></canvas>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>

        <!-- Grade Analysis Section -->
        <div class="row mt-4">
            <div class="col-12">
                <h2 class="section-title">Grade-wise Attrition</h2>
                <div class="row">
                    <div class="col-md-6">
                        <div class="card">
                            <div class="card-header">Grade Statistics</div>
                            <div class="card-body">
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
                        </div>
                    </div>
                    <div class="col-md-6">
                        <div class="card">
                            <div class="card-header">Grade Analysis Chart</div>
                            <div class="card-body">
                                <div class="chart-container">
                                    <canvas id="gradeChart"></canvas>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>

        <!-- Trend Analysis Section -->
        <div class="row mt-4">
            <div class="col-12">
                <h2 class="section-title">Attrition Trends</h2>
                <div class="row">
                    <div class="col-md-6">
                        <div class="card">
                            <div class="card-header">Quarterly Trend</div>
                            <div class="card-body">
                                <div class="chart-container">
                                    <canvas id="quarterlyChart"></canvas>
                                </div>
                            </div>
                        </div>
                    </div>
                    <div class="col-md-6">
                        <div class="card">
                            <div class="card-header">Monthly Trend</div>
                            <div class="card-body">
                                <div class="chart-container">
                                    <canvas id="monthlyChart"></canvas>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </div>

    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
    
    <script>
        // Global variables and settings
        const colors = {
            primary: '#004C99',
            secondary: '#4F81BD',
            accent: '#F79646',
            background: '#EDF3FE'
        };
        
        const chartColors = [
            '#4F81BD', '#C0504D', '#9BBB59', '#8064A2', '#4BACC6', 
            '#F79646', '#6D597A', '#355070', '#6B705C', '#B56576'
        ];
        
        let originalData = null;
        let filteredData = null;
        let charts = {};
        
        const activeFilters = {
            gender: [],
            location: [],
            function: [],
            grade: []
        };
        
        function initializeReport(data) {
            originalData = JSON.parse(JSON.stringify(data));
            filteredData = JSON.parse(JSON.stringify(data));
            
            document.getElementById('reportDate').textContent += new Date().toLocaleString();
            
            $('#genderFilter, #locationFilter, #functionFilter, #gradeFilter').select2({
                placeholder: "Select options",
                allowClear: true,
                width: '100%'
            });
            
            populateFilterOptions(data);
            setupFilterEventHandlers();
            updateDashboard(filteredData);
        }
        
        function populateFilterOptions(data) {
            const genderFilter = document.getElementById('genderFilter');
            if (data.slicers.gender) {
                data.slicers.gender.forEach(item => {
                    const option = new Option(item.label, item.value);
                    genderFilter.appendChild(option);
                });
            }
            
            const locationFilter = document.getElementById('locationFilter');
            if (data.slicers.location) {
                data.slicers.location.forEach(item => {
                    const option = new Option(item.label, item.value);
                    locationFilter.appendChild(option);
                });
            }
            
            const functionFilter = document.getElementById('functionFilter');
            if (data.slicers.function) {
                data.slicers.function.forEach(item => {
                    const option = new Option(item.label, item.value);
                    functionFilter.appendChild(option);
                });
            }
            
            const gradeFilter = document.getElementById('gradeFilter');
            if (data.slicers.grade) {
                data.slicers.grade.forEach(item => {
                    const option = new Option(item.label, item.value);
                    gradeFilter.appendChild(option);
                });
            }
        }
        
        function setupFilterEventHandlers() {
            $('#genderFilter').on('change', function() {
                activeFilters.gender = $(this).val() || [];
                applyFilters();
            });
            
            $('#locationFilter').on('change', function() {
                activeFilters.location = $(this).val() || [];
                applyFilters();
            });
            
            $('#functionFilter').on('change', function() {
                activeFilters.function = $(this).val() || [];
                applyFilters();
            });
            
            $('#gradeFilter').on('change', function() {
                activeFilters.grade = $(this).val() || [];
                applyFilters();
            });
            
            $('#clearAllFilters').on('click', function() {
                $('#genderFilter, #locationFilter, #functionFilter, #gradeFilter').val(null).trigger('change');
                
                for (const key in activeFilters) {
                    activeFilters[key] = [];
                }
                
                filteredData = JSON.parse(JSON.stringify(originalData));
                updateDashboard(filteredData);
                updateFilterSummary();
            });
        }
        
        function applyFilters() {
            filteredData = JSON.parse(JSON.stringify(originalData));
            
            if (activeFilters.gender.length > 0) {
                filteredData.gender = originalData.gender.filter(row => 
                    activeFilters.gender.includes(row.category)
                );
            }
            
            if (activeFilters.location.length > 0) {
                filteredData.location = originalData.location.filter(row => 
                    activeFilters.location.includes(row.category)
                );
            }
            
            if (activeFilters.function.length > 0) {
                filteredData.function = originalData.function.filter(row => 
                    activeFilters.function.includes(row.category)
                );
            }
            
            if (activeFilters.grade.length > 0) {
                filteredData.grade = originalData.grade.filter(row => 
                    activeFilters.grade.includes(row.category)
                );
            }
            
            updateOverallStatistics();
            updateFilterSummary();
            updateDashboard(filteredData);
        }
        
        function updateOverallStatistics() {
            let totalEmployees = 0;
            let totalExits = 0;
            
            filteredData.gender.forEach(row => {
                totalEmployees += row.totalEmployees;
                totalExits += row.attritionCount;
            });
            
            filteredData.overall = {
                totalEmployees: totalEmployees,
                totalExits: totalExits,
                attritionRate: totalEmployees > 0 ? ((totalExits / totalEmployees) * 100).toFixed(2) : 0
            };
        }
        
        function updateFilterSummary() {
            const filterSummary = document.getElementById('filterSummary');
            const totalFilters = Object.values(activeFilters).reduce(
                (sum, filters) => sum + filters.length, 0
            );
            
            if (totalFilters === 0) {
                filterSummary.textContent = "No filters applied";
            } else {
                filterSummary.textContent = `${totalFilters} filter(s) applied`;
            }
        }
        
        function updateDashboard(data) {
            updateOverallSection(data.overall);
            updateGenderSection(data.gender);
            updateLocationSection(data.location);
            updateFunctionSection(data.function);
            updateTenureSection(data.tenure);
            updateGradeSection(data.grade);
            updateTrendSection(data.quarterly, data.monthly);
        }
        
        function updateOverallSection(overall) {
            document.getElementById('totalEmployees').textContent = overall.totalEmployees;
            document.getElementById('totalExits').textContent = overall.totalExits;
            document.getElementById('attritionRate').textContent = `${overall.attritionRate}%`;
            
            if (charts.overall) {
                charts.overall.data.datasets[0].data = [
                    overall.totalEmployees - overall.totalExits,
                    overall.totalExits
                ];
                charts.overall.update();
            } else {
                const ctx = document.getElementById('overallChart').getContext('2d');
                charts.overall = new Chart(ctx, {
                    type: 'pie',
                    data: {
                        labels: ['Active Employees', 'Exited Employees'],
                        datasets: [{
                            data: [overall.totalEmployees - overall.totalExits, overall.totalExits],
                            backgroundColor: [colors.secondary, colors.accent]
                        }]
                    },
                    options: {
                        responsive: true,
                        maintainAspectRatio: false,
                        plugins: { legend: { position: 'bottom' } }
                    }
                });
            }
        }
        
        function updateGenderSection(genderData) {
            populateTable('genderTable', genderData);
            updateChart('genderBarChart', genderData, 'bar', 'Attrition Count by Gender');
            updateChart('genderPieChart', genderData, 'pie', 'Attrition Rate by Gender');
        }
        
        function updateLocationSection(locationData) {
            populateTable('locationTable', locationData);
            updateChart('locationChart', locationData, 'bar', 'Attrition by Location');
        }
        
        function updateFunctionSection(functionData) {
            populateTable('functionTable', functionData);
            updateChart('functionChart', functionData, 'bar', 'Attrition by Function');
        }
        
        function updateTenureSection(tenureData) {
            populateTable('tenureTable', tenureData);
            updateChart('tenureChart', tenureData, 'bar', 'Attrition by Tenure');
        }
        
        function updateGradeSection(gradeData) {
            populateTable('gradeTable', gradeData);
            updateChart('gradeChart', gradeData, 'bar', 'Attrition by Grade');
        }
        
        function updateTrendSection(quarterlyData, monthlyData) {
            updateChart('quarterlyChart', quarterlyData, 'bar', 'Quarterly Trend', 'period', 'exitCount');
            updateChart('monthlyChart', monthlyData, 'line', 'Monthly Trend', 'period', 'exitCount');
        }
        
        function populateTable(tableId, data) {
            const tbody = document.querySelector(`#${tableId} tbody`);
            tbody.innerHTML = '';
            
            data.forEach(row => {
                const tr = document.createElement('tr');
                if (tableId === 'tenureTable') {
                    tr.innerHTML = `<td>${row.category}</td><td>${row.count}</td><td>${row.percentage}%</td>`;
                } else {
                    tr.innerHTML = `<td>${row.category}</td><td>${row.attritionCount}</td><td>${row.totalEmployees}</td><td>${row.attritionRate}%</td>`;
                }
                tbody.appendChild(tr);
            });
        }
        
        function updateChart(chartId, data, type, title, labelKey = 'category', dataKey = 'attritionCount') {
            const ctx = document.getElementById(chartId);
            if (!ctx) return;
            
            if (charts[chartId]) {
                charts[chartId].destroy();
            }
            
            const chartCtx = ctx.getContext('2d');
            charts[chartId] = new Chart(chartCtx, {
                type: type,
                data: {
                    labels: data.map(row => row[labelKey]),
                    datasets: [{
                        label: title,
                        data: data.map(row => row[dataKey]),
                        backgroundColor: type === 'pie' ? chartColors : colors.secondary,
                        borderColor: type === 'line' ? colors.accent : undefined,
                        borderWidth: type === 'line' ? 2 : undefined,
                        tension: type === 'line' ? 0.2 : undefined
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {
                        title: { display: true, text: title },
                        legend: { display: type === 'pie' }
                    },
                    scales: type !== 'pie' ? {
                        y: { beginAtZero: true }
                    } : undefined
                }
            });
        }
        
        // Initialize dashboard
        const reportData = REPLACE_WITH_DATA;
        document.addEventListener('DOMContentLoaded', function() {
            initializeReport(reportData);
        });
    </script>
</body>
</html>'''