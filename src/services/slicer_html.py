import pandas as pd
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
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel

# Placeholder for setup_logging (replace with actual implementation if available)
def setup_logging():
    logging.basicConfig(level=logging.INFO)
    return logging.getLogger(__name__)

# Initialize logging
logger = setup_logging()

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
            result.append({
                'category': str(gender),
                'attritionCount': int(row['Attrition Count']),
                'totalEmployees': int(row['Total Employees']),
                'attritionRate': float(row['Attrition Rate %'])
            })
        
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
            result.append({
                'category': str(location),
                'attritionCount': int(row['Attrition Count']),
                'totalEmployees': int(row['Total Employees']),
                'attritionRate': float(row['Attrition Rate %'])
            })
        
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
            result.append({
                'category': str(function),
                'attritionCount': int(row['Attrition Count']),
                'totalEmployees': int(row['Total Employees']),
                'attritionRate': float(row['Attrition Rate %'])
            })
        
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
            result.append({
                'category': str(row['Tenure Band']),
                'count': int(row['Count']),
                'percentage': float(row['Percentage'])
            })
        
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
            result.append({
                'category': str(grade),
                'attritionCount': int(row['Attrition Count']),
                'totalEmployees': int(row['Total Employees']),
                'attritionRate': float(row['Attrition Rate %'])
            })
        
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
        logger.error(f"Error in trend analysis: {e}")
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
        logger.error(f"Error extracting slicer dimensions: {e}")
        return {}

def generate_interactive_html_report(df, output_dir, file_id):
    """
    Generate interactive HTML report with slicers from the analysis data
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
        
        # HTML template
        html_template = '''<!DOCTYPE html>
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
        .filter-title {
            color: #004C99;
            font-weight: bold;
            margin-bottom: 15px;
            border-bottom: 1px solid #eee;
            padding-bottom: 10px;
        }
        .filter-section {
            margin-bottom: 15px;
        }
        .filter-label {
            font-weight: 600;
            margin-bottom: 5px;
            color: #495057;
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
        .card-header {
            background-color: #EDF3FE;
            border-bottom: 1px solid rgba(0,0,0,0.125);
            font-weight: bold;
            color: #004C99;
            border-radius: 8px 8px 0 0 !important;
        }
        .stat-card {
            background-color: #f8f9fa;
            border-left: 4px solid #004C99;
            transition: transform 0.3s ease;
        }
        .stat-card:hover {
            transform: translateY(-5px);
        }
        .stat-value {
            font-size: 1.8rem;
            font-weight: bold;
            color: #004C99;
        }
        .stat-label {
            color: #6c757d;
            font-size: 0.9rem;
        }
        .chart-container {
            position: relative;
            height: 350px;
            margin-bottom: 15px;
        }
        .table th {
            background-color: #4F81BD;
            color: white;
            position: sticky;
            top: 0;
        }
        .table-striped tbody tr:nth-of-type(odd) {
            background-color: #EDF3FE;
        }
        .table-responsive {
            max-height: 350px;
            overflow-y: auto;
        }
        .select2-container--default .select2-selection--multiple {
            border-color: #ced4da;
        }
        .select2-container--default.select2-container--focus .select2-selection--multiple {
            border-color: #004C99;
        }
        .select2-container--default .select2-results__option--highlighted[aria-selected] {
            background-color: #004C99;
        }
        .btn-primary {
            background-color: #004C99;
            border-color: #004C99;
        }
        .btn-primary:hover {
            background-color: #4F81BD;
            border-color: #4F81BD;
        }
        .btn-outline-primary {
            color: #004C99;
            border-color: #004C99;
        }
        .btn-outline-primary:hover {
            background-color: #004C99;
            color: white;
        }
        .animate-on-scroll {
            opacity: 0;
            transform: translateY(20px);
            transition: opacity 0.6s ease, transform 0.6s ease;
        }
        .animate-on-scroll.visible {
            opacity: 1;
            transform: translateY(0);
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
                    <div class="d-flex justify-content-between align-items-center mb-3">
                        <h4 class="filter-title mb-0">Dashboard Filters</h4>
                        <button id="clearAllFilters" class="btn btn-outline-primary btn-sm">Clear All Filters</button>
                    </div>
                    <div class="row">
                        <div class="col-md-3 filter-section">
                            <label class="filter-label">Gender</label>
                            <select id="genderFilter" class="form-control select2-multi" multiple="multiple">
                                <!-- Will be populated by JavaScript -->
                            </select>
                        </div>
                        <div class="col-md-3 filter-section">
                            <label class="filter-label">Location</label>
                            <select id="locationFilter" class="form-control select2-multi" multiple="multiple">
                                <!-- Will be populated by JavaScript -->
                            </select>
                        </div>
                        <div class="col-md-3 filter-section">
                            <label class="filter-label">Function</label>
                            <select id="functionFilter" class="form-control select2-multi" multiple="multiple">
                                <!-- Will be populated by JavaScript -->
                            </select>
                        </div>
                        <div class="col-md-3 filter-section">
                            <label class="filter-label">Grade</label>
                            <select id="gradeFilter" class="form-control select2-multi" multiple="multiple">
                                <!-- Will be populated by JavaScript -->
                            </select>
                        </div>
                    </div>
                    <div class="row mt-2">
                        <div class="col-md-3 filter-section">
                            <label class="filter-label">Year</label>
                            <select id="yearFilter" class="form-control select2-multi" multiple="multiple">
                                <!-- Will be populated by JavaScript -->
                            </select>
                        </div>
                        <div class="col-md-3 filter-section">
                            <label class="filter-label">Tenure Band</label>
                            <select id="tenureFilter" class="form-control select2-multi" multiple="multiple">
                                <!-- Will be populated by JavaScript -->
                            </select>
                        </div>
                        <div class="col-md-6">
                            <div class="d-flex justify-content-end align-items-center h-100">
                                <p id="filterSummary" class="text-muted mb-0">No filters applied</p>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>

        <!-- Overall Statistics Section -->
        <div class="row animate-on-scroll">
            <div class="col-12">
                <h2 class="section-title">Overall Attrition Statistics</h2>
                <div class="row">
                    <div class="col-md-4">
                        <div class="card stat-card p-3">
                            <div class="stat-label">Total Employees</div>
                            <div id="totalEmployees" class="stat-value">-</div>
                            <div class="small">All employees in organization</div>
                        </div>
                    </div>
                    <div class="col-md-4">
                        <div class="card stat-card p-3">
                            <div class="stat-label">Total Exits</div>
                            <div id="totalExits" class="stat-value">-</div>
                            <div class="small">Employees who left</div>
                        </div>
                    </div>
                    <div class="col-md-4">
                        <div class="card stat-card p-3">
                            <div class="stat-label">Attrition Rate</div>
                            <div id="attritionRate" class="stat-value">-</div>
                            <div class="small">Percentage of employees who left</div>
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
        <div class="row mt-4 animate-on-scroll">
            <div class="col-12">
                <h2 class="section-title">Gender-wise Attrition</h2>
                <div class="row">
                    <div class="col-md-5">
                        <div class="card">
                            <div class="card-header">Gender Statistics</div>
                            <div class="card-body">
                                <div class="table-responsive">
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
                    </div>
                    <div class="col-md-7">
                        <div class="row">
                            <div class="col-md-6">
                                <div injecting class="card">
                                    <div class="card-header">Attrition Count by Gender</div>
                                    <div class="card-body">
                                        <div class="chart-container">
                                            <canvas id="genderBarChart"></canvas>
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
        <div class="row mt-4 animate-on-scroll">
            <div class="col-12">
                <h2 class="section-title">Location-wise Attrition</h2>
                <div class="row">
                    <div class="col-md-5">
                        <div class="card">
                            <div class="card-header">Location Statistics</div>
                            <div class="card-body">
                                <div class="table-responsive">
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
                    </div>
                    <div class="col-md-7">
                        <div class="row">
                            <div class="col-md-6">
                                <div class="card">
                                    <div class="card-header">Attrition Count by Location</div>
                                    <div class="card-body">
                                        <div class="chart-container">
                                            <canvas id="locationBarChart"></canvas>
                                        </div>
                                </div>
                            </div>
                            </div>
                            <div class="col-md-6">
                                <div class="card">
                                    <div class="card-header">Attrition Rate by Location</div>
                                    <div class="card-body">
                                        <div class="chart-container">
                                            <canvas id="locationPieChart"></canvas>
                                        </div>
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>

        <!-- Function Analysis Section -->
        <div class="row mt-4 animate-on-scroll">
            <div class="col-12">
                <h2 class="section-title">Function-wise Attrition</h2>
                <div class="row">
                    <div class="col-md-5">
                        <div class="card">
                            <div class="card-header">Function Statistics</div>
                            <div class="card-body">
                                <div class="table-responsive">
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
                    </td>
                    <div>
                    <div class="col-md-7">
                        <div class="row">
                            <div class="col-md-6">
                                <div class="card">
                                    <div class="card-header">Attrition Count by Function</div>
                                    <div class="card-body">
                                        <div class="chart-container">
                                            <canvas id="functionBarChart"></canvas>
                                        </div>
                                    </div>
                                </div>
                            </div>
                            <div class="col-md-6">
                                <div class="card">
                                    <div class="card-header">Attrition Rate by Function</div>
                                    <div class="card-body">
                                        <div class="chart-container">
                                            <canvas id="functionPieChart"></canvas>
                                        </div>
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>

        <!-- Tenure Analysis Section -->
        <div class="row mt-4 animate-on-scroll">
            <div class="col-12">
                <h2 class="section-title">Tenure Analysis of Exited Employees</h2>
                <div class="row">
                    <div class="col-md-5">
                        <div class="card">
                            <div class="card-header">Tenure Statistics</div>
                            <div class="card-body">
                                <div class="table-responsive">
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
                    </div>
                    <div class="col-md-7">
                        <div class="row">
                            <div class="col-md-6">
                                <div class="card">
                                    <div class="card-header">Attrition Count by Tenure</div>
                                    <div class="card-body">
                                        <div class="chart-container">
                                            <canvas id="tenureBarChart"></canvas>
                                        </div>
                                    </div>
                                </div>
                            </div>
                            <div class="col-md-6">
                                <div class="card">
                                    <div class="card-header">Attrition Distribution by Tenure</div>
                                    <div class="card-body">
                                        <div class="chart-container">
                                            <canvas id="tenurePieChart"></canvas>
                                        </div>
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>

        <!-- Grade Analysis Section -->
        <div class="row mt-4 animate-on-scroll">
            <div class="col-12">
                <h2 class="section-title">Grade-wise Attrition</h2>
                <div class="row">
                    <div class="col-md-5">
                        <div class="card">
                            <div class="card-header">Grade Statistics</div>
                            <div class="card-body">
                                <div class="table-responsive">
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
                    </div>
                    <div class="col-md-7">
                        <div class="row">
                            <div class="col-md-6">
                                <div class="card">
                                    <div class="card-header">Attrition Count by Grade</div>
                                    <div class="card-body">
                                        <div class="chart-container">
                                            <canvas id="gradeBarChart"></canvas>
                                        </div>
                                    </div>
                                </div>
                            </div>
                            <div class="col-md-6">
                                <div class="card">
                                    <div class="card-header">Attrition Rate by Grade</div>
                                    <div class="card-body">
                                        <div class="chart-container">
                                            <canvas id="gradePieChart"></canvas>
                                        </div>
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>

        <!-- Trend Analysis Section -->
        <div class="row mt-4 animate-on-scroll">
            <div class="col-12">
                <h2 class="section-title">Quarterly and Monthly Attrition Trends</h2>
                <div class="row">
                    <div class="col-md-6">
                        <div class="card">
                            <div class="card-header">Quarterly Attrition Trend</div>
                            <div class="card-body">
                                <div class="chart-container">
                                    <canvas id="quarterlyChart"></canvas>
                                </div>
                                <div class="table-responsive mt-3">
                                    <table class="table table-striped" id="quarterlyTable">
                                        <thead>
                                            <tr>
                                                <th>Quarter</th>
                                                <th>Exit Count</th>
                                            </tr>
                                        </thead>
                                        <tbody></tbody>
                                    </table>
                                </div>
                            </div>
                        </div>
                    </div>
                    <div class="col-md-6">
                        <div class="card">
                            <div class="card-header">Monthly Attrition Trend</div>
                            <div class="card-body">
                                <div class="chart-container">
                                    <canvas id="monthlyChart"></canvas>
                                </div>
                                <div class="table-responsive mt-3">
                                    <table class="table table-striped" id="monthlyTable">
                                        <thead>
                                            <tr>
                                                <th>Month</th>
                                                <th>Exit Count</th>
                                            </tr>
                                        </thead>
                                        <tbody></tbody>
                                    </table>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </div>

    <!-- Bootstrap JS and other libraries -->
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
    
    <script>
        // Global variables and settings
        const colors = {
            primary: '#004C99',
            secondary: '#4F81BD',
            accent: '#F79646',
            background: '#EDF3FE'
        };
        
        // Chart color palette
        const chartColors = [
            '#4F81BD', '#C0504D', '#9BBB59', '#8064A2', '#4BACC6',
            '#F79646', '#6DAA', '#355070', '#6B705C', '#B56576'
        ];
        
        // Global variables for data management
        let originalData = null;
        let filteredData = null;
        let charts = {};
        
        // Filter state
        const activeFilters = {
            gender: [],
            location: [],
            functionFilter: [],
            grade: [],
            yearFilter: [],
            tenure: []
        };
        
        // Initialize report with data
        function initializeReport(data) {
            originalData = JSON.parse(JSON.stringify(data));
            filteredData = JSON.parse(JSON.stringify(data));
            
            document.getElementById('reportDate').textContent += new Date().toLocaleString();
            
            $('.select2-multi').select2({
                placeholder: "Select options",
                allowClear: true,
                width: '100%'
            });
            
            populateFilterOptions(data);
            setupFilterEvents();
            updateDashboard(filteredData);
            setupScrollAnimations();
        }
        
        // Populate filter options from data
        function populateFilterOptions(data) {
            const genderItems = document.getElementById('genderFilter');
            data.slicers.gender.forEach(item => {
                const option = new Option(item.label, item.value);
                genderItems.appendChild(option);
            });
            
            const locationItems = document.getElementById('locationFilter');
            data.slicers.locationItems.forEach(item => {
                const option = new Option(item.label, item.value);
                locationItems.appendChild(option);
            });
            
            const functionItems = document.getElementById('functionFilter');
            functionItems.slicers.functionItems.forEach(item => {
                const option = new Option(item.label, item.value);
                functionItems.appendChild(option);
            });
            
            const gradeItems = document.getElementById('gradeFilter');
            data.slicers.grade.forEach(item => {
                const option = new Option(item.label, item.value);
                gradeItems.appendChild(option);
            });
            
            const yearItems = document.getElementById('yearFilter');
            data.slicers.year.forEach(item => {
                const option = new Option(item.label, item.value);
                yearItems.appendChild(option);
            });
            
            const tenureItems = document.getElementById('tenureFilter');
            const tenure_years = data.tenure.map(row => row.category);
            tenure_years.forEach(year => {
                const option = new Option(year, year);
                tenureItems.appendChild(option);
            });
        }
        
        // Set up filter event handlers
        function setupFilterEvents() {
            $('#genderFilter').on('change', function() {
                activeFilters.gender = $(this).val() || [];
                applyFilters();
            });
            
            $('#locationFilter').on('change', function() {
                activeFilters.location = $(this).val() || [];
                applyFilters();
            });
            
            $('#functionFilter').on('change', function() {
                activeFilters.functionFilter = $(this).val() || [];
                applyFilters();
            });
            
            $('#gradeFilter').on('change', function() {
                activeFilters.grade = $(this).val() || [];
                applyFilters();
            });
            
            $('#yearFilter').on('change', function() {
                activeFilters.yearFilter = $(this).val() || [];
                applyFilters();
            });
            
            $('#tenureFilter').on('change', function() {
                activeFilters.tenure = $(this).val() || [];
                applyFilters();
            });
            
            $('#clearAllFilters').on('click', function() {
                // Clear all Select2 filters
                const filters = [
                    '#genderFilter',
                    '#locationFilter',
                    '#functionFilter',
                    '#gradeFilter',
                    '#yearFilter',
                    '#tenureFilter'
                ];
                
                filters.forEach(filter => {
                    const $filter = $(filter);
                    $filter.val(null); // Clear programmatically
                    $filter.trigger('change.select2'); // Trigger Select2-specific change
                    $filter.select2('close'); // Ensure dropdown is closed
                });
                
                // Reset activeFilters
                for (const key in activeFilters) {
                    activeFilters[key] = [];
                }
                
                // Reset filteredData to originalData
                filteredData = JSON.parse(JSON.stringify(originalData));
                
                // Update dashboard and filter summary
                updateDashboard(filteredData);
                updateFilterSummary();
            });
        }
        
        // Apply all active filters to data
        function applyFilters() {
            // Check if any filters are applied
            const hasFilters = Object.values(activeFilters).some(filters => filters.length > 0);
            
            // If no filters, reset to original data
            if (!hasFilters) {
                filteredData = JSON.parse(JSON.stringify(originalData));
                updateOverallStatistics();
                updateFilterSummary();
                updateDashboard(filteredData);
                return;
            }
            
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
            
            if (activeFilters.functionFilter.length > 0) {
                filteredData.functionFilter = originalData.functionFilter.filter(row => 
                    activeFilters.functionFilter.includes(row.category)
                );
            }
            
            if (activeFilters.grade.length > 0) {
                filteredData.grade = originalData.grade.filter(row => 
                    activeFilters.grade.includes(row.category)
                );
            }
            
            if (activeFilters.tenure.length > 0) {
                filteredData.tenure = originalData.tenure.filter(row => 
                    activeFilters.tenure.includes(row.category)
                );
            }
            
            if (activeFilters.yearFilter.length > 0) {
                filteredData.quarterly = originalData.quarterly.filter(row => {
                    const yearMatch = row.period.match(/^(\d{4})/);
                    return yearMatch && activeFilters.yearFilter.includes(yearMatch[1]);
                });
                
                filteredData.monthly = originalData.monthly.filter(row => {
                    const yearMatch = row.period.match(/^(\d{4})/);
                    return yearMatch && activeFilters.yearFilter.includes(yearMatch[1]);
                });
            }
            
            updateOverallStatistics();
            updateFilterSummary();
            updateDashboard(filteredData);
        }
        
        // Update overall statistics based on current filters
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
        
        // Update filter summary text
        function updateFilterSummary() {
            const filterSummary = document.getElementById('filterSummary');
            const totalFilters = Object.values(activeFilters).reduce(
                (sum, filters) => sum + filters.length, 0
            );
            
            if (totalFilters === 0) {
                filterSummary.textContent = "No filters applied";
                return;
            }
            
            const filterTexts = [];
            if (activeFilters.gender.length > 0) {
                filterTexts.push(`Gender: ${activeFilters.gender.join(', ')}`);
            }
            if (activeFilters.location.length > 0) {
                filterTexts.push(`Location: ${activeFilters.location.length} selected`);
            }
            if (activeFilters.functionFilter.length > 0) {
                filterTexts.push(`Function: ${activeFilters.functionFilter.length} selected`);
            }
            if (activeFilters.grade.length > 0) {
                filterTexts.push(`Grade: ${activeFilters.grade.length} selected`);
            }
            if (activeFilters.tenure.length > 0) {
                filterTexts.push(`Tenure: ${activeFilters.tenure.length} selected`);
            }
            if (activeFilters.yearFilter.length > 0) {
                filterTexts.push(`Year: ${activeFilters.yearFilter.join(', ')}`);
            }
            
            filterSummary.textContent = `Filters applied: ${filterTexts.join(' | ')}`;
        }
        
        // Update dashboard with current data
        function updateDashboard(data) {
            updateOverallSection(data.overall);
            updateGenderSection(data.gender);
            updateLocationSection(data.location);
            updateFunctionSection(data.functionFilter);
            updateTenureSection(data.tenure);
            updateGradeSection(data.grade);
            updateTrendSection(data.quarterly, data.monthly);
        }
        
        // Update overall statistics section
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
                            backgroundColor: [colors.secondary, colors.accent],
                            borderWidth: 1
                        }]
                    },
                    options: {
                        responsive: true,
                        maintainAspectRatio: false,
                        plugins: {
                            legend: { position: 'bottom' },
                            tooltip: {
                                callbacks: {
                                    label: function(context) {
                                        const label = context.label || '';
                                        const value = context.raw || 0;
                                        const total = context.dataset.data.reduce((a, b) => a + b, 0);
                                        const percentage = Math.round((value / total) * 100);
                                        return `${label}: ${value} (${percentage}%)`;
                                    }
                                }
                            }
                        }
                    }
                });
            }
        }
        
        // Update gender section
        function updateGenderSection(genderData) {
            populateTable('genderTable', genderData);
            
            if (charts.genderBar) {
                charts.genderBar.data.labels = genderData.map(row => row.category);
                charts.genderBar.data.datasets[0].data = genderData.map(row => row.attritionCount);
                charts.genderBar.update();
            } else {
                const barCtx = document.getElementById('genderBarChart').getContext('2d');
                charts.genderBar = new Chart(barCtx, {
                    type: 'bar',
                    data: {
                        labels: genderData.map(row => row.category),
                        datasets: [{
                            label: 'Attrition Count',
                            data: genderData.map(row => row.attritionCount),
                            backgroundColor: chartColors
                        }]
                    },
                    options: {
                        responsive: true,
                        maintainAspectRatio: false,
                        plugins: { legend: { display: false } },
                        scales: {
                            y: {
                                beginAtZero: true,
                                title: { display: true, text: 'Number of Employees' }
                            }
                        }
                    }
                });
            }
            
            if (charts.genderPie) {
                charts.genderPie.data.labels = genderData.map(row => row.category);
                charts.genderPie.data.datasets[0].data = genderData.map(row => row.attritionRate);
                charts.genderPie.update();
            } else {
                const pieCtx = document.getElementById('genderPieChart').getContext('2d');
                charts.genderPie = new Chart(pieCtx, {
                    type: 'pie',
                    data: {
                        labels: genderData.map(row => row.category),
                        datasets: [{
                            data: genderData.map(row => row.attritionRate),
                            backgroundColor: chartColors
                        }]
                    },
                    options: {
                        responsive: true,
                        maintainAspectRatio: false,
                        plugins: {
                            tooltip: {
                                callbacks: {
                                    label: function(context) {
                                        return `${context.label}: ${context.raw}%`;
                                    }
                                }
                            }
                        }
                    }
                });
            }
        }
        
        // Update location section
        function updateLocationSection(locationData) {
            populateTable('locationTable', locationData);
            
            const topLocations = [...locationData]
                .sort((a, b) => b.attritionCount - a.attritionCount)
                .slice(0, 10);
            
            if (charts.locationBar) {
                charts.locationBar.data.labels = topLocations.map(row => row.category);
                charts.locationBar.data.datasets[0].data = topLocations.map(row => row.attritionCount);
                charts.locationBar.update();
            } else {
                const barCtx = document.getElementById('locationBarChart').getContext('2d');
                charts.locationBar = new Chart(barCtx, {
                    type: 'bar',
                    data: {
                        labels: topLocations.map(row => row.category),
                        datasets: [{
                            label: 'Attrition Count',
                            data: topLocations.map(row => row.attritionCount),
                            backgroundColor: chartColors
                        }]
                    },
                    options: {
                        responsive: true,
                        maintainAspectRatio: false,
                        plugins: { legend: { display: false } },
                        scales: {
                            y: {
                                beginAtZero: true,
                                title: { display: true, text: 'Number of Employees' }
                            },
                            x: {
                                ticks: { autoSkip: false, maxRotation: 45, minRotation: 45 }
                            }
                        }
                    }
                });
            }
            
            const topLocationsByRate = [...locationData]
                .sort((a, b) => b.attritionRate - a.attritionRate)
                .slice(0, 7);
                
            if (charts.locationPie) {
                charts.locationPie.data.labels = topLocationsByRate.map(row => row.category);
                charts.locationPie.data.datasets[0].data = topLocationsByRate.map(row => row.attritionRate);
                charts.locationPie.update();
            } else {
                const pieCtx = document.getElementById('locationPieChart').getContext('2d');
                charts.locationPie = new Chart(pieCtx, {
                    type: 'pie',
                    data: {
                        labels: topLocationsByRate.map(row => row.category),
                        datasets: [{
                            data: topLocationsByRate.map(row => row.attritionRate),
                            backgroundColor: chartColors
                        }]
                    },
                    options: {
                        responsive: true,
                        maintainAspectRatio: false,
                        plugins: {
                            tooltip: {
                                callbacks: {
                                    label: function(context) {
                                        return `${context.label}: ${context.raw}%`;
                                    }
                                }
                            }
                        }
                    }
                });
            }
        }
        
        // Update function section
        function updateFunctionSection(functionData) {
            populateTable('functionTable', functionData);
            
            const topFunctions = [...functionData]
                .sort((a, b) => b.attritionCount - a.attritionCount)
                .slice(0, 10);
            
            if (charts.functionBar) {
                charts.functionBar.data.labels = topFunctions.map(row => row.category);
                charts.functionBar.data.datasets[0].data = topFunctions.map(row => row.attritionCount);
                charts.functionBar.update();
            } else {
                const barCtx = document.getElementById('functionBarChart').getContext('2d');
                charts.functionBar = new Chart(barCtx, {
                    type: 'bar',
                    data: {
                        labels: topFunctions.map(row => row.category),
                        datasets: [{
                            label: 'Attrition Count',
                            data: topFunctions.map(row => row.attritionCount),
                            backgroundColor: chartColors
                        }]
                    },
                    options: {
                        responsive: true,
                        maintainAspectRatio: false,
                        plugins: { legend: { display: false } },
                        scales: {
                            y: {
                                beginAtZero: true,
                                title: { display: true, text: 'Number of Employees' }
                            },
                            x: {
                                ticks: { autoSkip: false, maxRotation: 45, minRotation: 45 }
                            }
                        }
                    }
                });
            }
            
            const topFunctionsByRate = [...functionData]
                .sort((a, b) => b.attritionRate - a.attritionRate)
                .slice(0, 7);
                
            if (charts.functionPie) {
                charts.functionPie.data.labels = topFunctionsByRate.map(row => row.category);
                charts.functionPie.data.datasets[0].data = topFunctionsByRate.map(row => row.attritionRate);
                charts.functionPie.update();
            } else {
                const pieCtx = document.getElementById('functionPieChart').getContext('2d');
                charts.functionPie = new Chart(pieCtx, {
                    type: 'pie',
                    data: {
                        labels: topFunctionsByRate.map(row => row.category),
                        datasets: [{
                            data: topFunctionsByRate.map(row => row.attritionRate),
                            backgroundColor: chartColors
                        }]
                    },
                    options: {
                        responsive: true,
                        maintainAspectRatio: false,
                        plugins: {
                            tooltip: {
                                callbacks: {
                                    label: function(context) {
                                        return `${context.label}: ${context.raw}%`;
                                    }
                                }
                            }
                        }
                    }
                });
            }
        }
        
        // Update tenure section
        function updateTenureSection(tenureData) {
            populateTable('tenureTable', tenureData);
            
            if (charts.tenureBar) {
                charts.tenureBar.data.labels = tenureData.map(row => row.category);
                charts.tenureBar.data.datasets[0].data = tenureData.map(row => row.count);
                charts.tenureBar.update();
            } else {
                const barCtx = document.getElementById('tenureBarChart').getContext('2d');
                charts.tenureBar = new Chart(barCtx, {
                    type: 'bar',
                    data: {
                        labels: tenureData.map(row => row.category),
                        datasets: [{
                            label: 'Number of Exits',
                            data: tenureData.map(row => row.count),
                            backgroundColor: chartColors
                        }]
                    },
                    options: {
                        responsive: true,
                        maintainAspectRatio: false,
                        plugins: { legend: { display: false } },
                        scales: {
                            y: {
                                beginAtZero: true,
                                title: { display: true, text: 'Number of Employees' }
                            }
                        }
                    }
                });
            }
            
            if (charts.tenurePie) {
                charts.tenurePie.data.labels = tenureData.map(row => row.category);
                charts.tenurePie.data.datasets[0].data = tenureData.map(row => row.percentage);
                charts.tenurePie.update();
            } else {
                const pieCtx = document.getElementById('tenurePieChart').getContext('2d');
                charts.tenurePie = new Chart(pieCtx, {
                    type: 'pie',
                    data: {
                        labels: tenureData.map(row => row.category),
                        datasets: [{
                            data: tenureData.map(row => row.percentage),
                            backgroundColor: chartColors
                        }]
                    },
                    options: {
                        responsive: true,
                        maintainAspectRatio: false,
                        plugins: {
                            tooltip: {
                                callbacks: {
                                    label: function(context) {
                                        return `${context.label}: ${context.raw}%`;
                                    }
                                }
                            }
                        }
                    }
                });
            }
        }
        
        // Update grade section
        function updateGradeSection(gradeData) {
            populateTable('gradeTable', gradeData);
            
            const topGrades = [...gradeData]
                .sort((a, b) => b.attritionCount - a.attritionCount)
                .slice(0, 10);
            
            if (charts.gradeBar) {
                charts.gradeBar.data.labels = topGrades.map(row => row.category);
                charts.gradeBar.data.datasets[0].data = topGrades.map(row => row.attritionCount);
                charts.gradeBar.update();
            } else {
                const barCtx = document.getElementById('gradeBarChart').getContext('2d');
                charts.gradeBar = new Chart(barCtx, {
                    type: 'bar',
                    data: {
                        labels: topGrades.map(row => row.category),
                        datasets: [{
                            label: 'Attrition Count',
                            data: topGrades.map(row => row.attritionCount),
                            backgroundColor: chartColors
                        }]
                    },
                    options: {
                        responsive: true,
                        maintainAspectRatio: false,
                        plugins: { legend: { display: false } },
                        scales: {
                            y: {
                                beginAtZero: true,
                                title: { display: true, text: 'Number of Employees' }
                            }
                        }
                    }
                });
            }
            
            const topGradesByRate = [...gradeData]
                .sort((a, b) => b.attritionRate - a.attritionRate)
                .slice(0, 7);
                
            if (charts.gradePie) {
                charts.gradePie.data.labels = topGradesByRate.map(row => row.category);
                charts.gradePie.data.datasets[0].data = topGradesByRate.map(row => row.attritionRate);
                charts.gradePie.update();
            } else {
                const pieCtx = document.getElementById('gradePieChart').getContext('2d');
                charts.gradePie = new Chart(pieCtx, {
                    type: 'pie',
                    data: {
                        labels: topGradesByRate.map(row => row.category),
                        datasets: [{
                            data: topGradesByRate.map(row => row.attritionRate),
                            backgroundColor: chartColors
                        }]
                    },
                    options: {
                        responsive: true,
                        maintainAspectRatio: false,
                        plugins: {
                            tooltip: {
                                callbacks: {
                                    label: function(context) {
                                        return `${context.label}: ${context.raw}%`;
                                    }
                                }
                            }
                        }
                    }
                });
            }
        }
        
        // Update trend section
        function updateTrendSection(quarterlyData, monthlyData) {
            const quarterlyTable = document.querySelector('#quarterlyTable tbody');
            quarterlyTable.innerHTML = '';
            quarterlyData.forEach(row => {
                const tr = document.createElement('tr');
                tr.innerHTML = `<td>${row.period}</td><td>${row.exitCount}</td>`;
                quarterlyTable.appendChild(tr);
            });
            
            const monthlyTable = document.querySelector('#monthlyTable tbody');
            monthlyTable.innerHTML = '';
            monthlyData.forEach(row => {
                const tr = document.createElement('tr');
                tr.innerHTML = `<td>${row.period}</td><td>${row.exitCount}</td>`;
                monthlyTable.appendChild(tr);
            });
            
            if (charts.quarterly) {
                charts.quarterly.data.labels = quarterlyData.map(row => row.period);
                charts.quarterly.data.datasets[0].data = quarterlyData.map(row => row.exitCount);
                charts.quarterly.update();
            } else {
                const ctx = document.getElementById('quarterlyChart').getContext('2d');
                charts.quarterly = new Chart(ctx, {
                    type: 'bar',
                    data: {
                        labels: quarterlyData.map(row => row.period),
                        datasets: [{
                            label: 'Exit Count',
                            data: quarterlyData.map(row => row.exitCount),
                            backgroundColor: colors.secondary
                        }]
                    },
                    options: {
                        responsive: true,
                        maintainAspectRatio: false,
                        plugins: { legend: { display: false } },
                        scales: {
                            y: {
                                beginAtZero: true,
                                title: { display: true, text: 'Number of Exits' }
                            }
                        }
                    }
                });
            }
            
            if (charts.monthly) {
                charts.monthly.data.labels = monthlyData.map(row => row.period);
                charts.monthly.data.datasets[0].data = monthlyData.map(row => row.exitCount);
                charts.monthly.update();
            } else {
                const ctx = document.getElementById('monthlyChart').getContext('2d');
                charts.monthly = new Chart(ctx, {
                    type: 'line',
                    data: {
                        labels: monthlyData.map(row => row.period),
                        datasets: [{
                            label: 'Exit Count',
                            data: monthlyData.map(row => row.exitCount),
                            borderColor: colors.accent,
                            backgroundColor: 'rgba(247, 150, 70, 0.2)',
                            borderWidth: 2,
                            tension: 0.2,
                            fill: true
                        }]
                    },
                    options: {
                        responsive: true,
                        maintainAspectRatio: false,
                        plugins: { legend: { display: false } },
                        scales: {
                            y: {
                                beginAtZero: true,
                                title: { display: true, text: 'Number of Exits' }
                            },
                            x: {
                                ticks: { autoSkip: true, maxTicksLimit: 12 }
                            }
                        }
                    }
                });
            }
        }
        
        // Helper function to populate tables
        function populateTable(tableId, data) {
            const tbody = document.querySelector(`#${tableId} tbody`);
            tbody.innerHTML = '';
            
            data.forEach(row => {
                const tr = document.createElement('tr');
                if (tableId === 'tenureTable') {
                    tr.innerHTML = `<td>${row.category}</td><td>${row.count}</td><td>${row.percentage}%</td>`;
                } else if (tableId === 'quarterlyTable' || tableId === 'monthlyTable') {
                    tr.innerHTML = `<td>${row.period}</td><td>${row.exitCount}</td>`;
                } else {
                    tr.innerHTML = `<td>${row.category}</td><td>${row.attritionCount}</td><td>${row.totalEmployees}</td><td>${row.attritionRate}%</td>`;
                }
                tbody.appendChild(tr);
            });
        }
        
        // Set up animations on scroll
        function setupScrollAnimations() {
            const animatedElements = document.querySelectorAll('.animate-on-scroll');
            checkElementsInView(animatedElements);
            window.addEventListener('scroll', () => checkElementsInView(animatedElements));
        }
        
        // Check if elements are in view
        function checkElementsInView(elements) {
            elements.forEach(element => {
                const elementTop = element.getBoundingClientRect().top;
                const elementVisible = 150;
                if (elementTop < window.innerHeight - elementVisible) {
                    element.classList.add('visible');
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
        
        # Replace data placeholder with actual data
        final_html = html_template.replace('REPLACE_WITH_DATA', json.dumps(report_data))
        
        # Save HTML file
        report_filename = f"Interactive_Attrition_Dashboard_{report_time}.html"
        html_path = f"{report_dir}/{report_filename}"
        with open(html_path, 'w', encoding='utf-8') as f:
            f.write(final_html)
        
        logger.info(f"Interactive HTML dashboard saved to {html_path}")
        return True, html_path, report_filename
    except Exception as e:
        logger.error(f"Failed to generate interactive HTML dashboard: {e}")
        return False, None, None
