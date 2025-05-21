# app.py
import pandas as pd
import numpy as np
import os
import shutil
import uuid
import logging
from datetime import datetime
from fastapi import FastAPI, UploadFile, File, HTTPException, BackgroundTasks
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.chart import BarChart, PieChart, LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.utils import get_column_letter
from src.core.logging import setup_logging

# Initialize logging
logger = setup_logging()


# Data validation functions
def ensure_employee_name_column(df):
    """Ensures 'Employee Name' column exists"""
    if 'Employee Name' in df.columns:
        return
    
    id_columns = ['Employee ID', 'EmployeeID', 'Emp ID', 'EmpID', 'ID']
    for col in id_columns:
        if col in df.columns:
            df['Employee Name'] = df[col]
            return
    
    raise ValueError("No 'Employee Name' or ID column found in the data")

def ensure_action_type_column(df):
    """Ensures 'Action Type' column exists"""
    if not hasattr(df, 'columns'):
        raise TypeError("Input must be a pandas DataFrame")
    
    if df.empty:
        raise ValueError("DataFrame is empty")
        
    if 'Action Type' in df.columns:
        df['Action Type'] = df['Action Type'].fillna('')
        return
    
    status_columns = ['Status', 'Employee Status', 'Employment Status', 
                      'Job Status', 'Action', 'Termination Status', 'Exit Status']
    
    for col in status_columns:
        if col in df.columns:
            df['Action Type'] = df[col]
            df['Action Type'] = df['Action Type'].fillna('')
            return
    
    raise ValueError("No 'Action Type' or status-related column found in the data")

def ensure_gender_column(df):
    """Ensures 'Gender' column exists"""
    if 'Gender' in df.columns:
        return
    
    if 'Sex' in df.columns:
        df['Gender'] = df['Sex']
        return

def ensure_function_column(df):
    """Ensures 'Function' column exists"""
    if 'Function' in df.columns:
        return
    
    if 'Department' in df.columns:
        df['Function'] = df['Department']
        return
    
    if 'Business Unit' in df.columns:
        df['Function'] = df['Business Unit']
        return

def ensure_grade_column(df):
    """Ensures 'Grade' column exists"""
    if 'Grade' in df.columns:
        return
    
    grade_alternatives = [
        "Job Grade", "Pay Grade", "Grade Level", "Salary Grade", 
        "Grade Code", "Band", "Position Grade", "Compensation Grade", 
        "Organizational Level", "Rank", "Position Rank"
    ]
    
    for alt_column in grade_alternatives:
        if alt_column in df.columns:
            df['Grade'] = df[alt_column]
            return

# Excel styling functions
def style_excel_table_headers(sheet, start_row, start_col, end_col, header_color="4F81BD"):
    """Apply styling to the header row of an Excel table"""
    try:
        fill = PatternFill(start_color=header_color, end_color=header_color, fill_type="solid")
        border = Border(
            left=Side(style='thin'), 
            right=Side(style='thin'), 
            top=Side(style='thin'), 
            bottom=Side(style='thin')
        )
        
        for col in range(start_col, end_col + 1):
            cell = sheet.cell(row=start_row, column=col)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.fill = fill
            cell.border = border
            
        return True
    except Exception as e:
        logger.warning(f"Failed to style Excel table headers: {e}")
        return False

def style_excel_table(sheet, start_row, start_col, end_row, end_col, 
                     header_color="4F81BD", 
                     alt_row_color="EDF3FE", 
                     border_style="thin", 
                     include_totals=False):
    """Apply comprehensive styling to an Excel table"""
    try:
        header_fill = PatternFill(start_color=header_color, end_color=header_color, fill_type="solid")
        alt_row_fill = PatternFill(start_color=alt_row_color, end_color=alt_row_color, fill_type="solid")
        
        border = Border(
            left=Side(style=border_style), 
            right=Side(style=border_style), 
            top=Side(style=border_style), 
            bottom=Side(style=border_style)
        )
        
        # Style header row
        for col in range(start_col, end_col + 1):
            cell = sheet.cell(row=start_row, column=col)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.fill = header_fill
            cell.border = border
            
        # Style data rows
        for row in range(start_row + 1, end_row + 1):
            row_fill = alt_row_fill if row % 2 == 0 else None
            
            for col in range(start_col, end_col + 1):
                cell = sheet.cell(row=row, column=col)
                cell.border = border
                
                if col == start_col:
                    cell.alignment = Alignment(horizontal='left', vertical='center')
                else:
                    cell.alignment = Alignment(horizontal='right', vertical='center')
                
                if row_fill:
                    cell.fill = row_fill
                    
        # Apply special formatting to totals row if included
        if include_totals:
            totals_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
            bottom_border = Border(
                left=Side(style=border_style),
                right=Side(style=border_style),
                top=Side(style=border_style),
                bottom=Side(style="medium")
            )
            
            for col in range(start_col, end_col + 1):
                cell = sheet.cell(row=end_row, column=col)
                cell.font = Font(bold=True)
                cell.border = bottom_border
                cell.fill = totals_fill
        
        # Auto-size columns
        for col in range(start_col, end_col + 1):
            column_letter = get_column_letter(col)
            max_length = 0
            
            for row in range(start_row, end_row + 1):
                cell = sheet.cell(row=row, column=col)
                if cell.value:
                    max_length = max(max_length, len(str(cell.value)))
            
            adjusted_width = max_length + 4
            sheet.column_dimensions[column_letter].width = min(adjusted_width, 40)
        
        return True
    except Exception as e:
        logger.error(f"Failed to style Excel table: {e}")
        return False

# Data loading function
def load_data(file_path):
    """Load and prepare the data"""
    try:
        df = pd.read_excel(file_path)
        
        # Apply data validation
        ensure_employee_name_column(df)
        ensure_action_type_column(df)
        ensure_gender_column(df)
        ensure_function_column(df)
        ensure_grade_column(df)
        
        df['Action Type'] = df['Action Type'].fillna('')
        logger.info(f"Successfully loaded data from {file_path}")
        return df
    except Exception as e:
        logger.error(f"Error loading data: {e}")
        return None

# Analysis functions
def add_excel_overall_statistics(wb, sheet, df, output_dir, report_time):
    """Add overall attrition statistics to the Excel report using native Excel charts"""
    try:
        # Add section title
        sheet['A1'] = '1. Overall Attrition Statistics'
        sheet['A1'].font = Font(size=14, bold=True)
        
        # Calculate overall statistics
        total_employees = len(df)
        exits = df[df['Action Type'].str.contains('Exit|Resignation|Termination', na=False)]
        total_exits = len(exits)
        attrition_rate = round((total_exits / total_employees * 100), 2)
        
        # Add data to sheet
        sheet['A3'] = "Total Employees:"
        sheet['B3'] = total_employees
        
        sheet['A4'] = "Total Exits:"
        sheet['B4'] = total_exits
        
        sheet['A5'] = "Overall Attrition Rate:"
        sheet['B5'] = f"{attrition_rate}%"
        
        # Style the cells
        for row in range(3, 6):
            sheet[f'A{row}'].font = Font(bold=True)
            
        # Add data for chart
        sheet['A7'] = "Metric"
        sheet['B7'] = "Percentage"
        sheet['A8'] = "Attrition"
        sheet['B8'] = attrition_rate
        sheet['A9'] = "Retention"
        sheet['B9'] = 100 - attrition_rate
            
        # Create native Excel chart
        chart = BarChart()
        chart.type = "col"
        chart.title = "Attrition vs. Retention Rate"
        chart.y_axis.title = "Percentage"
        chart.x_axis.title = "Category"
        
        data = Reference(sheet, min_col=2, min_row=7, max_row=9, max_col=2)
        categories = Reference(sheet, min_col=1, min_row=8, max_row=9)
        
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(categories)
        
        # Add data labels
        chart.dataLabels = DataLabelList()
        chart.dataLabels.showVal = True
        
        # Add chart to worksheet
        sheet.add_chart(chart, "D3")
        
        # Adjust column widths
        for col in range(1, 3):
            sheet.column_dimensions[get_column_letter(col)].width = 25
        
        logger.info("Successfully added overall statistics to Excel with native charts")
        return True
    except Exception as e:
        logger.error(f"Failed to add overall statistics to Excel: {e}")
        sheet['A1'] = "Error generating Overall Statistics section. Section skipped."
        return False

def add_excel_gender_analysis(wb, sheet, df, output_dir, report_time):
    """Add gender-wise attrition analysis to the Excel report using native Excel charts"""
    try:
        # Add section title
        sheet['A1'] = '2. Gender-wise Attrition'
        sheet['A1'].font = Font(size=14, bold=True)
        
        # Calculate gender statistics
        attrition_by_gender = df[df['Action Type'].str.contains('Exit|Resignation|Termination', na=False)].groupby('Gender').agg({
            'Employee Name': 'count'
        }).rename(columns={'Employee Name': 'Attrition Count'})
        
        total_by_gender = df.groupby('Gender')['Employee Name'].count()
        attrition_by_gender['Total Employees'] = total_by_gender
        attrition_by_gender['Attrition Rate %'] = round((attrition_by_gender['Attrition Count'] / attrition_by_gender['Total Employees'] * 100), 2)
        
        # Add table headers
        headers = ['Gender', 'Attrition Count', 'Total Employees', 'Attrition Rate %']
        for col, header in enumerate(headers, start=1):
            sheet.cell(row=3, column=col, value=header)
        
        # Style the headers
        style_excel_table_headers(sheet, 3, 1, len(headers))
        
        # Add data rows
        gender_df = attrition_by_gender.reset_index()
        for i, idx in enumerate(range(len(gender_df)), start=4):
            sheet.cell(row=i, column=1, value=str(gender_df.iloc[idx]['Gender']))
            sheet.cell(row=i, column=2, value=int(gender_df.iloc[idx]['Attrition Count']))
            sheet.cell(row=i, column=3, value=int(gender_df.iloc[idx]['Total Employees']))
            sheet.cell(row=i, column=4, value=float(gender_df.iloc[idx]['Attrition Rate %']))
        
        # Add chart data to specific cells
        start_row = 4 + len(gender_df) + 2
        sheet.cell(row=start_row, column=1, value="Chart Data")
        sheet.cell(row=start_row, column=1).font = Font(bold=True)
        
        # Add headers for chart data
        sheet.cell(row=start_row+1, column=1, value="Gender")
        sheet.cell(row=start_row+1, column=2, value="Attrition Count")
        sheet.cell(row=start_row+1, column=3, value="Attrition Rate %")
        
        # Add data for charts
        for i, idx in enumerate(range(len(gender_df)), start=0):
            sheet.cell(row=start_row+2+i, column=1, value=str(gender_df.iloc[idx]['Gender']))
            sheet.cell(row=start_row+2+i, column=2, value=int(gender_df.iloc[idx]['Attrition Count']))
            sheet.cell(row=start_row+2+i, column=3, value=float(gender_df.iloc[idx]['Attrition Rate %']))
        
        # Create bar chart for attrition count
        bar_chart = BarChart()
        bar_chart.title = "Attrition Count by Gender"
        bar_chart.y_axis.title = "Number of Employees"
        bar_chart.x_axis.title = "Gender"
        
        data = Reference(sheet, min_col=2, min_row=start_row+1, max_row=start_row+1+len(gender_df), max_col=2)
        categories = Reference(sheet, min_col=1, min_row=start_row+2, max_row=start_row+1+len(gender_df))
        
        bar_chart.add_data(data, titles_from_data=True)
        bar_chart.set_categories(categories)
        
        # Add data labels
        bar_chart.dataLabels = DataLabelList()
        bar_chart.dataLabels.showVal = True
        
        # Add bar chart to worksheet
        chart_position = "A" + str(start_row+2+len(gender_df)+2)
        sheet.add_chart(bar_chart, chart_position)
        
        # Create pie chart for attrition rate
        pie_chart = PieChart()
        pie_chart.title = "Attrition Rate Distribution by Gender"
        
        data = Reference(sheet, min_col=3, min_row=start_row+1, max_row=start_row+1+len(gender_df), max_col=3)
        labels = Reference(sheet, min_col=1, min_row=start_row+2, max_row=start_row+1+len(gender_df))
        
        pie_chart.add_data(data, titles_from_data=True)
        pie_chart.set_categories(labels)
        
        # Add data labels showing percentages
        pie_chart.dataLabels = DataLabelList()
        pie_chart.dataLabels.showPercent = True
        
        # Add pie chart to worksheet
        pie_position = "F" + str(start_row+2+len(gender_df)+2)
        sheet.add_chart(pie_chart, pie_position)
        
        # Adjust column widths
        for col in range(1, len(headers) + 1):
            sheet.column_dimensions[get_column_letter(col)].width = 20
        
        logger.info("Successfully added gender analysis to Excel with native charts")
        return True
    except Exception as e:
        logger.error(f"Failed to add gender analysis to Excel: {e}")
        sheet['A1'] = "Error generating Gender Analysis section. Section skipped."
        return False

def add_excel_location_analysis(wb, sheet, df, output_dir, report_time):
    """Add location-wise attrition analysis to the Excel report using native Excel charts"""
    try:
        # Add section title
        sheet['A1'] = '3. Location-wise Attrition (Locations with ≥50 Employees)'
        sheet['A1'].font = Font(size=14, bold=True)
        
        # Calculate location statistics
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
        
        # Add table headers
        headers = ['Location', 'Attrition Count', 'Total Employees', 'Attrition Rate %']
        for col, header in enumerate(headers, start=1):
            sheet.cell(row=3, column=col, value=header)
        
        # Style the headers
        style_excel_table_headers(sheet, 3, 1, len(headers))
        
        # Add data rows
        location_df = attrition_by_location.reset_index()
        for i, idx in enumerate(range(len(location_df)), start=4):
            sheet.cell(row=i, column=1, value=str(location_df.iloc[idx]['Job Location']))
            sheet.cell(row=i, column=2, value=int(location_df.iloc[idx]['Attrition Count']))
            sheet.cell(row=i, column=3, value=int(location_df.iloc[idx]['Total Employees']))
            sheet.cell(row=i, column=4, value=float(location_df.iloc[idx]['Attrition Rate %']))
        
        # Add chart data to specific cells
        start_row = 4 + len(location_df) + 2
        sheet.cell(row=start_row, column=1, value="Chart Data")
        sheet.cell(row=start_row, column=1).font = Font(bold=True)
        
        # Add headers for chart data
        sheet.cell(row=start_row+1, column=1, value="Location")
        sheet.cell(row=start_row+1, column=2, value="Attrition Count")
        sheet.cell(row=start_row+1, column=3, value="Attrition Rate %")
        
        # Add data for charts
        for i, idx in enumerate(range(len(location_df)), start=0):
            sheet.cell(row=start_row+2+i, column=1, value=str(location_df.iloc[idx]['Job Location']))
            sheet.cell(row=start_row+2+i, column=2, value=int(location_df.iloc[idx]['Attrition Count']))
            sheet.cell(row=start_row+2+i, column=3, value=float(location_df.iloc[idx]['Attrition Rate %']))
        
        # Create bar chart for attrition count
        bar_chart = BarChart()
        bar_chart.type = "col"
        bar_chart.title = "Attrition Count by Location"
        bar_chart.y_axis.title = "Number of Employees"
        bar_chart.x_axis.title = "Location"
        bar_chart.height = 15  # Make the chart taller (in units of character height)
        bar_chart.width = 20   # Make the chart wider (in units of character width)
        
        data = Reference(sheet, min_col=2, min_row=start_row+1, max_row=start_row+1+len(location_df), max_col=2)
        categories = Reference(sheet, min_col=1, min_row=start_row+2, max_row=start_row+1+len(location_df))
        
        bar_chart.add_data(data, titles_from_data=True)
        bar_chart.set_categories(categories)
        
        # Add data labels
        bar_chart.dataLabels = DataLabelList()
        bar_chart.dataLabels.showVal = True
        
        # Add bar chart to worksheet
        chart_position = "A" + str(start_row+2+len(location_df)+2)
        sheet.add_chart(bar_chart, chart_position)
        
        # Create pie chart for attrition rate
        pie_chart = PieChart()
        pie_chart.title = "Attrition Rate Distribution by Location"
        pie_chart.height = 15  # Make the chart taller
        pie_chart.width = 20   # Make the chart wider
        
        data = Reference(sheet, min_col=3, min_row=start_row+1, max_row=start_row+1+len(location_df), max_col=3)
        labels = Reference(sheet, min_col=1, min_row=start_row+2, max_row=start_row+1+len(location_df))
        
        pie_chart.add_data(data, titles_from_data=True)
        pie_chart.set_categories(labels)
        
        # Add data labels showing percentages
        pie_chart.dataLabels = DataLabelList()
        pie_chart.dataLabels.showPercent = True
        
        # Add pie chart to worksheet
        pie_position = "I" + str(start_row+2+len(location_df)+2)  # Place chart to the right
        sheet.add_chart(pie_chart, pie_position)
        
        # Adjust column widths
        for col in range(1, len(headers) + 1):
            sheet.column_dimensions[get_column_letter(col)].width = 20
        
        logger.info("Successfully added location analysis to Excel with native charts")
        return True
    except Exception as e:
        logger.error(f"Failed to add location analysis to Excel: {e}")
        sheet['A1'] = "Error generating Location Analysis section. Section skipped."
        return False

def add_excel_function_analysis(wb, sheet, df, output_dir, report_time):
    """Add function-wise attrition analysis to the Excel report using native Excel charts"""
    try:
        # Add section title
        sheet['A1'] = '4. Function-wise Attrition (Functions with ≥20 Employees)'
        sheet['A1'].font = Font(size=14, bold=True)
        
        # Calculate function statistics
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
        
        # Add table headers
        headers = ['Function', 'Attrition Count', 'Total Employees', 'Attrition Rate %']
        for col, header in enumerate(headers, start=1):
            sheet.cell(row=3, column=col, value=header)
        
        # Style the headers
        style_excel_table_headers(sheet, 3, 1, len(headers))
        
        # Add data rows
        function_df = attrition_by_function.reset_index()
        for i, idx in enumerate(range(len(function_df)), start=4):
            sheet.cell(row=i, column=1, value=str(function_df.iloc[idx]['Function']))
            sheet.cell(row=i, column=2, value=int(function_df.iloc[idx]['Attrition Count']))
            sheet.cell(row=i, column=3, value=int(function_df.iloc[idx]['Total Employees']))
            sheet.cell(row=i, column=4, value=float(function_df.iloc[idx]['Attrition Rate %']))
        
        # Add chart data to specific cells
        start_row = 4 + len(function_df) + 2
        sheet.cell(row=start_row, column=1, value="Chart Data")
        sheet.cell(row=start_row, column=1).font = Font(bold=True)
        
        # Add headers for chart data
        sheet.cell(row=start_row+1, column=1, value="Function")
        sheet.cell(row=start_row+1, column=2, value="Attrition Count")
        sheet.cell(row=start_row+1, column=3, value="Attrition Rate %")
        
        # Add data for charts
        for i, idx in enumerate(range(len(function_df)), start=0):
            sheet.cell(row=start_row+2+i, column=1, value=str(function_df.iloc[idx]['Function']))
            sheet.cell(row=start_row+2+i, column=2, value=int(function_df.iloc[idx]['Attrition Count']))
            sheet.cell(row=start_row+2+i, column=3, value=float(function_df.iloc[idx]['Attrition Rate %']))
        
        # Create bar chart for attrition count
        bar_chart = BarChart()
        bar_chart.type = "col"
        bar_chart.title = "Attrition Count by Function"
        bar_chart.y_axis.title = "Number of Employees"
        bar_chart.x_axis.title = "Function"
        bar_chart.height = 15
        bar_chart.width = 20
        
        data = Reference(sheet, min_col=2, min_row=start_row+1, max_row=start_row+1+len(function_df), max_col=2)
        categories = Reference(sheet, min_col=1, min_row=start_row+2, max_row=start_row+1+len(function_df))
        
        bar_chart.add_data(data, titles_from_data=True)
        bar_chart.set_categories(categories)
        
        # Add data labels
        bar_chart.dataLabels = DataLabelList()
        bar_chart.dataLabels.showVal = True
        
        # Add bar chart to worksheet
        chart_position = "A" + str(start_row+2+len(function_df)+2)
        sheet.add_chart(bar_chart, chart_position)
        
        # Create pie chart for attrition rate
        pie_chart = PieChart()
        pie_chart.title = "Attrition Rate Distribution by Function"
        pie_chart.height = 15
        pie_chart.width = 20
        
        data = Reference(sheet, min_col=3, min_row=start_row+1, max_row=start_row+1+len(function_df), max_col=3)
        labels = Reference(sheet, min_col=1, min_row=start_row+2, max_row=start_row+1+len(function_df))
        
        pie_chart.add_data(data, titles_from_data=True)
        pie_chart.set_categories(labels)
        
        # Add data labels showing percentages
        pie_chart.dataLabels = DataLabelList()
        pie_chart.dataLabels.showPercent = True
        
        # Add pie chart to worksheet
        pie_position = "I" + str(start_row+2+len(function_df)+2)
        sheet.add_chart(pie_chart, pie_position)
        
        # Adjust column widths
        for col in range(1, len(headers) + 1):
            sheet.column_dimensions[get_column_letter(col)].width = 25
        
        logger.info("Successfully added function analysis to Excel with native charts")
        return True
    except Exception as e:
        logger.error(f"Failed to add function analysis to Excel: {e}")
        sheet['A1'] = "Error generating Function Analysis section. Section skipped."
        return False

def add_excel_tenure_analysis(wb, sheet, df, output_dir, report_time):
    """Add tenure analysis of exited employees to the Excel report using native Excel charts"""
    try:
        # Add section title
        sheet['A1'] = '5. Tenure Analysis of Exited Employees'
        sheet['A1'].font = Font(size=14, bold=True)
        
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
        
        # Add table headers
        headers = ['Tenure Band', 'Number of Exits', 'Percentage']
        for col, header in enumerate(headers, start=1):
            sheet.cell(row=3, column=col, value=header)
        
        # Style the headers
        style_excel_table_headers(sheet, 3, 1, len(headers))
        
        # Add data rows
        for i, idx in enumerate(range(len(tenure_analysis)), start=4):
            sheet.cell(row=i, column=1, value=str(tenure_analysis.iloc[idx]['Tenure Band']))
            sheet.cell(row=i, column=2, value=int(tenure_analysis.iloc[idx]['Count']))
            sheet.cell(row=i, column=3, value=float(tenure_analysis.iloc[idx]['Percentage']))
        
        # Add chart data to specific cells
        start_row = 4 + len(tenure_analysis) + 2
        sheet.cell(row=start_row, column=1, value="Chart Data")
        sheet.cell(row=start_row, column=1).font = Font(bold=True)
        
        # Add headers for chart data
        sheet.cell(row=start_row+1, column=1, value="Tenure Band")
        sheet.cell(row=start_row+1, column=2, value="Number of Exits")
        sheet.cell(row=start_row+1, column=3, value="Percentage")
        
        # Add data for charts
        for i, idx in enumerate(range(len(tenure_analysis)), start=0):
            sheet.cell(row=start_row+2+i, column=1, value=str(tenure_analysis.iloc[idx]['Tenure Band']))
            sheet.cell(row=start_row+2+i, column=2, value=int(tenure_analysis.iloc[idx]['Count']))
            sheet.cell(row=start_row+2+i, column=3, value=float(tenure_analysis.iloc[idx]['Percentage']))
        
        # Create bar chart for attrition count by tenure
        bar_chart = BarChart()
        bar_chart.type = "col"
        bar_chart.title = "Attrition Count by Tenure"
        bar_chart.y_axis.title = "Number of Employees"
        bar_chart.x_axis.title = "Tenure Band"
        bar_chart.height = 15
        bar_chart.width = 20
        
        data = Reference(sheet, min_col=2, min_row=start_row+1, max_row=start_row+1+len(tenure_analysis), max_col=2)
        categories = Reference(sheet, min_col=1, min_row=start_row+2, max_row=start_row+1+len(tenure_analysis))
        
        bar_chart.add_data(data, titles_from_data=True)
        bar_chart.set_categories(categories)
        
        # Add data labels
        bar_chart.dataLabels = DataLabelList()
        bar_chart.dataLabels.showVal = True
        
        # Add bar chart to worksheet
        chart_position = "A" + str(start_row+2+len(tenure_analysis)+2)
        sheet.add_chart(bar_chart, chart_position)
        
        # Create pie chart for tenure distribution
        pie_chart = PieChart()
        pie_chart.title = "Attrition Distribution by Tenure"
        pie_chart.height = 15
        pie_chart.width = 20
       
        data = Reference(sheet, min_col=3, min_row=start_row+1, max_row=start_row+1+len(tenure_analysis), max_col=3)
        labels = Reference(sheet, min_col=1, min_row=start_row+2, max_row=start_row+1+len(tenure_analysis))
       
        pie_chart.add_data(data, titles_from_data=True)
        pie_chart.set_categories(labels)
       
        # Add data labels showing percentages
        pie_chart.dataLabels = DataLabelList()
        pie_chart.dataLabels.showPercent = True
       
        # Add pie chart to worksheet
        pie_position = "J" + str(start_row+2+len(tenure_analysis)+2)
        sheet.add_chart(pie_chart, pie_position)
       
        # Adjust column widths
        for col in range(1, len(headers) + 1):
            sheet.column_dimensions[get_column_letter(col)].width = 20
       
        logger.info("Successfully added tenure analysis to Excel with native charts")
        return True
    except Exception as e:
        logger.error(f"Failed to add tenure analysis to Excel: {e}")
        sheet['A1'] = "Error generating Tenure Analysis section. Section skipped."
        return False

def add_excel_grade_analysis(wb, sheet, df, output_dir, report_time):
    """Add grade-wise attrition analysis to the Excel report using native Excel charts with enhanced table formatting"""
    try:
        # Add section title
        sheet['A1'] = '6. Grade-wise Attrition'
        sheet['A1'].font = Font(size=14, bold=True)
        
        # Calculate grade statistics
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
        
        # Convert to DataFrame and sort for better presentation
        grade_df = attrition_by_grade.reset_index()
        # Optional: Sort by attrition rate descending
        grade_df = grade_df.sort_values('Attrition Rate %', ascending=False)
        
        # Add table headers
        headers = ['Grade', 'Attrition Count', 'Total Employees', 'Attrition Rate %']
        for col, header in enumerate(headers, start=1):
            sheet.cell(row=3, column=col, value=header)
        
        # Add data rows with better formatting
        for i, (idx, row) in enumerate(grade_df.iterrows(), start=4):
            # Grade
            sheet.cell(row=i, column=1, value=str(row['Grade']))
            
            # Attrition Count (formatted as whole number)
            sheet.cell(row=i, column=2, value=int(row['Attrition Count']))
            
            # Total Employees (formatted as whole number)
            sheet.cell(row=i, column=3, value=int(row['Total Employees']))
            
            # Attrition Rate % (percentage format with 2 decimal places)
            rate_cell = sheet.cell(row=i, column=4, value=float(row['Attrition Rate %']))
            rate_cell.number_format = '0.00"%"'
        
        # Add a totals row
        totals_row = 4 + len(grade_df)
        sheet.cell(row=totals_row, column=1, value="Total")
        sheet.cell(row=totals_row, column=2, value=int(grade_df['Attrition Count'].sum()))
        sheet.cell(row=totals_row, column=3, value=int(grade_df['Total Employees'].sum()))
        
        # Calculate weighted average for the total attrition rate
        overall_rate = (grade_df['Attrition Count'].sum() / grade_df['Total Employees'].sum() * 100)
        rate_cell = sheet.cell(row=totals_row, column=4, value=overall_rate)
        rate_cell.number_format = '0.00"%"'
        
        # Apply enhanced table styling including the totals row
        style_excel_table(
            sheet=sheet,
            start_row=3,
            start_col=1,
            end_row=totals_row,
            end_col=len(headers),
            header_color="4F81BD",  # Professional blue
            alt_row_color="EDF3FE",  # Light blue alternating rows
            border_style="thin",
            include_totals=True
        )
        
        # Add chart data to specific cells
        start_row = totals_row + 3
        sheet.cell(row=start_row, column=1, value="Chart Data")
        sheet.cell(row=start_row, column=1).font = Font(bold=True)
        
        # Add headers for chart data
        sheet.cell(row=start_row+1, column=1, value="Grade")
        sheet.cell(row=start_row+1, column=2, value="Attrition Count")
        sheet.cell(row=start_row+1, column=3, value="Attrition Rate %")
        
        # Add data for charts (use top 8 grades by attrition rate for better chart visibility if many grades)
        chart_data = grade_df.head(8) if len(grade_df) > 8 else grade_df
        for i, (idx, row) in enumerate(chart_data.iterrows(), start=0):
            sheet.cell(row=start_row+2+i, column=1, value=str(row['Grade']))
            sheet.cell(row=start_row+2+i, column=2, value=int(row['Attrition Count']))
            sheet.cell(row=start_row+2+i, column=3, value=float(row['Attrition Rate %']))
        
        # Create bar chart for attrition count by grade
        bar_chart = BarChart()
        bar_chart.type = "col"
        bar_chart.title = "Attrition Count by Grade"
        bar_chart.y_axis.title = "Number of Employees"
        bar_chart.x_axis.title = "Grade"
        bar_chart.height = 15
        bar_chart.width = 20
        bar_chart.style = 10  # Use a nicer chart style
        
        data = Reference(sheet, min_col=2, min_row=start_row+1, max_row=start_row+1+len(chart_data), max_col=2)
        categories = Reference(sheet, min_col=1, min_row=start_row+2, max_row=start_row+1+len(chart_data))
        
        bar_chart.add_data(data, titles_from_data=True)
        bar_chart.set_categories(categories)
        
        # Add data labels
        bar_chart.dataLabels = DataLabelList()
        bar_chart.dataLabels.showVal = True
        
        # Add bar chart to worksheet
        chart_position = "A" + str(start_row+2+len(chart_data)+2)
        sheet.add_chart(bar_chart, chart_position)
        
        # Create pie chart for grade distribution
        pie_chart = PieChart()
        pie_chart.title = "Attrition Rate Distribution by Grade"
        pie_chart.height = 15
        pie_chart.width = 20
        pie_chart.style = 10  # Use a nicer chart style
        
        data = Reference(sheet, min_col=3, min_row=start_row+1, max_row=start_row+1+len(chart_data), max_col=3)
        labels = Reference(sheet, min_col=1, min_row=start_row+2, max_row=start_row+1+len(chart_data))
        
        pie_chart.add_data(data, titles_from_data=True)
        pie_chart.set_categories(labels)
        
        # Add data labels showing percentages
        pie_chart.dataLabels = DataLabelList()
        pie_chart.dataLabels.showPercent = True
        
        # Add pie chart to worksheet
        pie_position = "I" + str(start_row+2+len(chart_data)+2)
        sheet.add_chart(pie_chart, pie_position)
        
        # Adjust column widths for better readability
        for col, width in enumerate([20, 18, 18, 18], start=1):
            sheet.column_dimensions[get_column_letter(col)].width = width
        
        logger.info("Successfully added grade analysis to Excel with native charts")
        return True
    except Exception as e:
        logger.error(f"Failed to add grade analysis to Excel: {e}")
        sheet['A1'] = "Error generating Grade Analysis section. Section skipped."
        return False

def add_excel_trend_analysis(wb, sheet, df, output_dir, report_time):
    """Add quarterly and monthly attrition trend analysis to the Excel report using native Excel charts"""
    try:
        # Add section title
        sheet['A1'] = '7. Quarterly and Monthly Attrition Trends'
        sheet['A1'].font = Font(size=14, bold=True)
        
        # Filter to include only exit actions
        exits_df = df[df['Action Type'].str.contains('Exit|Resignation|Termination', na=False)].copy()

        # Ensure dates are in datetime format
        exits_df['Action Date'] = pd.to_datetime(exits_df['Action Date'])

        # Add month and quarter columns
        exits_df['Month'] = exits_df['Action Date'].dt.strftime('%Y-%m')
        exits_df['Quarter'] = exits_df['Action Date'].dt.to_period('Q').astype(str)

        # Calculate quarterly trends
        quarterly_trends = exits_df.groupby('Quarter').size().reset_index(name='Exit Count')
        quarterly_trends = quarterly_trends.sort_values('Quarter')  # Sort by quarter chronologically
        
        # Calculate monthly trends
        monthly_trends = exits_df.groupby('Month').size().reset_index(name='Exit Count')
        monthly_trends = monthly_trends.sort_values('Month')  # Sort by month chronologically
        monthly_trends['Month'] = pd.to_datetime(monthly_trends['Month'])
        
        # Add quarterly trend section
        sheet['A3'] = 'Quarterly Attrition Trend'
        sheet['A3'].font = Font(size=12, bold=True)
        
        # Add quarterly table headers
        quarterly_headers = ['Quarter', 'Exit Count']
        for col, header in enumerate(quarterly_headers, start=1):
            sheet.cell(row=4, column=col, value=header)
        
        # Style the headers
        style_excel_table_headers(sheet, 4, 1, len(quarterly_headers))
        
        # Add quarterly data rows
        for i, idx in enumerate(range(len(quarterly_trends)), start=5):
            sheet.cell(row=i, column=1, value=str(quarterly_trends.iloc[idx]['Quarter']))
            sheet.cell(row=i, column=2, value=int(quarterly_trends.iloc[idx]['Exit Count']))
        
        # Create bar chart for quarterly trend
        bar_chart = BarChart()
        bar_chart.type = "col"
        bar_chart.title = "Quarterly Attrition Trend"
        bar_chart.y_axis.title = "Exit Count"
        bar_chart.x_axis.title = "Quarter"
        bar_chart.height = 15
        bar_chart.width = 20
        
        data = Reference(sheet, min_col=2, min_row=4, max_row=4+len(quarterly_trends), max_col=2)
        categories = Reference(sheet, min_col=1, min_row=5, max_row=4+len(quarterly_trends))
        
        bar_chart.add_data(data, titles_from_data=True)
        bar_chart.set_categories(categories)
        
        # Add data labels
        bar_chart.dataLabels = DataLabelList()
        bar_chart.dataLabels.showVal = True
        
        # Add chart to worksheet
        quarterly_chart_row = 5 + len(quarterly_trends) + 2
        sheet.add_chart(bar_chart, "A" + str(quarterly_chart_row))
        
        # Add monthly trend section
        monthly_start_row = quarterly_chart_row + 20  # Leave space after quarterly chart
        sheet.cell(row=monthly_start_row, column=1, value='Monthly Attrition Trend')
        sheet.cell(row=monthly_start_row, column=1).font = Font(size=12, bold=True)
        
        # Add monthly table headers
        monthly_headers = ['Month', 'Exit Count']
        for col, header in enumerate(monthly_headers, start=1):
            sheet.cell(row=monthly_start_row + 1, column=col, value=header)
        
        # Style the headers
        style_excel_table_headers(sheet, monthly_start_row + 1, 1, len(monthly_headers))
        
        # Add monthly data rows
        for i, idx in enumerate(range(len(monthly_trends)), start=0):
            month_cell_value = monthly_trends.iloc[idx]['Month'].strftime('%Y-%m')
            sheet.cell(row=monthly_start_row + 2 + i, column=1, value=month_cell_value)
            sheet.cell(row=monthly_start_row + 2 + i, column=2, value=int(monthly_trends.iloc[idx]['Exit Count']))
        
        # Create line chart for monthly trend
        line_chart = LineChart()
        line_chart.title = "Monthly Attrition Trend"
        line_chart.y_axis.title = "Exit Count"
        line_chart.x_axis.title = "Month"
        line_chart.height = 15
        line_chart.width = 30  # Make it wider to accommodate more months
        
        data = Reference(sheet, min_col=2, min_row=monthly_start_row + 1, max_row=monthly_start_row + 1 + len(monthly_trends), max_col=2)
        categories = Reference(sheet, min_col=1, min_row=monthly_start_row + 2, max_row=monthly_start_row + 1 + len(monthly_trends))
        
        line_chart.add_data(data, titles_from_data=True)
        line_chart.set_categories(categories)
        
        # Style the line chart
        s = line_chart.series[0]
        s.marker.symbol = "circle"
        s.marker.size = 10
        s.smooth = True  # Make the line smoother
        
        # Add data labels
        line_chart.dataLabels = DataLabelList()
        line_chart.dataLabels.showVal = True
        
        # Add line chart to worksheet
        monthly_chart_row = monthly_start_row + 2 + len(monthly_trends) + 2
        sheet.add_chart(line_chart, "A" + str(monthly_chart_row))
        
        # Adjust column widths
        for col in range(1, 3):
            sheet.column_dimensions[get_column_letter(col)].width = 20
        
        logger.info("Successfully added trend analysis to Excel with native charts")
        return True
    
    except Exception as e:
        logger.error(f"Failed to add trend analysis to Excel: {e}")
        sheet['A1'] = "Error generating Trend Analysis section. Section skipped."
        return False

def create_excel_attrition_report(df, output_dir):
    """Create a comprehensive automated attrition analysis report in Excel format with native Excel charts"""
    try:
        # Create timestamp for the report
        report_time = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        # Create output directory if it doesn't exist
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
        
        # Initialize Excel workbook
        wb = Workbook()
        
        # Get the active sheet (first sheet)
        title_sheet = wb.active
        title_sheet.title = "Dashboard"
        
        # Add title and report information
        title_sheet['B2'] = 'AUTOMATED ATTRITION ANALYSIS REPORT'
        title_sheet['B3'] = 'BY VIMAL SINGH'
        title_sheet['B4'] = f"Report Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
        
        # Style title
        title_sheet['B2'].font = Font(size=18, bold=True, color="004C99")
        title_sheet['B3'].font = Font(size=14, bold=True, color="004C99")
        title_sheet['B4'].font = Font(size=11)
        
        # Create sheets for each section
        section_names = [
            "Overall Statistics",
            "Gender Analysis",
            "Location Analysis",
            "Function Analysis", 
            "Tenure Analysis",
            "Grade Analysis",
            "Trend Analysis"
        ]
        
        for section_name in section_names:
            wb.create_sheet(title=section_name)
        
        # List of section functions to call
        section_functions = [
            (add_excel_overall_statistics, "Overall Statistics"),
            (add_excel_gender_analysis, "Gender Analysis"),
            (add_excel_location_analysis, "Location Analysis"),
            (add_excel_function_analysis, "Function Analysis"),
            (add_excel_tenure_analysis, "Tenure Analysis"),
            (add_excel_grade_analysis, "Grade Analysis"),
            (add_excel_trend_analysis, "Trend Analysis")
        ]
        
        # Execute each section, skipping any that fail
        successful_sections = 0
        for section_func, section_name in section_functions:
            logger.info(f"Generating Excel {section_name} section...")
            sheet = wb[section_name]
            
            try:
                result = section_func(wb, sheet, df, output_dir, report_time)
                if result:
                    successful_sections += 1
                else:
                    logger.warning(f"Section function returned False: {section_name}")
            except Exception as e:
                logger.error(f"Error in Excel {section_name} section: {e}")
                sheet['A1'] = f"Error generating {section_name} section. Section skipped."
                sheet['A1'].font = Font(color="FF0000")
        
        # Save the Excel workbook
        excel_path = f"{output_dir}/Attrition_Report_{report_time}.xlsx"
        wb.save(excel_path)
        logger.info(f"Excel report saved to {excel_path}")
        logger.info(f"Excel report generated with {successful_sections} of {len(section_functions)} sections completed successfully")
        return True, excel_path, report_time
    
    except Exception as e:
        logger.error(f"Failed to create Excel report: {e}")
        return False, None, None
