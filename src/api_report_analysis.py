# Import necessary libraries
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from datetime import datetime
import os
from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import matplotlib.gridspec as gridspec
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from fastapi import Body, FastAPI, Response, status, UploadFile, File, BackgroundTasks
from fastapi import HTTPException
from fastapi.responses import FileResponse
from pydantic import BaseModel, ValidationError
from typing import Optional
from fastapi.middleware.cors import CORSMiddleware
import shutil
import uuid
from pathlib import Path
from src.logging_config import setup_logging
from src.cors_config import setup_cors

# Assign the logger instance returned by setup_logging to a variable
logger = setup_logging()

# Create the FastAPI app
app = FastAPI(title="Attrition Analysis API", 
             description="API for generating automated attrition analysis reports",
             version="1.0.0")

# Add CORS middleware with fixed syntax
setup_cors(app)

# Define base directories
BASE_DIR = Path(__file__).parent.parent
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_DIR = os.path.join(SCRIPT_DIR, "file_uploads")
OUTPUT_DIR = os.path.join(SCRIPT_DIR, "attrition_reports")

# Log the directory paths
logger.info(f"UPLOAD_DIR: {UPLOAD_DIR}")
logger.info(f"OUTPUT_DIR: {OUTPUT_DIR}")

# Create directories if they don't exist
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

# Debug middleware to log requests
@app.middleware("http")
async def debug_request(request, call_next):
    print(f"Request: {request.method} {request.url}")
    print(f"Headers: {request.headers}")
    
    response = await call_next(request)
    print(f"Response: {response.status_code}")
    return response



# Helper functions for column validation
def ensure_employee_name_column(df):
    """
    Ensures 'Employee Name' column exists, creating it from other ID columns if needed.
    
    Args:
        df: pandas DataFrame containing employee data
        
    Returns:
        None (modifies DataFrame in place)
        
    Raises:
        ValueError: If no suitable column for 'Employee Name' is found
    """
    # Do nothing if 'Employee Name' already exists
    if 'Employee Name' in df.columns:
        return
    
    # Look for alternative ID columns
    id_columns = ['Employee ID', 'EmployeeID', 'Emp ID', 'EmpID', 'ID']
    
    # Use the first ID column that exists
    for col in id_columns:
        if col in df.columns:
            df['Employee Name'] = df[col]
            return
    
    # If we get here, no suitable column was found
    raise ValueError("No 'Employee Name' or ID column found in the data")


def ensure_action_type_column(df):
    """
    Ensures 'Action Type' column exists, creating it from other status columns if needed.
    
    Args:
        df: pandas DataFrame containing employee data
        
    Returns:
        None (modifies DataFrame in place)
        
    Raises:
        TypeError: If df is not a pandas DataFrame
        ValueError: If the DataFrame is empty
    """
    # Basic validation
    if not hasattr(df, 'columns'):
        raise TypeError("Input must be a pandas DataFrame")
    
    if df.empty:
        raise ValueError("DataFrame is empty")
        
    # Initialize or clean Action Type column
    if 'Action Type' in df.columns:
        df['Action Type'] = df['Action Type'].fillna('')
        return
    
    # Look for alternative status columns
    status_columns = ['Status', 'Employee Status', 'Employment Status', 
                      'Job Status', 'Action', 'Termination Status', 'Exit Status']
    
    # Use the first status column that exists
    for col in status_columns:
        if col in df.columns:
            df['Action Type'] = df[col]
            df['Action Type'] = df['Action Type'].fillna('')
            return
    
    # If no status column found, raise an error
    raise ValueError("No 'Action Type' or status-related column found in the data")


def ensure_gender_column(df):
    """
    Ensures 'Gender' column exists, creating it from 'Sex' if needed.
    """
    # Do nothing if 'Gender' already exists
    if 'Gender' in df.columns:
        return
    
    # Check for 'Sex' column as alternative
    if 'Sex' in df.columns:
        df['Gender'] = df['Sex']
        return
    
    # If neither column exists, create a default Gender column
    #df['Gender'] = "Unknown"



def ensure_function_column(df):
    """
    Ensures 'Function' column exists, creating it from Department or Business Unit if needed.
    """
    # Do nothing if 'Function' already exists
    if 'Function' in df.columns:
        return
    
    # Check for 'Department' column as alternative
    if 'Department' in df.columns:
        df['Function'] = df['Department']
        return
    
    # Check for 'Business Unit' column as alternative
    if 'Business Unit' in df.columns:
        df['Function'] = df['Business Unit']
        return
    
    # If none of the columns exist, create a default Function column
    # df['Function'] = "Unknown"


def ensure_grade_column(df):
    """
    Ensures 'Grade' column exists, creating it from alternative column names if needed.
    """
    # Do nothing if 'Grade' already exists
    if 'Grade' in df.columns:
        return
    
    # List of alternative column names for Grade
    grade_alternatives = [
        "Job Grade", "Pay Grade", "Grade Level", "Salary Grade", 
        "Grade Code", "Band", "Position Grade", "Compensation Grade", 
        "Organizational Level", "Rank", "Position Rank"
    ]
    
    # Check each alternative and use the first one found
    for alt_column in grade_alternatives:
        if alt_column in df.columns:
            df['Grade'] = df[alt_column]
            return
    
    # No need to add default values if no alternatives found


##New validation code ends

def load_data(file_path):
    """
    Load and prepare the data
    """
    try:
        df = pd.read_excel(file_path)

        # Add the new function calls here
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

def style_table_headers(table, header_color="ADD8E6"):
    """
    Apply styling to the header row of a table
    
    Parameters:
    - table: the table object to style
    - header_color: hex color code for header background (default: light blue)
    """
    try:
        header_cells = table.rows[0].cells
        
        for cell in header_cells:
            # Make text bold
            cell.paragraphs[0].runs[0].bold = True
            # Center text
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            # Apply background color
            shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{header_color}"/>')
            cell._tc.get_or_add_tcPr().append(shading_elm)
        return True
    except Exception as e:
        logger.warning(f"Failed to style table headers: {e}")
        return False

def create_report_document(output_dir):
    """
    Create and initialize the Word document for the report
    """
    try:
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
        
        # Create Word document
        doc = Document()
        
        # Add title
        doc.add_heading('AUTOMATED ATTRITION ANALYSIS REPORT BY VIMAL SINGH', 0)
        doc.add_paragraph(f"Report Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        
        return doc, datetime.now().strftime("%Y%m%d_%H%M%S")
    except Exception as e:
        logger.error(f"Failed to create report document: {e}")
        return None, None

def add_overall_statistics(doc, df, output_dir, report_time):
    """
    Add overall attrition statistics to the report
    """
    try:
        doc.add_heading('1. Overall Attrition Statistics', level=1)
        
        # Calculate total employees and exits
        total_employees = len(df)
        exits = df[df['Action Type'].str.contains('Exit|Resignation|Termination', na=False)]
        total_exits = len(exits)
        attrition_rate = (total_exits / total_employees * 100).round(2)
        
        # Add overall statistics paragraph
        doc.add_paragraph(f"Total Employee Count: {total_employees}")
        doc.add_paragraph(f"Total Exits: {total_exits}")
        doc.add_paragraph(f"Overall Attrition Rate: {attrition_rate}%")
        
        # Create overall pie chart
        plt.figure(figsize=(10, 6))
        labels = ['Active Employees', 'Exited Employees']
        values = [total_employees - total_exits, total_exits]
        plt.pie(values, labels=labels, autopct='%1.1f%%', startangle=90, explode=[0, 0.1])
        plt.title('Overall Attrition', fontsize=16, fontweight='bold')
        plt.tight_layout()
        
        overall_plot_path = f"{output_dir}/Overall_Statistics_{report_time}.png"
        plt.savefig(overall_plot_path)
        plt.close()
        
        doc.add_picture(overall_plot_path, width=Inches(6))
        last_paragraph = doc.paragraphs[-1]
        last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        doc.add_page_break()
        
        logger.info("Successfully added overall statistics section")
        return True
    except Exception as e:
        logger.error(f"Failed to add overall statistics section: {e}")
        doc.add_paragraph("Error generating Overall Statistics section. Section skipped.")
        doc.add_page_break()
        return False

def add_gender_analysis(doc, df, output_dir, report_time):
    """
    Add gender-wise attrition analysis to the report
    """
    try:
        doc.add_heading('2. Gender-wise Attrition', level=1)
        attrition_by_gender = df[df['Action Type'].str.contains('Exit|Resignation|Termination', na=False)].groupby('Gender').agg({
            'Employee Name': 'count'
        }).rename(columns={'Employee Name': 'Attrition Count'})
        
        total_by_gender = df.groupby('Gender')['Employee Name'].count()
        attrition_by_gender['Total Employees'] = total_by_gender
        attrition_by_gender['Attrition Rate %'] = (attrition_by_gender['Attrition Count'] / attrition_by_gender['Total Employees'] * 100).round(2)
        
        # Add gender data table
        table = doc.add_table(rows=1, cols=4)
        table.style = 'Table Grid'
        header_cells = table.rows[0].cells
        header_cells[0].text = 'Gender'
        header_cells[1].text = 'Attrition Count'
        header_cells[2].text = 'Total Employees'
        header_cells[3].text = 'Attrition Rate %'

        style_table_headers(table)
        
        for index, row in attrition_by_gender.reset_index().iterrows():
            row_cells = table.add_row().cells
            row_cells[0].text = str(row['Gender'])
            row_cells[1].text = str(row['Attrition Count'])
            row_cells[2].text = str(row['Total Employees'])
            row_cells[3].text = f"{row['Attrition Rate %']}%"
        
        # Create and save gender plots
        fig = plt.figure(figsize=(16, 7))
        gs = gridspec.GridSpec(1, 2, width_ratios=[2, 3])  # ax1 will be smaller than ax2
        ax1 = plt.subplot(gs[0])  # Smaller subplot
        ax2 = plt.subplot(gs[1])  # Larger subplot

        sns.barplot(x=attrition_by_gender.index, y=attrition_by_gender['Attrition Count'], ax=ax1, hue=attrition_by_gender.index, palette="coolwarm", width=0.4)
        ax1.set_title('Attrition Count by Gender\n', fontsize=15, fontweight='bold', color='black')
        ax1.set_ylabel('Number of Employees', fontsize=14)
        ax1.set_xlabel('Gender', fontsize=14)
        
        ax2.pie(attrition_by_gender['Attrition Rate %'], labels=attrition_by_gender.index, autopct='%1.1f%%', textprops={'fontsize': 14})
        ax2.set_title('Attrition Rate Distribution by Gender\n', fontsize=15, fontweight='bold', color='black')
        plt.tight_layout()
        
        gender_plot_path = f"{output_dir}/Gender_Analysis_{report_time}.png"
        plt.savefig(gender_plot_path)
        plt.close()

        doc.add_paragraph("\n")
        doc.add_picture(gender_plot_path, width=Inches(7))
        
        last_paragraph = doc.paragraphs[-1]
        last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_page_break()
        
        logger.info("Successfully added gender analysis section")
        return True
    except Exception as e:
        logger.error(f"Failed to add gender analysis section: {e}")
        doc.add_paragraph("Error generating Gender Analysis section. Section skipped.")
        doc.add_page_break()
        return False

def add_location_analysis(doc, df, output_dir, report_time):
    """
    Add location-wise attrition analysis to the report
    """
    try:
        doc.add_heading('3. Location-wise Attrition (Locations with ≥50 Employees)', level=1)
        
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
        attrition_by_location['Attrition Rate %'] = (attrition_by_location['Attrition Count'] / attrition_by_location['Total Employees'] * 100).round(2)
        
        # Add location data table
        table = doc.add_table(rows=1, cols=4)
        table.style = 'Table Grid'
        header_cells = table.rows[0].cells
        header_cells[0].text = 'Location'
        header_cells[1].text = 'Attrition Count'
        header_cells[2].text = 'Total Employees'
        header_cells[3].text = 'Attrition Rate %'

        style_table_headers(table)
        
        for index, row in attrition_by_location.reset_index().iterrows():
            row_cells = table.add_row().cells
            row_cells[0].text = str(row['Job Location'])
            row_cells[1].text = str(row['Attrition Count'])
            row_cells[2].text = str(row['Total Employees'])
            row_cells[3].text = f"{row['Attrition Rate %']}%"

        doc.add_paragraph("\n")
        
        # Create and save location plots
        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(16, 7))
        sns.barplot(data=attrition_by_location.reset_index(), x='Job Location', y='Attrition Count', ax=ax1)
        ax1.set_title('Attrition Count by Location\n', fontsize=16, fontweight='bold')
        ax1.set_xlabel('Location', fontsize=14)
        ax1.set_ylabel('Number of Employees', fontsize=14)
        ax1.tick_params(axis='x', rotation=45)
        
        ax2.pie(attrition_by_location['Attrition Rate %'], labels=attrition_by_location.index, autopct='%1.1f%%', textprops={'fontsize': 14})
        ax2.set_title('Attrition Rate Distribution by Location\n', fontsize=16, fontweight='bold')
        plt.tight_layout()
        
        location_plot_path = f"{output_dir}/Location_Analysis_{report_time}.png"
        plt.savefig(location_plot_path, bbox_inches='tight')
        plt.close()
        
        doc.add_picture(location_plot_path, width=Inches(7))
        last_paragraph = doc.paragraphs[-1]
        last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

        doc.add_page_break()
        
        logger.info("Successfully added location analysis section")
        return True
    except Exception as e:
        logger.error(f"Failed to add location analysis section: {e}")
        doc.add_paragraph("Error generating Location Analysis section. Section skipped.")
        doc.add_page_break()
        return False

def add_function_analysis(doc, df, output_dir, report_time):
    """
    Add function-wise attrition analysis to the report
    """
    try:
        doc.add_heading('4. Function-wise Attrition (Functions with ≥20 Employees)', level=1)
        
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
        attrition_by_function['Attrition Rate %'] = (attrition_by_function['Attrition Count'] / attrition_by_function['Total Employees'] * 100).round(2)
        
        # Add function data table
        table = doc.add_table(rows=1, cols=4)
        table.style = 'Table Grid'
        header_cells = table.rows[0].cells
        header_cells[0].text = 'Function'
        header_cells[1].text = 'Attrition Count'
        header_cells[2].text = 'Total Employees'
        header_cells[3].text = 'Attrition Rate %'

        style_table_headers(table)
        
        for index, row in attrition_by_function.reset_index().iterrows():
            row_cells = table.add_row().cells
            row_cells[0].text = str(row['Function'])
            row_cells[1].text = str(row['Attrition Count'])
            row_cells[2].text = str(row['Total Employees'])
            row_cells[3].text = f"{row['Attrition Rate %']}%"

        doc.add_paragraph("\n")
        
        # Create and save function plots
        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(17, 7))
        sns.barplot(data=attrition_by_function.reset_index(), x='Function', y='Attrition Count', hue='Function', ax=ax1, legend=False)
        ax1.set_title('Attrition Count by Function\n', fontsize=16, fontweight='bold')
        ax1.set_ylabel('Number of Employees', fontsize=14)
        ax1.set_xlabel('Function', fontsize=14)
        ax1.tick_params(axis='x', rotation=45)
        
        ax2.pie(attrition_by_function['Attrition Rate %'], labels=attrition_by_function.index, autopct='%1.1f%%', textprops={'fontsize': 14})
        ax2.set_title('Attrition Rate Distribution by Function\n', fontsize=16, fontweight='bold')
        plt.tight_layout()
        
        function_plot_path = f"{output_dir}/Function_Analysis_{report_time}.png"
        plt.savefig(function_plot_path, bbox_inches='tight')
        plt.close()
        
        doc.add_picture(function_plot_path, width=Inches(7))
        last_paragraph = doc.paragraphs[-1]
        last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

        doc.add_page_break()
        
        logger.info("Successfully added function analysis section")
        return True
    except Exception as e:
        logger.error(f"Failed to add function analysis section: {e}")
        doc.add_paragraph("Error generating Function Analysis section. Section skipped.")
        doc.add_page_break()
        return False

def add_tenure_analysis(doc, df, output_dir, report_time):
    """
    Add tenure analysis of exited employees to the report
    """
    try:
        doc.add_heading('5. Tenure Analysis of Exited Employees', level=1)

        # Filter to only include employees who have exited
        exits_df = df[df['Action Type'].str.contains('Exit|Resignation|Termination', na=False)].copy()

        # Calculate tenure in years (difference between joining date and action date)
        exits_df['Tenure in Days'] = (pd.to_datetime(exits_df['Action Date']) - pd.to_datetime(exits_df['Date of Joining'])).dt.days
        exits_df['Tenure in Years'] = exits_df['Tenure in Days'] / 365.25

        # Create tenure bins
        tenure_bins = [0, 1, 2, 3, 5, 10, float('inf')]
        tenure_labels = ['<1 year', '1-2 years', '2-3 years', '3-5 years', '5-10 years', '>10 years']
        exits_df['Tenure Band'] = pd.cut(exits_df['Tenure in Years'], bins=tenure_bins, labels=tenure_labels, right=False)

        # Group by tenure band - Fix for FutureWarning
        tenure_analysis = exits_df.groupby('Tenure Band', observed=True).size().reset_index(name='Count')
        tenure_analysis['Percentage'] = (tenure_analysis['Count'] / tenure_analysis['Count'].sum() * 100).round(2)

        # Add tenure data table
        table = doc.add_table(rows=1, cols=3)
        table.style = 'Table Grid'
        header_cells = table.rows[0].cells
        header_cells[0].text = 'Tenure Band'
        header_cells[1].text = 'Number of Exits'
        header_cells[2].text = 'Percentage'

        style_table_headers(table)

        for index, row in tenure_analysis.iterrows():
            row_cells = table.add_row().cells
            row_cells[0].text = str(row['Tenure Band'])
            row_cells[1].text = str(row['Count'])
            row_cells[2].text = f"{row['Percentage']}%"

        doc.add_paragraph("\n")
        
        # Create and save tenure plots
        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(16, 7))

        # Bar chart - Fix for seaborn FutureWarning
        sns.barplot(data=tenure_analysis, x='Tenure Band', y='Count', hue='Tenure Band', ax=ax1, legend=False)
        ax1.set_title('Attrition Count by Tenure\n', fontsize=16, fontweight='bold')
        ax1.set_ylabel('Number of Employees', fontsize=14)
        ax1.set_xlabel('Tenure Band', fontsize=14)
        ax1.tick_params(axis='x', rotation=45)

        # Pie chart
        ax2.pie(tenure_analysis['Percentage'], labels=tenure_analysis['Tenure Band'], autopct='%1.1f%%', textprops={'fontsize': 14})
        ax2.set_title('Attrition Distribution by Tenure\n', fontsize=16, fontweight='bold')
        plt.tight_layout()

        tenure_plot_path = f"{output_dir}/Tenure_Analysis_{report_time}.png"
        plt.savefig(tenure_plot_path, bbox_inches='tight')
        plt.close()

        doc.add_picture(tenure_plot_path, width=Inches(7))
        last_paragraph = doc.paragraphs[-1]
        last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        doc.add_page_break()
        
        logger.info("Successfully added tenure analysis section")
        return True
    except Exception as e:
        logger.error(f"Failed to add tenure analysis section: {e}")
        doc.add_paragraph("Error generating Tenure Analysis section. Section skipped.")
        doc.add_page_break()
        return False

def add_grade_analysis(doc, df, output_dir, report_time):
    """
    Add grade-wise attrition analysis to the report
    """
    try:
        doc.add_heading('6. Grade-wise Attrition', level=1)

        # Get counts of employees by grade
        total_by_grade = df.groupby('Grade')['Employee Name'].count()
        valid_grades = total_by_grade[total_by_grade >= 10].index  # Only include grades with at least 10 employees

        # Filter to include only valid grades and exit actions
        attrition_by_grade = df[
            (df['Action Type'].str.contains('Exit|Resignation|Termination', na=False)) & 
            (df['Grade'].isin(valid_grades))
        ].groupby('Grade').agg({
            'Employee Name': 'count'
        }).rename(columns={'Employee Name': 'Attrition Count'})

        # Calculate attrition rates
        total_by_grade = total_by_grade[valid_grades]
        attrition_by_grade['Total Employees'] = total_by_grade
        attrition_by_grade['Attrition Rate %'] = (attrition_by_grade['Attrition Count'] / attrition_by_grade['Total Employees'] * 100).round(2)

        # Add grade data table
        table = doc.add_table(rows=1, cols=4)
        table.style = 'Table Grid'
        header_cells = table.rows[0].cells
        header_cells[0].text = 'Grade'
        header_cells[1].text = 'Attrition Count'
        header_cells[2].text = 'Total Employees'
        header_cells[3].text = 'Attrition Rate %'

        style_table_headers(table)
        
        for index, row in attrition_by_grade.reset_index().iterrows():
            row_cells = table.add_row().cells
            row_cells[0].text = str(row['Grade'])
            row_cells[1].text = str(row['Attrition Count'])
            row_cells[2].text = str(row['Total Employees'])
            row_cells[3].text = f"{row['Attrition Rate %']}%"

        doc.add_paragraph("\n")
        
        # Create and save grade plots
        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(16, 7))

        # Bar chart
        sns.barplot(data=attrition_by_grade.reset_index(), x='Grade', y='Attrition Count', hue='Grade', ax=ax1, legend=False)
        ax1.set_title('Attrition Count by Grade\n', fontsize=16, fontweight='bold')
        ax1.set_ylabel('Number of Employees', fontsize=14)
        ax1.set_xlabel('Grade', fontsize=14)
        ax1.tick_params(axis='x', rotation=45)

        # Pie chart
        ax2.pie(attrition_by_grade['Attrition Rate %'], labels=attrition_by_grade.index, autopct='%1.1f%%', textprops={'fontsize': 14})
        ax2.set_title('Attrition Rate Distribution by Grade\n', fontsize=16, fontweight='bold')
        plt.tight_layout()

        grade_plot_path = f"{output_dir}/Grade_Analysis_{report_time}.png"
        plt.savefig(grade_plot_path, bbox_inches='tight')
        plt.close()

        doc.add_picture(grade_plot_path, width=Inches(7))
        last_paragraph = doc.paragraphs[-1]
        last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

        doc.add_page_break()
        
        logger.info("Successfully added grade analysis section")
        return True
    except Exception as e:
        logger.error(f"Failed to add grade analysis section: {e}")
        doc.add_paragraph("Error generating Grade Analysis section. Section skipped.")
        doc.add_page_break()
        return False

def add_trend_analysis(doc, df, output_dir, report_time):
    """
    Add quarterly and monthly attrition trend analysis to the report
    """
    try:
        doc.add_heading('7. Quarterly and Monthly Attrition Trends', level=1)

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
        quarterly_trends['Quarter'] = pd.PeriodIndex(quarterly_trends['Quarter'], freq='Q')

        # Calculate monthly trends
        monthly_trends = exits_df.groupby('Month').size().reset_index(name='Exit Count')
        monthly_trends = monthly_trends.sort_values('Month')  # Sort by month chronologically
        monthly_trends['Month'] = pd.to_datetime(monthly_trends['Month'])

        # Add quarterly trend table first
        doc.add_heading('Quarterly Attrition Trend', level=2)
        table = doc.add_table(rows=1, cols=2)
        table.style = 'Table Grid'
        header_cells = table.rows[0].cells
        header_cells[0].text = 'Quarter'
        header_cells[1].text = 'Exit Count'

        style_table_headers(table)

        for index, row in quarterly_trends.iterrows():
            row_cells = table.add_row().cells
            row_cells[0].text = str(row['Quarter'])
            row_cells[1].text = str(row['Exit Count'])

        doc.add_paragraph("\n")
        
        # Create and save quarterly trend chart
        plt.figure(figsize=(10, 6))
        ax = sns.barplot(data=quarterly_trends, x=quarterly_trends['Quarter'].astype(str), y='Exit Count', hue=quarterly_trends['Quarter'].astype(str), legend=False)
    
        plt.title('Quarterly Attrition Trend\n', fontsize=16, fontweight='bold')
        plt.ylabel('Number of Exits', fontsize=14)
        plt.xlabel('Quarter', fontsize=14)
        plt.grid(True, linestyle='--', alpha=0.7, axis='y')
        plt.tight_layout()

        # Add count labels on top of each bar
        for i, v in enumerate(quarterly_trends['Exit Count']):
            ax.text(i, v + 0.5, str(v), ha='center', fontweight='bold')

        quarterly_trend_path = f"{output_dir}/Quarterly_Trend_{report_time}.png"
        plt.savefig(quarterly_trend_path, bbox_inches='tight')
        plt.close()

        doc.add_picture(quarterly_trend_path, width=Inches(7))
        last_paragraph = doc.paragraphs[-1]
        last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # Now add monthly trend table
        doc.add_heading('Monthly Attrition Trend', level=2)
        table = doc.add_table(rows=1, cols=2)
        table.style = 'Table Grid'
        header_cells = table.rows[0].cells
        header_cells[0].text = 'Month'
        header_cells[1].text = 'Exit Count'

        style_table_headers(table)
        
        for index, row in monthly_trends.iterrows():
            row_cells = table.add_row().cells
            row_cells[0].text = str(row['Month'])
            row_cells[1].text = str(row['Exit Count'])

        doc.add_paragraph("\n")

        # Create and save monthly trend chart
        plt.figure(figsize=(12, 6))
        sns.lineplot(data=monthly_trends, x='Month', y='Exit Count', marker='o', linewidth=2)
        plt.title('Monthly Attrition Trend\n', fontsize=16, fontweight='bold')
        plt.ylabel('Number of Exits', fontsize=14)
        plt.xlabel('Month', fontsize=14)
        plt.xticks(rotation=45)
        plt.grid(True, linestyle='--', alpha=0.7)
        plt.tight_layout()

        monthly_trend_path = f"{output_dir}/Monthly_Trend_{report_time}.png"
        plt.savefig(monthly_trend_path, bbox_inches='tight')
        plt.close()

        doc.add_picture(monthly_trend_path, width=Inches(7))
        last_paragraph = doc.paragraphs[-1]
        last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        logger.info("Successfully added trend analysis section")
        return True
    except Exception as e:
        logger.error(f"Failed to add trend analysis section: {e}")
        doc.add_paragraph("Error generating Trend Analysis section. Section skipped.")
        return False

def create_attrition_report(df, output_dir):
    """
    Create a comprehensive automated attrition analysis report in Word format
    with modular sections and error handling
    """
    # Create document
    doc, report_time = create_report_document(output_dir)
    today= datetime.now().strftime("%Y-%m-%d");
    time = datetime.now().strftime("%H:%M");

    if doc is None:
        logger.error("Failed to create report document. Aborting report generation.")
        return False
    
    # Define all report sections to run
    report_sections = [
        {"func": add_overall_statistics, "name": "Overall Statistics"},
        {"func": add_gender_analysis, "name": "Gender Analysis"},
        {"func": add_location_analysis, "name": "Location Analysis"},
        {"func": add_function_analysis, "name": "Function Analysis"},
        {"func": add_tenure_analysis, "name": "Tenure Analysis"},
        {"func": add_grade_analysis, "name": "Grade Analysis"},
        {"func": add_trend_analysis, "name": "Trend Analysis"}
    ]
    
    # Execute each section, skipping any that fail
    successful_sections = 0
    
    for section in report_sections:
        section_func = section["func"]
        section_name = section["name"]
        
        logger.info(f"Generating {section_name} section...")
        
        # Call all section functions with the same 4 parameters
        try:
            result = section_func(doc, df, output_dir, report_time)
            if result:
                successful_sections += 1
            else:
                logger.warning(f"Section function returned False: {section_name}")
        except Exception as e:
            logger.error(f"Error in {section_name} section: {e}")
            doc.add_paragraph(f"Error generating {section_name} section. Section skipped.")
            doc.add_page_break()
    

    # Save the Word document
    try:
        report_path = f"{output_dir}/Attrition_Report_{today}_{time}.docx"
        doc.save(report_path)
        logger.info(f"Report saved to {report_path}")
        logger.info(f"Report generated with {successful_sections} of {len(report_sections)} sections completed successfully")
        return True, report_path, report_time
    except Exception as e:
        logger.error(f"Failed to save report: {e}")
        return False, None, None

# Root endpoint
@app.get("/")
def read_root():
    return {"message": "Welcome to the Attrition Analysis API"}

# File upload endpoint
@app.post("/api/upload/")
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

# Generate report endpoint
@app.post("/api/generate-report/")
def generate_report(file_id: str, background_tasks: BackgroundTasks):
    """
    Generate an attrition report for the uploaded file
    """
    try:
        # Find the file with the given ID
        files = [f for f in os.listdir(UPLOAD_DIR) if f.startswith(f"{file_id}_")]
        
        if not files:
            raise HTTPException(status_code=404, detail="File not found. Please upload the file first.")
        
        file_path = f"{UPLOAD_DIR}/{files[0]}"
        report_dir = f"{OUTPUT_DIR}/{file_id}"
        os.makedirs(report_dir, exist_ok=True)
        
        # Load data
        df = load_data(file_path)
        if df is None:
            raise HTTPException(status_code=500, detail="Failed to load data from the uploaded file.")
            
        # Generate report
        success, report_path, report_time = create_attrition_report(df, report_dir)
        
        if success and report_path:
            return {
                "status": "success",
                "message": "Report generated successfully",
                "file_id": file_id,
                "report_file": os.path.basename(report_path),
                "download_url": f"/download/{file_id}/{os.path.basename(report_path)}"
            }
        else:
            raise HTTPException(status_code=500, detail="Failed to generate report.")
            
    except Exception as e:
        logger.error(f"Report generation error: {e}")
        raise HTTPException(status_code=500, detail=f"Report generation failed: {str(e)}")

# Download report endpoint
@app.get("/api/download/")
async def download_report(file_id: str, filename: str):
    """
    Download a generated report
    """
    try:
        report_path = f"{OUTPUT_DIR}/{file_id}/{filename}"
        
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

# Main function for direct execution
def main():
    """
    Main function to run the attrition report generation
    """
    logger.info("Starting Attrition Report Generation...")
    
    INPUT_FILE = "HRIS.xlsx"  # Change this to your input file name
    OUTPUT_DIR = "attrition_reports"
    
    logger.info(f"Loading data from {INPUT_FILE}...")
    df = load_data(INPUT_FILE)
    
    if df is not None:
        logger.info("Generating report...")
        success, report_path, _ = create_attrition_report(df, OUTPUT_DIR)
        if success:
            logger.info(f"Report generated successfully. Please check the {OUTPUT_DIR} directory.")
        else:
            logger.error("Report generation completed with errors. Check log for details.")
    else:
       logger.error("Failed to generate report due to data loading error.")

# if __name__ == "__main__":
#    import uvicorn
#    uvicorn.run(app, host="127.0.0.1", port=8000)