from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
from datetime import datetime
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import os
from sklearn.model_selection import train_test_split
from sklearn.ensemble import RandomForestClassifier
from sklearn.preprocessing import LabelEncoder

def load_data(file_path: str, logger):
    """
    Load and prepare the data
    
    Args:
        file_path: Path to the Excel file
        logger: Logger instance for logging operations
        
    Returns:
        DataFrame if successful, None if error occurs
    """
    try:
        df = pd.read_excel(file_path)
        
        # Ensure all required columns exist
        ensure_required_columns(df)
        
        # Fill missing values in Action Type column
        df['Action Type'] = df['Action Type'].fillna('')
        
        logger.info(f"Successfully loaded data from {file_path}")
        return df
    except Exception as e:
        logger.error(f"Error loading data: {e}")
        return None

def style_table_headers(table, header_color="ADD8E6"):
    """Apply styling to the header row of a table"""
    try:
        header_cells = table.rows[0].cells
        for cell in header_cells:
            if cell.paragraphs and cell.paragraphs[0].runs:
                cell.paragraphs[0].runs[0].bold = True
            cell.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
            shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{header_color}"/>')
            cell._tc.get_or_add_tcPr().append(shading_elm)
        return True
    except Exception as e:
        return False

def create_predictive_report_document(output_dir):
    """Create and initialize the Word document for the predictive report"""
    try:
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
        
        doc = Document()
        doc.add_heading('PREDICTIVE ANALYTICS REPORT: ATTRITION RISK BY VIMAL SINGH', 0)
        doc.add_paragraph(f"Report Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        
        return doc, datetime.now().strftime("%Y%m%d_%H%M%S")
    except Exception as e:
        return None, None

def add_predictive_analytics(doc, df, output_dir, report_time, logger):
    """Add predictive analytics section listing employees with attrition risk > 70%"""
    try:
        doc.add_heading('1. Predictive Analytics for Attrition Risk', level=1)
        doc.add_paragraph("This section lists active employees with a predicted attrition probability greater than 70%, sorted in descending order by probability, based on a Random Forest Classifier model trained on historical employee data.")
        
        # Prepare data for training
        df['Is_Separation'] = df['Action Type'].str.contains('Exit|Resignation|Termination|Separation', na=False).astype(int)
        
        # Calculate tenure in years
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
        
        if active_df.empty:
            logger.warning("No active employees found for prediction.")
            doc.add_paragraph("No active employees available for attrition risk prediction.")
            return None, True
        
        # Prepare active employee features
        X_active = active_df[features].copy()
        for col in X_active.select_dtypes(include=['object']).columns:
            X_active[col] = X_active[col].astype(str).fillna('Unknown')
            if col in label_encoders:
                # Handle unseen categories
                le = label_encoders[col]
                X_active[col] = X_active[col].apply(lambda x: le.transform([x])[0] if x in le.classes_ else 0)
        
        X_active = X_active.fillna(0)
        
        # Predict probabilities
        active_df['Attrition_Probability'] = model.predict_proba(X_active)[:, 1]
        
        # Filter high-risk employees
        high_risk_employees = active_df[active_df['Attrition_Probability'] > 0.7].copy()
        
        if not high_risk_employees.empty:
            high_risk_employees = high_risk_employees.sort_values(by='Attrition_Probability', ascending=False)
            
            # Add summary statistics
            total_high_risk = len(high_risk_employees)
            doc.add_paragraph(f"Total Number of High-Risk Employees: {total_high_risk}")
            
            # Show top 50 for main display
            high_risk_display = high_risk_employees.head(50)
            
            doc.add_paragraph("Top 50 Employees with High Attrition Risk (Probability > 0.7):")
            table = doc.add_table(rows=1, cols=4)
            table.style = 'Table Grid'
            header_cells = table.rows[0].cells
            header_cells[0].text = 'Employee Name'
            header_cells[1].text = 'Function'
            header_cells[2].text = 'Job Location'
            header_cells[3].text = 'Attrition Probability'
            style_table_headers(table)
            
            for _, row in high_risk_display.iterrows():
                row_cells = table.add_row().cells
                row_cells[0].text = str(row['Employee Name'])
                row_cells[1].text = str(row['Function'])
                row_cells[2].text = str(row['Job Location'])
                row_cells[3].text = f"{row['Attrition_Probability']:.3f}"
        else:
            logger.warning("No high-risk active employees found.")
            doc.add_paragraph("No active employees with attrition probability > 0.7.")
            high_risk_employees = pd.DataFrame()
        
        doc.add_page_break()
        logger.info("Successfully added predictive analytics section")
        return high_risk_employees, True
    except Exception as e:
        logger.error(f"Failed to add predictive analytics section: {e}")
        doc.add_paragraph("Error generating Predictive Analytics section.")
        doc.add_page_break()
        return None, False

def add_gender_analysis(doc, high_risk_df, output_dir, report_time, logger):
    """Add gender-wise analysis of high-risk employees"""
    try:
        doc.add_heading('2. Gender-wise Attrition Risk Analysis', level=1)
        
        if high_risk_df is None or high_risk_df.empty:
            doc.add_paragraph("No high-risk employees available for gender analysis.")
            return True
        
        # Group by Gender
        attrition_by_gender = high_risk_df.groupby('Gender').agg({
            'Employee Name': 'count'
        }).rename(columns={'Employee Name': 'High-Risk Count'})
        
        # Add table
        table = doc.add_table(rows=1, cols=2)
        table.style = 'Table Grid'
        header_cells = table.rows[0].cells
        header_cells[0].text = 'Gender'
        header_cells[1].text = 'High-Risk Count'
        style_table_headers(table)
        
        for index, row in attrition_by_gender.reset_index().iterrows():
            row_cells = table.add_row().cells
            row_cells[0].text = str(row['Gender'])
            row_cells[1].text = str(row['High-Risk Count'])
        
        # Create and save plot
        plt.figure(figsize=(8, 6))
        sns.barplot(x=attrition_by_gender.index, y=attrition_by_gender['High-Risk Count'], 
                   hue=attrition_by_gender.index, palette="coolwarm")
        plt.title('High-Risk Count by Gender', fontsize=15, fontweight='bold')
        plt.ylabel('Number of Employees', fontsize=14)
        plt.xlabel('Gender', fontsize=14)
        plt.tight_layout()
        
        gender_plot_path = f"{output_dir}/Predictive_Gender_Analysis_{report_time}.png"
        plt.savefig(gender_plot_path)
        plt.close()

        doc.add_paragraph("\n")
        doc.add_picture(gender_plot_path, width=Inches(5))
        last_paragraph = doc.paragraphs[-1]
        last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_page_break()
        
        logger.info("Successfully added gender analysis section")
        return True
    except Exception as e:
        logger.error(f"Failed to add gender analysis section: {e}")
        doc.add_paragraph("Error generating Gender Analysis section.")
        doc.add_page_break()
        return False

def add_complete_high_risk_list(doc, high_risk_df, logger):
    """Add complete list of high-risk employees in descending order"""
    try:
        doc.add_heading('7. Complete List of High-Risk Employees', level=1)
        doc.add_paragraph("This section provides the complete list of active employees with a predicted attrition probability greater than 70%, sorted in descending order by probability.")
        
        if high_risk_df is None or high_risk_df.empty:
            doc.add_paragraph("No high-risk employees available.")
            return True
        
        # Sort by Attrition_Probability in descending order
        high_risk_df = high_risk_df.sort_values(by='Attrition_Probability', ascending=False)
        
        # Add table
        table = doc.add_table(rows=1, cols=4)
        table.style = 'Table Grid'
        header_cells = table.rows[0].cells
        header_cells[0].text = 'Employee Name'
        header_cells[1].text = 'Function'
        header_cells[2].text = 'Job Location'
        header_cells[3].text = 'Attrition Probability'
        style_table_headers(table)
        
        for _, row in high_risk_df.iterrows():
            row_cells = table.add_row().cells
            row_cells[0].text = str(row['Employee Name'])
            row_cells[1].text = str(row['Function'])
            row_cells[2].text = str(row['Job Location'])
            row_cells[3].text = f"{row['Attrition_Probability']:.3f}"
        
        doc.add_page_break()
        logger.info("Successfully added complete high-risk list section")
        return True
    except Exception as e:
        logger.error(f"Failed to add complete high-risk list section: {e}")
        doc.add_paragraph("Error generating Complete High-Risk List section.")
        doc.add_page_break()
        return False

def generate_predictive_attrition_report(df, output_dir, logger):
    """Create a comprehensive predictive analytics report"""
    try:
        # Create document
        doc, report_time = create_predictive_report_document(output_dir)
        
        if doc is None:
            logger.error("Failed to create report document. Aborting report generation.")
            return False, None, None
        
        # Generate predictive analytics section
        high_risk_df, success = add_predictive_analytics(doc, df, output_dir, report_time, logger)
        
        if success and high_risk_df is not None and not high_risk_df.empty:
            # Add gender analysis
            add_gender_analysis(doc, high_risk_df, output_dir, report_time, logger)
            
            # Add complete list
            add_complete_high_risk_list(doc, high_risk_df, logger)
        
        # Save the document
        report_path = f"{output_dir}/Predictive_Attrition_Report_{report_time}.docx"
        doc.save(report_path)
        logger.info(f"Predictive report saved to {report_path}")
        
        return True, report_path, report_time
        
    except Exception as e:
        logger.error(f"Failed to generate predictive report: {e}")
        return False, None, None
    
    