# src/services/enhanced_predictive_generator.py
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
import numpy as np
from src.utils.validators import ensure_required_columns

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
        doc.add_heading('ENHANCED PREDICTIVE ANALYTICS REPORT: ATTRITION RISK BY VIMAL SINGH', 0)
        doc.add_paragraph(f"Report Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        
        return doc, datetime.now().strftime("%Y%m%d_%H%M%S")
    except Exception as e:
        return None, None

def build_predictive_model(df, logger):
    """Build and train the predictive model"""
    try:
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
        
        # Get model performance
        train_score = model.score(X_train, y_train)
        test_score = model.score(X_test, y_test)
        
        logger.info(f"Model trained successfully. Train accuracy: {train_score:.3f}, Test accuracy: {test_score:.3f}")
        return model, label_encoders, features, train_score, test_score
        
    except Exception as e:
        logger.error(f"Failed to build predictive model: {e}")
        return None, None, None, 0, 0

def predict_attrition_risk(df, model, label_encoders, features, logger):
    """Predict attrition risk for active employees"""
    try:
        # Filter active employees
        active_df = df[~df['Action Type'].str.contains('Exit|Resignation|Termination|Separation', na=False)].copy()
        
        if active_df.empty:
            logger.warning("No active employees found for prediction.")
            return pd.DataFrame()
        
        # Calculate tenure for active employees
        active_df['Tenure in Years'] = (pd.to_datetime(active_df['Action Date']) - pd.to_datetime(active_df['Date of Joining'])).dt.days / 365.25
        active_df['Tenure in Years'] = active_df['Tenure in Years'].fillna(0)
        
        # Prepare features
        X_active = active_df[features].copy()
        for col in X_active.select_dtypes(include=['object']).columns:
            X_active[col] = X_active[col].astype(str).fillna('Unknown')
            if col in label_encoders:
                le = label_encoders[col]
                X_active[col] = X_active[col].apply(lambda x: le.transform([x])[0] if x in le.classes_ else 0)
        
        X_active = X_active.fillna(0)
        
        # Predict probabilities
        active_df['Attrition_Probability'] = model.predict_proba(X_active)[:, 1]
        
        # Filter high-risk employees (probability > 0.7)
        high_risk_employees = active_df[active_df['Attrition_Probability'] > 0.7].copy()
        
        if not high_risk_employees.empty:
            high_risk_employees = high_risk_employees.sort_values(by='Attrition_Probability', ascending=False)
        
        logger.info(f"Found {len(high_risk_employees)} high-risk employees out of {len(active_df)} active employees")
        return high_risk_employees
        
    except Exception as e:
        logger.error(f"Failed to predict attrition risk: {e}")
        return pd.DataFrame()

def add_predictive_analytics_section(doc, df, output_dir, report_time, logger):
    """Add predictive analytics section to the report"""
    try:
        doc.add_heading('1. Enhanced Predictive Analytics for Attrition Risk', level=1)
        doc.add_paragraph("This section uses advanced machine learning algorithms to identify employees at high risk of attrition. The Random Forest Classifier model has been trained on historical employee data to predict attrition probability with enhanced accuracy.")
        
        # Build and train model
        model, label_encoders, features, train_score, test_score = build_predictive_model(df, logger)
        
        if model is None:
            doc.add_paragraph("Error: Could not build predictive model.")
            return None, False
        
        # Add model performance
        doc.add_paragraph(f"Model Performance Metrics:")
        doc.add_paragraph(f"• Training Accuracy: {train_score:.1%}")
        doc.add_paragraph(f"• Testing Accuracy: {test_score:.1%}")
        doc.add_paragraph(f"• Features Used: {', '.join(features)}")
        
        # Predict high-risk employees
        high_risk_employees = predict_attrition_risk(df, model, label_encoders, features, logger)
        
        if high_risk_employees.empty:
            doc.add_paragraph("No active employees with attrition probability > 70% were found.")
            return pd.DataFrame(), True
        
        # Add summary statistics
        total_high_risk = len(high_risk_employees)
        total_active = len(df[~df['Action Type'].str.contains('Exit|Resignation|Termination|Separation', na=False)])
        risk_percentage = (total_high_risk / total_active * 100) if total_active > 0 else 0
        
        doc.add_paragraph(f"Summary Statistics:")
        doc.add_paragraph(f"• Total Active Employees: {total_active}")
        doc.add_paragraph(f"• High-Risk Employees: {total_high_risk}")
        doc.add_paragraph(f"• Risk Percentage: {risk_percentage:.1f}%")
        
        # Show top 50 for main display
        high_risk_display = high_risk_employees.head(50)
        
        doc.add_paragraph("Top 50 Employees with Highest Attrition Risk (Probability > 0.7):")
        table = doc.add_table(rows=1, cols=5)
        table.style = 'Table Grid'
        header_cells = table.rows[0].cells
        header_cells[0].text = 'Employee Name'
        header_cells[1].text = 'Function'
        header_cells[2].text = 'Job Location'
        header_cells[3].text = 'Grade'
        header_cells[4].text = 'Risk Probability'
        style_table_headers(table)
        
        for _, row in high_risk_display.iterrows():
            row_cells = table.add_row().cells
            row_cells[0].text = str(row['Employee Name'])
            row_cells[1].text = str(row['Function'])
            row_cells[2].text = str(row['Job Location'])
            row_cells[3].text = str(row['Grade'])
            row_cells[4].text = f"{row['Attrition_Probability']:.3f}"
        
        doc.add_page_break()
        logger.info("Successfully added enhanced predictive analytics section")
        return high_risk_employees, True
        
    except Exception as e:
        logger.error(f"Failed to add predictive analytics section: {e}")
        doc.add_paragraph("Error generating Enhanced Predictive Analytics section.")
        doc.add_page_break()
        return None, False

def add_enhanced_gender_analysis(doc, high_risk_df, output_dir, report_time, logger):
    """Add enhanced gender-wise analysis of high-risk employees"""
    try:
        doc.add_heading('2. Enhanced Gender-wise Risk Analysis', level=1)
        
        if high_risk_df is None or high_risk_df.empty:
            doc.add_paragraph("No high-risk employees available for gender analysis.")
            return True
        
        # Group by Gender with additional statistics
        gender_stats = high_risk_df.groupby('Gender').agg({
            'Employee Name': 'count',
            'Attrition_Probability': ['mean', 'std', 'min', 'max']
        }).round(3)
        
        gender_stats.columns = ['Count', 'Avg_Probability', 'Std_Probability', 'Min_Probability', 'Max_Probability']
        
        # Add table
        table = doc.add_table(rows=1, cols=6)
        table.style = 'Table Grid'
        header_cells = table.rows[0].cells
        headers = ['Gender', 'Count', 'Avg Risk', 'Std Dev', 'Min Risk', 'Max Risk']
        for i, header in enumerate(headers):
            header_cells[i].text = header
        style_table_headers(table)
        
        for gender, row in gender_stats.iterrows():
            row_cells = table.add_row().cells
            row_cells[0].text = str(gender)
            row_cells[1].text = str(int(row['Count']))
            row_cells[2].text = f"{row['Avg_Probability']:.3f}"
            row_cells[3].text = f"{row['Std_Probability']:.3f}"
            row_cells[4].text = f"{row['Min_Probability']:.3f}"
            row_cells[5].text = f"{row['Max_Probability']:.3f}"
        
        # Create enhanced visualization
        plt.figure(figsize=(15, 8))
        
        # Subplot 1: Count by gender
        plt.subplot(2, 3, 1)
        sns.barplot(x=gender_stats.index, y=gender_stats['Count'], 
                   hue=gender_stats.index, palette="viridis", legend=False)
        plt.title('High-Risk Count by Gender', fontsize=12, fontweight='bold')
        plt.ylabel('Number of Employees')
        
        # Subplot 2: Average probability by gender
        plt.subplot(2, 3, 2)
        sns.barplot(x=gender_stats.index, y=gender_stats['Avg_Probability'], 
                   hue=gender_stats.index, palette="plasma", legend=False)
        plt.title('Average Risk Probability', fontsize=12, fontweight='bold')
        plt.ylabel('Average Probability')
        
        # Subplot 3: Risk distribution violin plot
        plt.subplot(2, 3, 3)
        if len(high_risk_df['Gender'].unique()) > 1:
            sns.violinplot(data=high_risk_df, x='Gender', y='Attrition_Probability', palette="coolwarm")
        else:
            sns.boxplot(data=high_risk_df, x='Gender', y='Attrition_Probability', palette="coolwarm")
        plt.title('Risk Distribution', fontsize=12, fontweight='bold')
        plt.ylabel('Risk Probability')
        
        # Subplot 4: Histogram of probabilities
        plt.subplot(2, 3, (4, 6))
        for gender in high_risk_df['Gender'].unique():
            gender_data = high_risk_df[high_risk_df['Gender'] == gender]['Attrition_Probability']
            plt.hist(gender_data, alpha=0.7, label=gender, bins=15, density=True)
        plt.title('Risk Probability Distribution by Gender', fontsize=12, fontweight='bold')
        plt.xlabel('Attrition Probability')
        plt.ylabel('Density')
        plt.legend()
        
        plt.tight_layout()
        
        gender_plot_path = f"{output_dir}/Enhanced_Gender_Analysis_{report_time}.png"
        plt.savefig(gender_plot_path, dpi=300, bbox_inches='tight')
        plt.close()

        doc.add_paragraph("\n")
        doc.add_picture(gender_plot_path, width=Inches(7))
        last_paragraph = doc.paragraphs[-1]
        last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        doc.add_page_break()
        
        logger.info("Successfully added enhanced gender analysis section")
        return True
    except Exception as e:
        logger.error(f"Failed to add enhanced gender analysis section: {e}")
        doc.add_paragraph("Error generating Enhanced Gender Analysis section.")
        doc.add_page_break()
        return False

def add_complete_high_risk_list(doc, high_risk_df, logger):
    """Add complete list of high-risk employees"""
    try:
        doc.add_heading('3. Complete List of High-Risk Employees', level=1)
        doc.add_paragraph("Complete listing of all active employees with attrition probability > 70%, ranked by risk level.")
        
        if high_risk_df is None or high_risk_df.empty:
            doc.add_paragraph("No high-risk employees available.")
            return True
        
        # Sort by Attrition_Probability in descending order
        high_risk_df = high_risk_df.sort_values(by='Attrition_Probability', ascending=False)
        
        # Add table with ranking
        table = doc.add_table(rows=1, cols=6)
        table.style = 'Table Grid'
        header_cells = table.rows[0].cells
        header_cells[0].text = 'Rank'
        header_cells[1].text = 'Employee Name'
        header_cells[2].text = 'Function'
        header_cells[3].text = 'Location'
        header_cells[4].text = 'Grade'
        header_cells[5].text = 'Risk Probability'
        style_table_headers(table)
        
        for idx, (_, row) in enumerate(high_risk_df.iterrows(), 1):
            row_cells = table.add_row().cells
            row_cells[0].text = str(idx)
            row_cells[1].text = str(row['Employee Name'])
            row_cells[2].text = str(row['Function'])
            row_cells[3].text = str(row['Job Location'])
            row_cells[4].text = str(row['Grade'])
            row_cells[5].text = f"{row['Attrition_Probability']:.3f}"
        
        doc.add_page_break()
        logger.info("Successfully added complete high-risk list section")
        return True
    except Exception as e:
        logger.error(f"Failed to add complete high-risk list section: {e}")
        doc.add_paragraph("Error generating Complete High-Risk List section.")
        doc.add_page_break()
        return False

def load_data(file_path: str, logger):
    """
    Load and prepare the data for predictive analytics
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

def generate_enhanced_predictive_attrition_report(df, output_dir, logger):
    """Create a comprehensive enhanced predictive analytics report"""
    try:
        # Create document
        doc, report_time = create_predictive_report_document(output_dir)
        
        if doc is None:
            logger.error("Failed to create report document. Aborting report generation.")
            return False, None, None
        
        # Generate predictive analytics section
        high_risk_df, success = add_predictive_analytics_section(doc, df, output_dir, report_time, logger)
        
        if success and high_risk_df is not None and not high_risk_df.empty:
            # Define report sections
            report_sections = [
                (add_enhanced_gender_analysis, "Enhanced Gender Analysis"),
                (add_complete_high_risk_list, "Complete High-Risk List")
            ]
            
            # Execute each section
            successful_sections = 1  # Start with 1 for the main predictive section
            
            for section_func, section_name in report_sections:
                logger.info(f"Generating {section_name} section...")
                
                try:
                    if section_name == "Complete High-Risk List":
                        result = section_func(doc, high_risk_df, logger)
                    else:
                        result = section_func(doc, high_risk_df, output_dir, report_time, logger)
                    
                    if result:
                        successful_sections += 1
                    else:
                        logger.warning(f"Section function returned False: {section_name}")
                except Exception as e:
                    logger.error(f"Error in {section_name} section: {e}")
                    doc.add_paragraph(f"Error generating {section_name} section. Section skipped.")
                    doc.add_page_break()
        
        # Save the document
        try:
            report_path = f"{output_dir}/Enhanced_Predictive_Attrition_Report_{report_time}.docx"
            doc.save(report_path)
            logger.info(f"Enhanced predictive report saved to {report_path}")
            logger.info(f"Report generated with {successful_sections} sections completed successfully")
            
            return True, report_path, report_time
            
        except Exception as e:
            logger.error(f"Failed to save report: {e}")
            return False, None, None
            
    except Exception as e:
        logger.error(f"Failed to generate enhanced predictive report: {e}")
        return False, None, None
