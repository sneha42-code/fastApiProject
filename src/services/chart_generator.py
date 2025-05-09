import matplotlib.pyplot as plt
import seaborn as sns
import matplotlib.gridspec as gridspec
from datetime import datetime
import pandas as pd

def create_overall_pie_chart(total_employees, total_exits, output_dir, report_time):
    """Create overall attrition pie chart"""
    plt.figure(figsize=(10, 6))
    labels = ['Active Employees', 'Exited Employees']
    values = [total_employees - total_exits, total_exits]
    plt.pie(values, labels=labels, autopct='%1.1f%%', startangle=90, explode=[0, 0.1])
    plt.title('Overall Attrition', fontsize=16, fontweight='bold')
    plt.tight_layout()
    
    chart_path = f"{output_dir}/Overall_Statistics_{report_time}.png"
    plt.savefig(chart_path)
    plt.close()
    
    return chart_path

def create_gender_charts(attrition_by_gender, output_dir, report_time):
    """Create gender analysis charts"""
    fig = plt.figure(figsize=(16, 7))
    gs = gridspec.GridSpec(1, 2, width_ratios=[2, 3])
    ax1 = plt.subplot(gs[0])
    ax2 = plt.subplot(gs[1])

    sns.barplot(x=attrition_by_gender.index, y=attrition_by_gender['Attrition Count'], 
                ax=ax1, hue=attrition_by_gender.index, palette="coolwarm", width=0.4)
    ax1.set_title('Attrition Count by Gender\n', fontsize=15, fontweight='bold', color='black')
    ax1.set_ylabel('Number of Employees', fontsize=14)
    ax1.set_xlabel('Gender', fontsize=14)
    
    ax2.pie(attrition_by_gender['Attrition Rate %'], labels=attrition_by_gender.index, 
            autopct='%1.1f%%', textprops={'fontsize': 14})
    ax2.set_title('Attrition Rate Distribution by Gender\n', fontsize=15, fontweight='bold', color='black')
    plt.tight_layout()
    
    chart_path = f"{output_dir}/Gender_Analysis_{report_time}.png"
    plt.savefig(chart_path)
    plt.close()
    
    return chart_path

def create_location_charts(attrition_by_location, output_dir, report_time):
    """Create location analysis charts"""
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(16, 7))
    sns.barplot(data=attrition_by_location.reset_index(), x='Job Location', y='Attrition Count', ax=ax1)
    ax1.set_title('Attrition Count by Location\n', fontsize=16, fontweight='bold')
    ax1.set_xlabel('Location', fontsize=14)
    ax1.set_ylabel('Number of Employees', fontsize=14)
    ax1.tick_params(axis='x', rotation=45)
    
    ax2.pie(attrition_by_location['Attrition Rate %'], labels=attrition_by_location.index, 
            autopct='%1.1f%%', textprops={'fontsize': 14})
    ax2.set_title('Attrition Rate Distribution by Location\n', fontsize=16, fontweight='bold')
    plt.tight_layout()
    
    chart_path = f"{output_dir}/Location_Analysis_{report_time}.png"
    plt.savefig(chart_path, bbox_inches='tight')
    plt.close()
    
    return chart_path

def create_function_charts(attrition_by_function, output_dir, report_time):
    """Create function analysis charts"""
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(17, 7))
    sns.barplot(data=attrition_by_function.reset_index(), x='Function', y='Attrition Count', 
                hue='Function', ax=ax1, legend=False)
    ax1.set_title('Attrition Count by Function\n', fontsize=16, fontweight='bold')
    ax1.set_ylabel('Number of Employees', fontsize=14)
    ax1.set_xlabel('Function', fontsize=14)
    ax1.tick_params(axis='x', rotation=45)
    
    ax2.pie(attrition_by_function['Attrition Rate %'], labels=attrition_by_function.index, 
            autopct='%1.1f%%', textprops={'fontsize': 14})
    ax2.set_title('Attrition Rate Distribution by Function\n', fontsize=16, fontweight='bold')
    plt.tight_layout()
    
    chart_path = f"{output_dir}/Function_Analysis_{report_time}.png"
    plt.savefig(chart_path, bbox_inches='tight')
    plt.close()
    
    return chart_path

def create_tenure_charts(tenure_analysis, output_dir, report_time):
    """Create tenure analysis charts"""
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(16, 7))

    # Bar chart - Fix for seaborn FutureWarning
    sns.barplot(data=tenure_analysis, x='Tenure Band', y='Count', hue='Tenure Band', ax=ax1, legend=False)
    ax1.set_title('Attrition Count by Tenure\n', fontsize=16, fontweight='bold')
    ax1.set_ylabel('Number of Employees', fontsize=14)
    ax1.set_xlabel('Tenure Band', fontsize=14)
    ax1.tick_params(axis='x', rotation=45)

    # Pie chart
    ax2.pie(tenure_analysis['Percentage'], labels=tenure_analysis['Tenure Band'], 
            autopct='%1.1f%%', textprops={'fontsize': 14})
    ax2.set_title('Attrition Distribution by Tenure\n', fontsize=16, fontweight='bold')
    plt.tight_layout()

    chart_path = f"{output_dir}/Tenure_Analysis_{report_time}.png"
    plt.savefig(chart_path, bbox_inches='tight')
    plt.close()
    
    return chart_path

def create_grade_charts(attrition_by_grade, output_dir, report_time):
    """Create grade analysis charts"""
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(16, 7))

    # Bar chart
    sns.barplot(data=attrition_by_grade.reset_index(), x='Grade', y='Attrition Count', 
                hue='Grade', ax=ax1, legend=False)
    ax1.set_title('Attrition Count by Grade\n', fontsize=16, fontweight='bold')
    ax1.set_ylabel('Number of Employees', fontsize=14)
    ax1.set_xlabel('Grade', fontsize=14)
    ax1.tick_params(axis='x', rotation=45)

    # Pie chart
    ax2.pie(attrition_by_grade['Attrition Rate %'], labels=attrition_by_grade.index, 
            autopct='%1.1f%%', textprops={'fontsize': 14})
    ax2.set_title('Attrition Rate Distribution by Grade\n', fontsize=16, fontweight='bold')
    plt.tight_layout()

    chart_path = f"{output_dir}/Grade_Analysis_{report_time}.png"
    plt.savefig(chart_path, bbox_inches='tight')
    plt.close()
    
    return chart_path

def create_quarterly_trend_chart(quarterly_trends, output_dir, report_time):
    """Create quarterly trend chart"""
    plt.figure(figsize=(10, 6))
    ax = sns.barplot(data=quarterly_trends, x=quarterly_trends['Quarter'].astype(str), 
                     y='Exit Count', hue=quarterly_trends['Quarter'].astype(str), legend=False)
    
    plt.title('Quarterly Attrition Trend\n', fontsize=16, fontweight='bold')
    plt.ylabel('Number of Exits', fontsize=14)
    plt.xlabel('Quarter', fontsize=14)
    plt.grid(True, linestyle='--', alpha=0.7, axis='y')
    plt.tight_layout()

    # Add count labels on top of each bar
    for i, v in enumerate(quarterly_trends['Exit Count']):
        ax.text(i, v + 0.5, str(v), ha='center', fontweight='bold')

    chart_path = f"{output_dir}/Quarterly_Trend_{report_time}.png"
    plt.savefig(chart_path, bbox_inches='tight')
    plt.close()
    
    return chart_path

def create_monthly_trend_chart(monthly_trends, output_dir, report_time):
    """Create monthly trend chart"""
    plt.figure(figsize=(12, 6))
    sns.lineplot(data=monthly_trends, x='Month', y='Exit Count', marker='o', linewidth=2)
    plt.title('Monthly Attrition Trend\n', fontsize=16, fontweight='bold')
    plt.ylabel('Number of Exits', fontsize=14)
    plt.xlabel('Month', fontsize=14)
    plt.xticks(rotation=45)
    plt.grid(True, linestyle='--', alpha=0.7)
    plt.tight_layout()

    chart_path = f"{output_dir}/Monthly_Trend_{report_time}.png"
    plt.savefig(chart_path, bbox_inches='tight')
    plt.close()
    
    return chart_path