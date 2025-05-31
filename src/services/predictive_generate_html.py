import pandas as pd
import json
import os
from datetime import datetime
from sklearn.model_selection import train_test_split
from sklearn.ensemble import RandomForestClassifier
from sklearn.preprocessing import LabelEncoder

def calculate_predictive_analytics(df, logger):
    """Calculate predictive analytics and return high-risk employees data"""
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
        
        # Get model accuracy
        train_score = model.score(X_train, y_train)
        test_score = model.score(X_test, y_test)
        
        # Predict for active employees
        active_df = df[~df['Action Type'].str.contains('Exit|Resignation|Termination|Separation', na=False)].copy()
        
        if active_df.empty:
            logger.warning("No active employees found for prediction.")
            # Return a default result structure with zeros/empty lists
            return {
                'model_performance': {
                    'train_accuracy': round(train_score * 100, 2),
                    'test_accuracy': round(test_score * 100, 2)
                },
                'total_active_employees': 0,
                'total_high_risk': 0,
                'high_risk_employees': pd.DataFrame([])
            }
        
        # Prepare active employee features
        X_active = active_df[features].copy()
        for col in X_active.select_dtypes(include=['object']).columns:
            X_active[col] = X_active[col].astype(str).fillna('Unknown')
            if col in label_encoders:
                le = label_encoders[col]
                X_active[col] = X_active[col].apply(lambda x: le.transform([x])[0] if x in le.classes_ else 0)
        
        X_active = X_active.fillna(0)
        
        # Predict probabilities
        active_df['Attrition_Probability'] = model.predict_proba(X_active)[:, 1]
        
        # Filter high-risk employees
        high_risk_employees = active_df[active_df['Attrition_Probability'] > 0.7].copy()
        
        if not high_risk_employees.empty:
            high_risk_employees = high_risk_employees.sort_values(by='Attrition_Probability', ascending=False)
        
        return {
            'model_performance': {
                'train_accuracy': round(train_score * 100, 2),
                'test_accuracy': round(test_score * 100, 2)
            },
            'total_active_employees': len(active_df),
            'total_high_risk': len(high_risk_employees),
            'high_risk_employees': high_risk_employees
        }
        
    except Exception as e:
        logger.error(f"Error in predictive analytics calculation: {e}")
        return {
            'model_performance': {'train_accuracy': 0, 'test_accuracy': 0},
            'total_active_employees': 0,
            'total_high_risk': 0,
            'high_risk_employees': pd.DataFrame([])
        }

def calculate_high_risk_gender_analysis(high_risk_df):
    """Calculate gender-wise analysis of high-risk employees"""
    try:
        if high_risk_df.empty:
            return []
        
        gender_analysis = high_risk_df.groupby('Gender').agg({
            'Employee Name': 'count',
            'Attrition_Probability': 'mean'
        }).rename(columns={'Employee Name': 'High-Risk Count', 'Attrition_Probability': 'Avg Probability'})
        
        result = []
        for gender, row in gender_analysis.iterrows():
            result.append([
                str(gender),
                int(row['High-Risk Count']),
                round(float(row['Avg Probability']), 3)
            ])
        
        return result
    except Exception as e:
        return []

def calculate_high_risk_location_analysis(high_risk_df):
    """Calculate location-wise analysis of high-risk employees"""
    try:
        if high_risk_df.empty:
            return []
        
        # Filter locations with more than 1 high-risk employee
        location_counts = high_risk_df.groupby('Job Location')['Employee Name'].count()
        valid_locations = location_counts[location_counts > 1].index
        
        if len(valid_locations) == 0:
            return []
        
        location_analysis = high_risk_df[high_risk_df['Job Location'].isin(valid_locations)].groupby('Job Location').agg({
            'Employee Name': 'count',
            'Attrition_Probability': 'mean'
        }).rename(columns={'Employee Name': 'High-Risk Count', 'Attrition_Probability': 'Avg Probability'})
        
        result = []
        for location, row in location_analysis.iterrows():
            result.append([
                str(location),
                int(row['High-Risk Count']),
                round(float(row['Avg Probability']), 3)
            ])
        
        return result
    except Exception as e:
        return []

def calculate_high_risk_function_analysis(high_risk_df):
    """Calculate function-wise analysis of high-risk employees"""
    try:
        if high_risk_df.empty:
            return []
        
        # Filter functions with more than 1 high-risk employee
        function_counts = high_risk_df.groupby('Function')['Employee Name'].count()
        valid_functions = function_counts[function_counts > 1].index
        
        if len(valid_functions) == 0:
            return []
        
        function_analysis = high_risk_df[high_risk_df['Function'].isin(valid_functions)].groupby('Function').agg({
            'Employee Name': 'count',
            'Attrition_Probability': 'mean'
        }).rename(columns={'Employee Name': 'High-Risk Count', 'Attrition_Probability': 'Avg Probability'})
        
        result = []
        for function, row in function_analysis.iterrows():
            result.append([
                str(function),
                int(row['High-Risk Count']),
                round(float(row['Avg Probability']), 3)
            ])
        
        return result
    except Exception as e:
        return []

def calculate_high_risk_grade_analysis(high_risk_df):
    """Calculate grade-wise analysis of high-risk employees"""
    try:
        if high_risk_df.empty:
            return []
        
        # Filter grades with more than 1 high-risk employee
        grade_counts = high_risk_df.groupby('Grade')['Employee Name'].count()
        valid_grades = grade_counts[grade_counts > 1].index
        
        if len(valid_grades) == 0:
            return []
        
        grade_analysis = high_risk_df[high_risk_df['Grade'].isin(valid_grades)].groupby('Grade').agg({
            'Employee Name': 'count',
            'Attrition_Probability': 'mean'
        }).rename(columns={'Employee Name': 'High-Risk Count', 'Attrition_Probability': 'Avg Probability'})
        
        result = []
        for grade, row in grade_analysis.iterrows():
            result.append([
                str(grade),
                int(row['High-Risk Count']),
                round(float(row['Avg Probability']), 3)
            ])
        
        return result
    except Exception as e:
        return []

def get_top_risk_employees(high_risk_df, limit=50):
    """Get top high-risk employees for display"""
    try:
        if high_risk_df.empty:
            return []
        
        top_employees = high_risk_df.head(limit)
        result = []
        
        for _, row in top_employees.iterrows():
            result.append([
                str(row['Employee Name']),
                str(row['Function']),
                str(row['Job Location']),
                str(row['Grade']),
                round(float(row['Attrition_Probability']), 3)
            ])
        
        return result
    except Exception as e:
        return []

def generate_predictive_html_report(df, output_dir, file_id, logger):
    """Generate HTML report for predictive analytics"""
    try:
        # Create timestamp for the report
        report_time = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        # Create output directory for this report
        report_dir = f"{output_dir}/{file_id}"
        os.makedirs(report_dir, exist_ok=True)
        
        # Calculate predictive analytics
        predictive_data = calculate_predictive_analytics(df, logger)
        
        # Defensive: If predictive_data is None or missing keys, handle gracefully
        if not predictive_data or 'high_risk_employees' not in predictive_data:
            logger.error("Failed to calculate predictive analytics")
            return False, None, None
        
        high_risk_df = predictive_data['high_risk_employees']
        # If high_risk_df is not a DataFrame, make it an empty DataFrame
        if not isinstance(high_risk_df, pd.DataFrame):
            high_risk_df = pd.DataFrame([])
        
        # Calculate analyses
        gender_data = calculate_high_risk_gender_analysis(high_risk_df)
        location_data = calculate_high_risk_location_analysis(high_risk_df)
        function_data = calculate_high_risk_function_analysis(high_risk_df)
        grade_data = calculate_high_risk_grade_analysis(high_risk_df)
        top_employees = get_top_risk_employees(high_risk_df, 50)
        all_employees = get_top_risk_employees(high_risk_df, len(high_risk_df))
        
        # Prepare data for JavaScript
        report_data = {
            'model_performance': predictive_data.get('model_performance', {'train_accuracy': 0, 'test_accuracy': 0}),
            'summary': {
                'total_active_employees': predictive_data.get('total_active_employees', 0),
                'total_high_risk': predictive_data.get('total_high_risk', 0),
                'risk_percentage': round((predictive_data.get('total_high_risk', 0) / predictive_data.get('total_active_employees', 1) * 100), 2) if predictive_data.get('total_active_employees', 0) > 0 else 0
            },
            'top_employees': top_employees,
            'all_employees': all_employees,
            'gender': gender_data,
            'location': location_data,
            'function': function_data,
            'grade': grade_data
        }
        
        # HTML template for predictive analytics
        html_template = '''<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Predictive Attrition Analytics Report</title>
    <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
    <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    <style>
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background-color: #f8f9fa;
        }
        .dashboard-header {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 30px;
            margin-bottom: 30px;
        }
        .section-title {
            color: #667eea;
            font-weight: bold;
            margin-bottom: 20px;
            border-bottom: 2px solid #667eea;
            padding-bottom: 10px;
        }
        .card {
            margin-bottom: 20px;
            box-shadow: 0 4px 8px rgba(0,0,0,0.1);
            border: none;
            border-radius: 10px;
        }
        .stat-card {
            background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%);
            color: white;
            border-radius: 10px;
        }
        .risk-card {
            background: linear-gradient(135deg, #4facfe 0%, #00f2fe 100%);
            color: white;
            border-radius: 10px;
        }
        .performance-card {
            background: linear-gradient(135deg, #43e97b 0%, #38f9d7 100%);
            color: white;
            border-radius: 10px;
        }
        .chart-container {
            position: relative;
            height: 400px;
            background: white;
            border-radius: 10px;
            padding: 20px;
        }
        .table th {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            border: none;
        }
        .table-striped tbody tr:nth-of-type(odd) {
            background-color: #f8f9ff;
        }
        .risk-high {
            background-color: #ffebee !important;
            border-left: 4px solid #f44336;
        }
        .risk-medium {
            background-color: #fff3e0 !important;
            border-left: 4px solid #ff9800;
        }
        .nav-pills .nav-link.active {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        }
        .employee-list {
            max-height: 500px;
            overflow-y: auto;
        }
    </style>
</head>
<body>
    <div class="dashboard-header text-center">
        <h1><i class="fas fa-brain"></i> PREDICTIVE ATTRITION ANALYTICS REPORT</h1>
        <h3>BY VIMAL SINGH</h3>
        <p id="reportDate">Report Generated on: </p>
    </div>

    <div class="container-fluid">
        <!-- Summary Cards -->
        <div class="row mb-4">
            <div class="col-md-3">
                <div class="card stat-card p-3 text-center">
                    <h4>Total Active Employees</h4>
                    <h2 id="totalActive">-</h2>
                </div>
            </div>
            <div class="col-md-3">
                <div class="card risk-card p-3 text-center">
                    <h4>High-Risk Employees</h4>
                    <h2 id="totalHighRisk">-</h2>
                </div>
            </div>
            <div class="col-md-3">
                <div class="card performance-card p-3 text-center">
                    <h4>Risk Percentage</h4>
                    <h2 id="riskPercentage">-</h2>
                </div>
            </div>
            <div class="col-md-3">
                <div class="card performance-card p-3 text-center">
                    <h4>Model Accuracy</h4>
                    <h2 id="modelAccuracy">-</h2>
                </div>
            </div>
        </div>

        <!-- Top Risk Employees -->
        <div class="row">
            <div class="col-12">
                <h2 class="section-title">1. High-Risk Employees (Probability > 70%)</h2>
                
                <!-- Navigation Tabs -->
                <ul class="nav nav-pills mb-3" id="employeeTabs" role="tablist">
                    <li class="nav-item" role="presentation">
                        <button class="nav-link active" id="top50-tab" data-bs-toggle="pill" data-bs-target="#top50" type="button">
                            Top 50 Employees
                        </button>
                    </li>
                    <li class="nav-item" role="presentation">
                        <button class="nav-link" id="all-tab" data-bs-toggle="pill" data-bs-target="#all" type="button">
                            All High-Risk Employees
                        </button>
                    </li>
                </ul>

                <!-- Tab Content -->
                <div class="tab-content" id="employeeTabContent">
                    <div class="tab-pane fade show active" id="top50" role="tabpanel">
                        <div class="card">
                            <div class="card-body employee-list">
                                <table class="table table-striped" id="top50Table">
                                    <thead>
                                        <tr>
                                            <th>Employee Name</th>
                                            <th>Function</th>
                                            <th>Location</th>
                                            <th>Grade</th>
                                            <th>Risk Probability</th>
                                        </tr>
                                    </thead>
                                    <tbody></tbody>
                                </table>
                            </div>
                        </div>
                    </div>
                    <div class="tab-pane fade" id="all" role="tabpanel">
                        <div class="card">
                            <div class="card-body employee-list">
                                <table class="table table-striped" id="allTable">
                                    <thead>
                                        <tr>
                                            <th>Employee Name</th>
                                            <th>Function</th>
                                            <th>Location</th>
                                            <th>Grade</th>
                                            <th>Risk Probability</th>
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

        <!-- Analysis Sections -->
        <div class="row mt-5">
            <div class="col-md-6">
                <h2 class="section-title">2. Risk Analysis by Gender</h2>
                <div class="row">
                    <div class="col-md-6">
                        <table class="table table-striped" id="genderTable">
                            <thead>
                                <tr>
                                    <th>Gender</th>
                                    <th>High-Risk Count</th>
                                    <th>Avg Probability</th>
                                </tr>
                            </thead>
                            <tbody></tbody>
                        </table>
                    </div>
                    <div class="col-md-6">
                        <div class="chart-container">
                            <canvas id="genderChart"></canvas>
                        </div>
                    </div>
                </div>
            </div>
            <div class="col-md-6">
                <h2 class="section-title">3. Risk Analysis by Location</h2>
                <div class="chart-container">
                    <canvas id="locationChart"></canvas>
                </div>
            </div>
        </div>

        <div class="row mt-5">
            <div class="col-md-6">
                <h2 class="section-title">4. Risk Analysis by Function</h2>
                <div class="chart-container">
                    <canvas id="functionChart"></canvas>
                </div>
            </div>
            <div class="col-md-6">
                <h2 class="section-title">5. Risk Analysis by Grade</h2>
                <div class="chart-container">
                    <canvas id="gradeChart"></canvas>
                </div>
            </div>
        </div>
    </div>

    <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/js/bootstrap.bundle.min.js"></script>
    <script>
        // Color scheme for charts
        const colors = {
            primary: '#667eea',
            secondary: '#764ba2',
            accent: '#f093fb',
            warning: '#ff9800',
            success: '#43e97b',
            info: '#4facfe'
        };

        function initializeReport(data) {
            // Set report date
            document.getElementById('reportDate').textContent += new Date().toLocaleString();

            // Summary statistics
            document.getElementById('totalActive').textContent = data.summary.total_active_employees;
            document.getElementById('totalHighRisk').textContent = data.summary.total_high_risk;
            document.getElementById('riskPercentage').textContent = data.summary.risk_percentage + '%';
            document.getElementById('modelAccuracy').textContent = data.model_performance.test_accuracy + '%';

            // Populate employee tables
            populateEmployeeTable('top50Table', data.top_employees);
            populateEmployeeTable('allTable', data.all_employees);

            // Populate analysis tables and charts
            populateAnalysisTable('genderTable', data.gender);
            createAnalysisCharts(data);
        }

        function populateEmployeeTable(tableId, employees) {
            const tbody = document.querySelector(`#${tableId} tbody`);
            tbody.innerHTML = '';
            
            employees.forEach(employee => {
                const tr = document.createElement('tr');
                const probability = parseFloat(employee[4]);
                
                // Add risk level styling
                if (probability >= 0.9) {
                    tr.className = 'risk-high';
                } else if (probability >= 0.8) {
                    tr.className = 'risk-medium';
                }
                
                employee.forEach(value => {
                    const td = document.createElement('td');
                    if (typeof value === 'number' && value < 1) {
                        td.textContent = (value * 100).toFixed(1) + '%';
                    } else {
                        td.textContent = value;
                    }
                    tr.appendChild(td);
                });
                tbody.appendChild(tr);
            });
        }

        function populateAnalysisTable(tableId, data) {
            const tbody = document.querySelector(`#${tableId} tbody`);
            tbody.innerHTML = '';
            
            data.forEach(row => {
                const tr = document.createElement('tr');
                row.forEach((value, index) => {
                    const td = document.createElement('td');
                    if (index === 2 && typeof value === 'number') {
                        td.textContent = (value * 100).toFixed(1) + '%';
                    } else {
                        td.textContent = value;
                    }
                    tr.appendChild(td);
                });
                tbody.appendChild(tr);
            });
        }

        function createAnalysisCharts(data) {
            // Gender chart
            if (data.gender.length > 0) {
                createBarChart('genderChart', data.gender, 'High-Risk Count by Gender');
            }

            // Location chart
            if (data.location.length > 0) {
                createBarChart('locationChart', data.location, 'High-Risk Count by Location');
            }

            // Function chart
            if (data.function.length > 0) {
                createBarChart('functionChart', data.function, 'High-Risk Count by Function');
            }

            // Grade chart
            if (data.grade.length > 0) {
                createBarChart('gradeChart', data.grade, 'High-Risk Count by Grade');
            }
        }

        function createBarChart(chartId, data, title) {
            const ctx = document.getElementById(chartId).getContext('2d');
            new Chart(ctx, {
                type: 'bar',
                data: {
                    labels: data.map(row => row[0]),
                    datasets: [{
                        label: 'High-Risk Count',
                        data: data.map(row => row[1]),
                        backgroundColor: data.map((_, index) => 
                            index % 3 === 0 ? colors.primary : 
                            index % 3 === 1 ? colors.accent : colors.info
                        ),
                        borderRadius: 5
                    }]
                },
                options: {
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: {
                        title: {
                            display: true,
                            text: title,
                            font: { size: 16, weight: 'bold' }
                        },
                        legend: {
                            display: false
                        }
                    },
                    scales: {
                        y: {
                            beginAtZero: true,
                            ticks: {
                                stepSize: 1
                            }
                        }
                    }
                }
            });
        }

        // Initialize report with data
        const reportData = REPLACE_WITH_DATA;
        initializeReport(reportData);
    </script>
</body>
</html>'''
        
        # Replace data placeholder with actual data
        final_html = html_template.replace('REPLACE_WITH_DATA', json.dumps(report_data))
        
        # Save HTML file
        report_filename = f"Predictive_Attrition_Report_{report_time}.html"
        html_path = f"{report_dir}/{report_filename}"
        with open(html_path, 'w', encoding='utf-8') as f:
            f.write(final_html)
        
        logger.info(f"Predictive HTML report saved to {html_path}")
        return True, html_path, report_filename
        
    except Exception as e:
        logger.error(f"Failed to generate predictive HTML report: {e}")
        return False, None, None