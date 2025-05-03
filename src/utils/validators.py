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

def ensure_required_columns(df):
    """Ensure all required columns exist in the dataframe"""
    ensure_employee_name_column(df)
    ensure_action_type_column(df)
    ensure_gender_column(df)
    ensure_function_column(df)
    ensure_grade_column(df)
