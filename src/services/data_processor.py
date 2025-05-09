import pandas as pd
from src.utils.validators import ensure_required_columns

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