import re
import pandas as pd
import numpy as np

def normalize_text(text):
    """
    Normalize text for matching:
    - Convert to string
    - Strip whitespace
    - Convert to uppercase
    - Handle None/NaN values
    """
    if pd.isna(text) or text is None:
        return ""
    
    try:
        return str(text).strip().upper()
    except:
        return ""

def extract_job_number(job_string):
    """
    Extract numeric part from job number
    Examples:
    "196" -> "196"
    "SGL-25-00196" -> "196"
    "00196" -> "196"
    "JOB-25-00196" -> "196"
    """
    if pd.isna(job_string) or job_string is None:
        return ""
    
    job_str = str(job_string).strip()
    
    # Use regex to find all digits
    digits = re.findall(r'\d+', job_str)
    
    if not digits:
        return ""
    
    # Join all digits
    all_digits = ''.join(digits)
    
    # Remove leading zeros
    result = re.sub(r'^0+', '', all_digits)
    
    # If result is empty (all zeros), return "0"
    return result if result else "0"

def normalize_dataframe(df):
    """
    Create a normalized copy of dataframe with additional columns for matching
    """
    df_norm = df.copy()
    
    # Define columns to normalize
    text_columns = ['Order No', 'Job No', 'Buyer Name', 'Style Name', 'Job Year']
    
    for col in text_columns:
        if col in df_norm.columns:
            norm_col_name = f"{col}_NORM"
            df_norm[norm_col_name] = df_norm[col].apply(normalize_text)
    
    # Ensure Order No column exists
    if 'Order No' in df_norm.columns:
        df_norm['Order No_NORM'] = df_norm['Order No'].apply(normalize_text)
    else:
        df_norm['Order No_NORM'] = ""
    
    # Fill NaN values in numeric columns with 0
    numeric_columns = ['Order Qty.', 'Plan Cut Qty', 'Total Cut Qty', 
                       'Cutting balance', 'Total Sew Input Qty', 
                       'Total Sew Output Qty', 'Total Iron Qty', 
                       'Total Packing Finish Qty', 'Total Ship Out']
    
    for col in numeric_columns:
        if col in df_norm.columns:
            df_norm[col] = pd.to_numeric(df_norm[col], errors='coerce').fillna(0)
    
    return df_norm

def safe_float_conversion(value):
    """
    Safely convert value to float, handling errors
    """
    try:
        if pd.isna(value) or value is None:
            return 0.0
        return float(value)
    except (ValueError, TypeError):
        return 0.0

def detect_file_type(file_path):
    """
    Detect file type based on extension
    Returns: 'excel' or 'csv'
    """
    ext = file_path.lower()
    if ext.endswith('.csv'):
        return 'csv'
    elif ext.endswith(('.xlsx', '.xls')):
        return 'excel'
    else:
        return 'unknown'

def validate_columns(df, required_columns, df_name="DataFrame"):
    """
    Validate that required columns exist in dataframe
    Returns: (bool, list) - (is_valid, missing_columns)
    """
    missing = [col for col in required_columns if col not in df.columns]
    return (len(missing) == 0, missing)

def format_excel_workbook(writer, df, sheet_name):
    """
    Apply formatting to Excel workbook
    """
    workbook = writer.book
    worksheet = writer.sheets[sheet_name]
    
    # Define formats
    header_format = workbook.add_format({
        'bold': True,
        'bg_color': '#D3D3D3',
        'border': 1
    })
    
    cell_format = workbook.add_format({
        'border': 1
    })
    
    # Write headers with formatting
    for col_num, value in enumerate(df.columns.values):
        worksheet.write(0, col_num, value, header_format)
    
    # Write data with formatting
    for row_num in range(len(df)):
        for col_num, value in enumerate(df.iloc[row_num]):
            worksheet.write(row_num + 1, col_num, value, cell_format)
    
    # Auto-adjust column widths
    for col_num, column in enumerate(df.columns):
        max_length = max(
            df[column].astype(str).map(len).max() if len(df) > 0 else 0,
            len(column)
        )
        worksheet.set_column(col_num, col_num, min(max_length + 2, 50))