import pandas as pd
import numpy as np
import os
import re
from utils import normalize_text, extract_job_number, normalize_dataframe

def extract_numeric_from_job(job_string):
    """
    Extract numeric part from job number in File 2 format
    Examples:
    "SGL-25-00196" -> "196"
    "JOB-25-01062" -> "1062"
    "RAL-25-00678" -> "678"
    "196" -> "196"
    """
    if pd.isna(job_string) or job_string is None:
        return ""
    
    job_str = str(job_string).strip()
    
    # Find the last numeric part (after the last hyphen)
    parts = job_str.split('-')
    if len(parts) > 1:
        last_part = parts[-1]
        # Remove leading zeros
        return re.sub(r'^0+', '', last_part)
    else:
        # If no hyphens, just remove leading zeros
        return re.sub(r'^0+', '', job_str)

def find_job_pos(df2, job_input):
    """
    Find all POs for a given job number
    
    Args:
        df2: Data Sheet 2 dataframe (production data)
        job_input: Job number to search for (e.g., "196" or "SGL-25-00196")
    
    Returns:
        DataFrame with matching POs
    """
    
    print(f"\n=== Job Lookup ===")
    print(f"Searching for job: {job_input}")
    
    # Extract numeric part from input job
    search_job = extract_numeric_from_job(job_input)
    print(f"Extracted search job number: '{search_job}'")
    
    if not search_job:
        print("No job number extracted")
        return pd.DataFrame()
    
    # Create a copy
    df_copy = df2.copy()
    
    # Print column names for debugging
    print(f"Available columns in production data: {list(df_copy.columns)}")
    
    # Find Job No column in production data - based on your actual data
    job_col = None
    # Your production data has 'Job No' column (from the analysis)
    possible_job_cols = ['Job No', 'JOB NO', 'job no', 'Job_No']
    
    for col in possible_job_cols:
        if col in df_copy.columns:
            job_col = col
            print(f"Found Job No column: '{job_col}'")
            break
    
    if job_col is None:
        print("No Job No column found in production data")
        return pd.DataFrame()
    
    # Production data already has numeric job numbers (196, 202, etc.)
    # Convert to string for comparison
    df_copy['JOB_STR'] = df_copy[job_col].astype(str).str.strip()
    
    # Print sample job numbers
    sample_jobs = df_copy['JOB_STR'].head(10).tolist()
    print(f"Sample job numbers in production data: {sample_jobs}")
    
    # Direct match (production data has numeric like "196")
    mask = df_copy['JOB_STR'] == search_job
    match_count = mask.sum()
    print(f"Found {match_count} matching rows for job '{search_job}'")
    
    if match_count == 0:
        # Try without removing leading zeros
        mask = df_copy['JOB_STR'].str.zfill(3) == search_job.zfill(3)
        match_count = mask.sum()
        print(f"Found {match_count} matching rows with zero-padded match")
    
    if not mask.any():
        print(f"No matches found for job: {job_input}")
        return pd.DataFrame()
    
    # Select relevant columns for display - BASED ON YOUR ACTUAL DATA
    # From your proddata.xls, the columns are:
    # 'Order No', 'Style Name', 'Item Name', 'Order Qty.', 'Ship Date'
    
    column_mapping = {
        'Order No': ['Order No', 'ORDER NO', 'Order_No'],
        'Style Name': ['Style Name', 'STYLE NAME', 'Style'],
        'Item Name': ['Item Name', 'ITEM NAME', 'Item'],
        'Order Qty.': ['Order Qty.', 'ORDER QTY', 'Qty'],
        'Ship Date': ['Ship Date', 'SHIP DATE', 'Ex-Factory Date']
    }
    
    # Find available columns
    available_cols = []
    col_rename = {}
    
    for display_col, possible_names in column_mapping.items():
        found = False
        for name in possible_names:
            if name in df_copy.columns:
                available_cols.append(name)
                col_rename[name] = display_col
                print(f"Found column '{name}' for '{display_col}'")
                found = True
                break
        if not found:
            print(f"Warning: No column found for '{display_col}'")
    
    # If we have the basic columns, create a result
    if available_cols:
        result_df = df_copy.loc[mask, available_cols].copy()
        result_df = result_df.rename(columns=col_rename)
    else:
        # Fallback: return whatever we have
        print("No display columns found, returning basic info")
        # Try to find at least Order No
        order_cols = ['Order No', 'ORDER NO', 'Order_No']
        order_col = None
        for col in order_cols:
            if col in df_copy.columns:
                order_col = col
                break
        
        if order_col:
            result_df = df_copy.loc[mask, [job_col, order_col]].copy()
            result_df = result_df.rename(columns={job_col: 'Job No', order_col: 'Order No'})
        else:
            result_df = df_copy.loc[mask, [job_col]].copy()
            result_df = result_df.rename(columns={job_col: 'Job No'})
    
    print(f"Returning {len(result_df)} results")
    return result_df

def process_files(file1_path, file2_path, sheet_name, output_path, status_callback=None):
    """
    Main processing function to match and merge the two Excel files
    
    Args:
        file1_path: Path to Data Sheet 1 (Schedule file with buyer orders)
        file2_path: Path to Data Sheet 2 (Production data file)
        sheet_name: Sheet name to read from Data Sheet 1 (e.g., 'Target', 'Kmart')
        output_path: Path to save the output file
        status_callback: Optional callback function for status updates
    
    Returns:
        Boolean indicating success/failure
    """
    
    def log(message):
        if status_callback:
            status_callback(message)
        print(message)
    
    try:
        # Step 1: Load the files
        log("\n=== Loading Files ===")
        log(f"File 1 (Schedule): {os.path.basename(file1_path)}")
        log(f"File 2 (Production): {os.path.basename(file2_path)}")
        log(f"Selected sheet: {sheet_name}")
        
        # Load File 1 (Schedule - Buyer Orders)
        file_ext1 = os.path.splitext(file1_path)[1].lower()
        
        if file_ext1 == '.csv':
            df1 = pd.read_csv(file1_path, dtype=str)
        else:
            df1 = pd.read_excel(file1_path, sheet_name=sheet_name, dtype=str)
        
        log(f"Loaded {len(df1)} rows from Schedule file")
        
        # Load File 2 (Production Data)
        file_ext2 = os.path.splitext(file2_path)[1].lower()
        
        if file_ext2 == '.csv':
            df2 = pd.read_csv(file2_path, dtype=str)
        else:
            # Production data is in the first/only sheet
            df2 = pd.read_excel(file2_path, dtype=str)
        
        log(f"Loaded {len(df2)} rows from Production data")
        
        # Step 2: Map columns in File 1 (Schedule)
        log("\n=== Mapping Schedule File Columns ===")
        
        # Schedule file columns: SL, JOB NO, Order No, Style, Color
        col_mapping_df1 = {
            'JOB NO': ['JOB NO', 'Job No', 'JOB_NUMBER'],
            'Order No': ['Order No', 'ORDER NO', 'PO No'],
            'STYLE NO': ['Style', 'STYLE', 'Style No', 'STYLE NO'],
            'COLOR': ['Color', 'COLOR', 'Colour']
        }
        
        df1_renamed = {}
        for std_col, possible_names in col_mapping_df1.items():
            for name in possible_names:
                if name in df1.columns:
                    df1_renamed[name] = std_col
                    log(f"  Mapped '{name}' to '{std_col}'")
                    break
        
        if df1_renamed:
            df1 = df1.rename(columns=df1_renamed)
        
        # Step 3: Map columns in File 2 (Production Data)
        log("\n=== Mapping Production File Columns ===")
        
        # Production data columns based on the actual file
        col_mapping_df2 = {
            'Job No': ['Job No', 'JOB NO'],
            'Order No': ['Order No', 'ORDER NO'],
            'Order Qty.': ['Order Qty.', 'ORDER QTY'],
            'Plan Cut Qty': ['Plan Cut Qty', 'PLAN CUT QTY'],
            'Total Cut Qty': ['Total Cut Qty', 'TOTAL CUT QTY'],
            'Cutting balance': ['Cutting balance', 'CUTTING BALANCE'],
            'Total Sew Input Qty': ['Total Sew Input Qty', 'TOTAL SEW INPUT'],
            'Total Sew Output Qty': ['Total Sew Output Qty', 'TOTAL SEW OUTPUT'],
            'Total Iron Qty': ['Total Iron Qty', 'TOTAL IRON QTY'],
            'Total Packing Finish Qty': ['Total Packing Finish Qty', 'TOTAL PACKING FINISH'],
            'Total Ship Out': ['Total Ship Out', 'TOTAL SHIP OUT'],
            'Style Name': ['Style Name', 'STYLE NAME'],
            'Item Name': ['Item Name', 'ITEM NAME'],
            'Ship Date': ['Ship Date', 'SHIP DATE', 'Ex-Factory Date']
        }
        
        df2_renamed = {}
        for std_col, possible_names in col_mapping_df2.items():
            for name in possible_names:
                if name in df2.columns:
                    df2_renamed[name] = std_col
                    log(f"  Mapped '{name}' to '{std_col}'")
                    break
        
        if df2_renamed:
            df2 = df2.rename(columns=df2_renamed)
        
        # Step 4: Convert numeric columns in production data
        log("\n=== Converting Data Types ===")
        
        numeric_cols_df2 = ['Order Qty.', 'Plan Cut Qty', 'Total Cut Qty', 
                           'Cutting balance', 'Total Sew Input Qty',
                           'Total Sew Output Qty', 'Total Iron Qty', 
                           'Total Packing Finish Qty', 'Total Ship Out']
        
        for col in numeric_cols_df2:
            if col in df2.columns:
                df2[col] = pd.to_numeric(df2[col], errors='coerce').fillna(0)
        
        # Step 5: Prepare File 1 for matching
        log("\n=== Preparing Schedule Data for Matching ===")
        
        # Extract numeric job number from File 1 (Schedule)
        if 'JOB NO' in df1.columns:
            df1['EXTRACTED_JOB'] = df1['JOB NO'].apply(extract_numeric_from_job)
            log(f"Extracted job numbers from Schedule data")
            # Show sample
            sample = df1[['JOB NO', 'EXTRACTED_JOB']].head(3)
            for _, row in sample.iterrows():
                log(f"  {row['JOB NO']} -> {row['EXTRACTED_JOB']}")
        else:
            log("Warning: 'JOB NO' column not found in Schedule file")
            df1['EXTRACTED_JOB'] = ""
        
        # Normalize Order No
        if 'Order No' in df1.columns:
            df1['Order No_NORM'] = df1['Order No'].apply(normalize_text)
        else:
            df1['Order No_NORM'] = ""
        
        # Step 6: Prepare File 2 for matching
        log("\n=== Preparing Production Data for Matching ===")
        
        # Production data already has numeric job numbers
        if 'Job No' in df2.columns:
            df2['JOB_STR'] = df2['Job No'].astype(str).str.strip()
            unique_jobs = df2['JOB_STR'].unique()[:10]
            log(f"Job numbers in production data: {unique_jobs}")
        else:
            df2['JOB_STR'] = ""
            log("Warning: 'Job No' column not found in Production data")
        
        # Normalize Order No in production data
        if 'Order No' in df2.columns:
            df2['Order No_NORM'] = df2['Order No'].apply(normalize_text)
        else:
            df2['Order No_NORM'] = ""
        
        # Step 7: Create lookup dictionary for fast matching
        log("\n=== Creating Lookup Dictionary ===")
        
        job_lookup = {}
        lookup_count = 0
        
        for idx, row in df2.iterrows():
            job_num = row.get('JOB_STR', '')
            order_norm = row.get('Order No_NORM', '')
            
            if job_num and job_num != "":
                if job_num not in job_lookup:
                    job_lookup[job_num] = {}
                
                job_lookup[job_num][order_norm] = row.to_dict()
                lookup_count += 1
        
        log(f"Created lookup with {lookup_count} entries for {len(job_lookup)} unique jobs")
        
        # Step 8: Process each row in File 1
        log("\n=== Matching Rows ===")
        
        # Ensure SL column exists
        if 'SL' not in df1.columns:
            df1['SL'] = range(1, len(df1) + 1)
        
        # Define output columns
        output_cols = [
            'SL', 'JOB NO', 'Order No', 'STYLE NO', 'COLOR',
            'Order Qty.', 'Plan Cut Qty', 'Total Cut Qty', 'Cutting balance',
            'Total Sew Input Qty', 'Total Sew Output Qty', 'Sewing Balance',
            'Total Iron Qty', 'Total Packing Finish Qty', 'Total Ship Out'
        ]
        
        output_df = pd.DataFrame(columns=output_cols)
        
        matched_count = 0
        unmatched_count = 0
        
        for idx, row in df1.iterrows():
            if idx % 50 == 0 and idx > 0:
                log(f"Processed {idx} rows...")
            
            extracted_job = row.get('EXTRACTED_JOB', '')
            order_norm = row.get('Order No_NORM', '')
            
            # Try to find match
            if extracted_job and extracted_job in job_lookup:
                # Check if this specific Order No exists for this job
                if order_norm in job_lookup[extracted_job]:
                    match_dict = job_lookup[extracted_job][order_norm]
                    matched_count += 1
                    
                    # Calculate sewing balance
                    sew_input = match_dict.get('Total Sew Input Qty', 0) or 0
                    sew_output = match_dict.get('Total Sew Output Qty', 0) or 0
                    sewing_balance = sew_input - sew_output
                    
                    output_row = {
                        'SL': row.get('SL', ''),
                        'JOB NO': row.get('JOB NO', ''),
                        'Order No': row.get('Order No', ''),
                        'STYLE NO': row.get('STYLE NO', ''),
                        'COLOR': row.get('COLOR', ''),
                        'Order Qty.': match_dict.get('Order Qty.', ''),
                        'Plan Cut Qty': match_dict.get('Plan Cut Qty', ''),
                        'Total Cut Qty': match_dict.get('Total Cut Qty', ''),
                        'Cutting balance': match_dict.get('Cutting balance', ''),
                        'Total Sew Input Qty': sew_input,
                        'Total Sew Output Qty': sew_output,
                        'Sewing Balance': sewing_balance,
                        'Total Iron Qty': match_dict.get('Total Iron Qty', ''),
                        'Total Packing Finish Qty': match_dict.get('Total Packing Finish Qty', ''),
                        'Total Ship Out': match_dict.get('Total Ship Out', '')
                    }
                else:
                    # Job matched but Order No didn't match
                    unmatched_count += 1
                    output_row = create_empty_row(row)
                    if idx < 10:  # Log first few unmatched for debugging
                        log(f"  Job '{extracted_job}' matched but Order No '{order_norm}' not found")
            else:
                # No job match
                unmatched_count += 1
                output_row = create_empty_row(row)
                if extracted_job and idx < 10:  # Log first few unmatched for debugging
                    log(f"  Job '{extracted_job}' not found in production data")
            
            output_df = pd.concat([output_df, pd.DataFrame([output_row])], ignore_index=True)
        
        log(f"\n=== Match Results ===")
        log(f"Matched: {matched_count}")
        log(f"Unmatched: {unmatched_count}")
        log(f"Total: {matched_count + unmatched_count}")
        
        # Step 9: Save output file
        log(f"\n=== Saving Output ===")
        log(f"Saving to: {output_path}")
        
        # Reorder columns
        existing_cols = [col for col in output_cols if col in output_df.columns]
        output_df = output_df[existing_cols]
        
        # Save to Excel
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            output_df.to_excel(writer, sheet_name='Matched Results', index=False)
            
            # Auto-adjust column widths
            worksheet = writer.sheets['Matched Results']
            for i, column in enumerate(output_df.columns):
                if len(output_df) > 0:
                    column_width = max(
                        output_df[column].astype(str).map(len).max(),
                        len(column)
                    )
                else:
                    column_width = len(column)
                worksheet.column_dimensions[chr(65 + i)].width = min(column_width + 2, 50)
        
        log("✅ File saved successfully!")
        return True
        
    except Exception as e:
        log(f"❌ Error in processing: {str(e)}")
        import traceback
        traceback.print_exc()
        return False

def create_empty_row(row):
    """Create an empty row template for unmatched entries"""
    return {
        'SL': row.get('SL', ''),
        'JOB NO': row.get('JOB NO', ''),
        'Order No': row.get('Order No', ''),
        'STYLE NO': row.get('STYLE NO', ''),
        'COLOR': row.get('COLOR', ''),
        'Order Qty.': '',
        'Plan Cut Qty': '',
        'Total Cut Qty': '',
        'Cutting balance': '',
        'Total Sew Input Qty': '',
        'Total Sew Output Qty': '',
        'Sewing Balance': '',
        'Total Iron Qty': '',
        'Total Packing Finish Qty': '',
        'Total Ship Out': ''
    }