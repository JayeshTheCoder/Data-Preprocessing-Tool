import pandas as pd
import os
import calendar
from datetime import datetime
import re # Added for parsing filename

def get_comp_numbers(directory_file_path):
    """
    Reads the Directory_Processed_Output.xlsx file and returns a list of
    unique Comp_No values where Type is 'PO' or 'MOPO'.
    """
    print(f"Reading directory file from: {directory_file_path}")
    try:
        df_dir = pd.read_excel(directory_file_path)
    except FileNotFoundError:
        print(f"❌ ERROR: Directory file not found at: {directory_file_path}")
        return None
    except Exception as e:
        print(f"❌ ERROR: Could not read directory file. Error: {e}")
        return None

    if "Comp_No" not in df_dir.columns or "Type" not in df_dir.columns:
        print("❌ ERROR: Directory file must contain 'Comp_No' and 'Type' columns.")
        return None

    try:
        filtered_dir = df_dir[df_dir['Type'].astype(str).str.contains("PO|MOPO", na=False)]
        # Convert all comp numbers to string for reliable matching
        comp_numbers = set(filtered_dir['Comp_No'].astype(str).unique())
        
        if not comp_numbers:
            print("⚠️ Warning: No 'PO' or 'MOPO' types found in directory file.")
        else:
            print(f"Found {len(comp_numbers)} relevant Comp_No's to process: {comp_numbers}")
        return comp_numbers
    except Exception as e:
        print(f"❌ ERROR: Failed to filter directory file. Error: {e}")
        return None

def _clean_dataframe(df):
    """Applies standard cleaning rules to a dataframe."""
    df_cleaned = df.copy()
    df_cleaned.replace({r'\n': ' '}, regex=True, inplace=True)
    if 8 < df_cleaned.shape[1]:
        df_cleaned[8] = df_cleaned[8].astype(str).str.replace('#', 'non-holding', regex=False)
    if 4 < df_cleaned.shape[1]:
        df_cleaned[4] = df_cleaned[4].astype(str).replace('Std Industrial', 'Standard Industrial', regex=False)
    
    if len(df_cleaned) > 1:
        def safe_replace(cell_value):
            if isinstance(cell_value, str):
                return cell_value.replace('/', '_').replace(',', '_')
            return cell_value
        for col in df_cleaned.columns:
            df_cleaned.loc[1:, col] = df_cleaned.loc[1:, col].apply(safe_replace)
    return df_cleaned

def _get_headers_and_parts(filename):
    """Parses filename to get headers and name components."""
    base_name = os.path.splitext(filename)[0]
    
    # Regex: Sales_UNIT_COMPNO_MM_YYYY
    match = re.search(r'^Sales_([A-Z0-9]+)_(\d+)_(\d{2})_(\d{4})$', base_name)
    
    if match:
        parts = {
            "unit_name": match.group(1),
            "comp_no": match.group(2),
            "month_part": match.group(3),
            "year_part_full": match.group(4),
            "year_part_short": match.group(4)[-2:],
            "month_int": int(match.group(3)),
            "year_int": int(match.group(4))
        }
    else:
        # Fallback for old filename format: ..._UNIT_MM_YYYY
        match_old = re.search(r'_([A-Z0-9]+)_(\d{2})_(\d{4})$', base_name)
        if match_old:
             parts = {
                "unit_name": match_old.group(1),
                "comp_no": None, # No comp_no in this format
                "month_part": match_old.group(2),
                "year_part_full": match_old.group(3),
                "year_part_short": match_old.group(3)[-2:],
                "month_int": int(match_old.group(2)),
                "year_int": int(match_old.group(3))
            }
        else:
            # Full fallback
            print(f"⚠️  Could not parse filename structure: '{filename}'. Using current date and default name.")
            now = datetime.now()
            parts = {
                "unit_name": "UnknownUnit",
                "comp_no": None,
                "month_part": f"{now.month:02d}",
                "year_part_full": str(now.year),
                "year_part_short": f"{now.year:02d}"[-2:],
                "month_int": now.month,
                "year_int": now.year
            }

    month_abbr = calendar.month_abbr[parts["month_int"]].upper()
    parts["sales_col"] = f"Sales {month_abbr} {parts['year_int']}"
    parts["py_col"] = f"PY Sales {month_abbr} {parts['year_int'] - 1}"
    
    return parts

def _process_segregated_file(df, output_folder, file_parts, base_headers, segregate_types):
    """NEW LOGIC: Processes a file by segregating into 3RD and IC."""
    processed_files = []
    fixed_headers = base_headers + [file_parts["sales_col"], file_parts["py_col"]]
    
    for seg_type in segregate_types:
        print(f"   ... segregating for: {seg_type}")
        
        # 1. Filter for rows where Column L (index 11) is an exact match
        df_segregated = df.loc[df.index.isin([0]) | (df[11].astype(str).str.strip() == seg_type)].copy()
        
        if len(df_segregated) <= 1:
            print(f"   ℹ️  No data found for '{seg_type}'.")
            continue

        # 2. Select final columns (C-K, M, O)
        try:
            final_columns_indices = list(range(2, 11)) + [12, 14] # C-K, M, O
            df_final = df_segregated.iloc[:, final_columns_indices].copy()
        except IndexError:
            print(f"   ❌ Could not process for '{seg_type}'. Missing required columns (C-K, M, O).")
            continue

        # 3. Remove original header row
        df_final = df_final.iloc[1:].copy()
        
        # 4. Assign new headers
        df_final.columns = fixed_headers

        if df_final.empty:
            print(f"   ⚠️  Result for '{seg_type}' is empty after all operations.")
            continue

        # 5. Create new output name and save
        # --- UPDATED ---
        # Add Comp_No (which user calls Profit Center) to the filename
        output_filename = f"Sales_Data_Processed_{file_parts['unit_name']}_{file_parts['comp_no']}_{file_parts['month_part']}{file_parts['year_part_short']}_{seg_type}.csv"
        # --- END UPDATE ---
        output_path = os.path.join(output_folder, output_filename)
        
        df_final.to_csv(output_path, index=False, encoding='utf-8-sig')
        print(f"   ✅ Successfully created: '{output_filename}'")
        processed_files.append(output_filename)
        
    return processed_files

def _process_3rd_only_file(df, output_folder, file_parts, base_headers):
    """OLD LOGIC: Processes a file by filtering for 3RD only."""
    processed_files = []
    print("   ... processing with fallback (3RD only) logic.")
    
    # 1. Filter for rows where Column L (index 11) contains '3RD'.
    #    Using .str.contains("3RD") for legacy compatibility
    df_filtered = df.loc[df.index.isin([0]) | df[11].astype(str).str.contains("3RD", na=False)].copy()
    
    if len(df_filtered) <= 1:
        print(f"   ℹ️  No data found for '3RD'.")
        return []

    # 2. Select final columns (C-K, M, O)
    try:
        final_columns_indices = list(range(2, 11)) + [12, 14]
        df_final = df_filtered.iloc[:, final_columns_indices].copy()
    except IndexError:
        print(f"   ❌ Could not process. Missing required columns (C-K, M, O).")
        return []

    # 3. Remove original header row
    df_final = df_final.iloc[1:].copy()
    
    # 4. Assign new headers
    fixed_headers = base_headers + [file_parts["sales_col"], file_parts["py_col"]]
    df_final.columns = fixed_headers

    if df_final.empty:
        print(f"   ⚠️  Result is empty after all operations.")
        return []

    # 5. Create old output name and save
    # --- UPDATED ---
    # Add comp_no (profit center) to filename if it was parsed
    if file_parts["comp_no"]:
        name_parts = [file_parts['unit_name'], file_parts['comp_no'], f"{file_parts['month_part']}{file_parts['year_part_short']}"]
    else:
        name_parts = [file_parts['unit_name'], f"{file_parts['month_part']}{file_parts['year_part_short']}"]
    
    output_filename = f"Sales_Data_Processed_{'_'.join(name_parts)}.csv"
    # --- END UPDATE ---
    output_path = os.path.join(output_folder, output_filename)
    
    df_final.to_csv(output_path, index=False, encoding='utf-8-sig')
    print(f"   ✅ Successfully created: '{output_filename}'")
    processed_files.append(output_filename)
    
    return processed_files


def process_files_to_csv(input_folder, output_folder, directory_file_path):
    """
    Processes all Excel files from input_folder.
    - If filename Comp_No matches directory file, segregates by '3RD' and 'IC'.
    - Otherwise, filters for '3RD' only.
    """
    print("Script started...")
    processed_files = []

    # --- 1. Get Comp_No's from Directory File ---
    comp_numbers_to_process = get_comp_numbers(directory_file_path)
    if comp_numbers_to_process is None:
        print("Script aborted due to error in directory file processing.")
        return []

    # --- 2. Define Headers ---
    base_headers = [
        "Product/Service", "P1-Division", "P2-DPC", "P3-SBU", "P4-SPG",
        "Customer Group", "Holding", "Distribution Channel", "Sales doc. type"
    ]
    segregate_types = ["3RD", "IC"] # Types to split files by

    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
        print(f"Created output folder: {output_folder}")

    # --- 3. Find all Excel files in input folder ---
    try:
        all_files = [f for f in os.listdir(input_folder) if f.endswith('.xlsx')]
        if not all_files:
            print(f"Warning: No .xlsx files found in {input_folder}")
            return []
    except FileNotFoundError:
        print(f"Error: Input folder not found at {input_folder}. Please check the path.")
        return []

    print(f"Found {len(all_files)} total Excel file(s) to process.")

    # --- 4. Process all files ---
    for filename in all_files:
        input_path = os.path.join(input_folder, filename)
        print(f"\nProcessing file: {filename}")
        
        try:
            # --- A. Load Data ---
            df = pd.read_excel(input_path, header=None, sheet_name="Raw")
            if df.empty:
                print(f"⚠️  Skipping '{filename}' because the 'Raw' sheet is empty.")
                continue
            
            if 11 >= df.shape[1]:
                print(f"⚠️  Skipping '{filename}' because it lacks Column L (index 11) for filtering.")
                continue
                
            # --- B. Clean Data ---
            df_cleaned = _clean_dataframe(df)

            # --- C. Get Filename Parts ---
            file_parts = _get_headers_and_parts(filename)

            # --- D. ROUTE TO CORRECT LOGIC ---
            # Check if comp_no was found AND is in the approved list
            if file_parts["comp_no"] and file_parts["comp_no"] in comp_numbers_to_process:
                print(f"   Found matching Comp_No: {file_parts['comp_no']}. Running segregation logic.")
                new_files = _process_segregated_file(df_cleaned, output_folder, file_parts, base_headers, segregate_types)
                processed_files.extend(new_files)
            else:
                if file_parts["comp_no"]:
                    print(f"   Comp_No {file_parts['comp_no']} not in directory list. Running fallback logic.")
                else:
                    print(f"   No matching Comp_No found in filename. Running fallback logic.")
                new_files = _process_3rd_only_file(df_cleaned, output_folder, file_parts, base_headers)
                processed_files.extend(new_files)

        except Exception as e:
            print(f"❌ Could not process '{filename}'. Error: {e}")

    print(f"\n--- Script finished. Total files created: {len(processed_files)} ---")
    return processed_files

if __name__ == "__main__":
    # --- IMPORTANT ---
    # Replace these paths with your actual folder and file paths
    #
    input_dir = "path/to/your/input_sales_files"
    output_dir = "path/to/your/output_folder"
    directory_file = "path/to/your/Directory_Processed_Output.xlsx"
    #
    # -----------------

    if input_dir == "path/to/your/input_sales_files":
        print("="*60)
        print("⚠️ PLEASE UPDATE THE 'input_dir', 'output_dir',")
        print("   and 'directory_file' variables in the __main__ block")
        print("   at the bottom of the script before running.")
        print("="*60)
    else:
        process_files_to_csv(input_dir, output_dir, directory_file)

