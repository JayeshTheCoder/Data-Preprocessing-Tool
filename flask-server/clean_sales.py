import pandas as pd
import os
import calendar
from datetime import datetime
import re

def load_directory_info(directory_file_path):
    """
    Reads the Directory_Processed_Output.xlsx file.
    Returns:
    1. A dictionary mapping Comp_No to its Type (e.g., 'MO', 'PO', 'MOPO').
    2. A dictionary mapping Comp_No to its currency info.
    """
    print(f"Reading directory file from: {directory_file_path}")
    try:
        df_dir = pd.read_excel(directory_file_path)
    except FileNotFoundError:
        print(f"❌ ERROR: Directory file not found at: {directory_file_path}")
        return None, None
    except Exception as e:
        print(f"❌ ERROR: Could not read directory file. Error: {e}")
        return None, None

    # --- 1. Check for all required columns ---
    required_cols = ["Comp_No", "Type", "Original Currency", "Conversion Currency"]
    if not all(col in df_dir.columns for col in required_cols):
        print(f"❌ ERROR: Directory file must contain {required_cols} columns.")
        return None, None

    # --- 2. NEW: Create Comp_No to Type Map ---
    comp_type_map = None
    try:
        # Get Comp_No and Type, drop duplicates, convert Comp_No to string
        df_types = df_dir[['Comp_No', 'Type']].drop_duplicates(subset=['Comp_No']).copy()
        df_types['Comp_No'] = df_types['Comp_No'].astype(str)
        df_types['Type'] = df_types['Type'].astype(str).str.strip().str.upper() # Clean up type field
        
        # Filter for only the types we care about
        valid_types = ['PO', 'MO', 'MOPO']
        df_types_filtered = df_types[df_types['Type'].isin(valid_types)]
        
        df_types_filtered = df_types_filtered.set_index('Comp_No')
        comp_type_map = df_types_filtered['Type'].to_dict()
        
        if not comp_type_map:
            print("⚠️ Warning: No 'PO', 'MO', or 'MOPO' types found in directory file.")
        else:
            print(f"Found {len(comp_type_map)} relevant Comp_No's for processing: {comp_type_map}")
    except Exception as e:
        print(f"❌ ERROR: Failed to build Comp_No -> Type map. Error: {e}")
        return None, None
    # --- END NEW ---

    # --- 3. Create Currency Map ---
    try:
        currency_cols = ['Comp_No', 'Original Currency', 'Conversion Currency']
        df_currencies = df_dir[currency_cols].drop_duplicates(subset=['Comp_No']).copy()
        df_currencies['Comp_No'] = df_currencies['Comp_No'].astype(str)
        df_currencies = df_currencies.set_index('Comp_No')
        currency_map = df_currencies.to_dict('index')
        print(f"Loaded currency mapping for {len(currency_map)} Comp_No's.")
    except Exception as e:
        print(f"❌ ERROR: Failed to build currency map. Error: {e}")
        return None, None

    return comp_type_map, currency_map

def load_currency_rates(currency_file_path, month_int, year_int):
    """
    Loads currency rates from the specified file for the given month/year.
    
    --- UPDATED ---
    Dynamically reads from the correct sheet based on the month (e.g., 'Sep', 'Oct').
    """
    try:
        # Get the 3-letter month abbreviation (e.g., 9 -> 'Sep')
        month_abbr = calendar.month_abbr[month_int]
        
        print(f"   ... loading currency rates for {month_int}/{year_int} from sheet '{month_abbr}'")

        # Get the full month name (e.g., "September")
        month_name = calendar.month_name[month_int]
        
        # Define the expected column headers
        current_year_col = f"{month_name} {year_int}"
        prev_year_col = f"{month_name} {year_int - 1}"

        # --- MODIFIED: Use sheet_name=month_abbr ---
        df_headers = pd.read_excel(currency_file_path, sheet_name=month_abbr, header=1, nrows=0)

        # Find matching columns case-insensitively and stripping spaces
        current_year_col_actual = next((col for col in df_headers.columns if col.strip().lower() == current_year_col.lower()), None)
        prev_year_col_actual = next((col for col in df_headers.columns if col.strip().lower() == prev_year_col.lower()), None)
        # --- END UPDATE ---

        if not current_year_col_actual or not prev_year_col_actual:
            print(f"❌ ERROR: Currency file (sheet '{month_abbr}') missing required columns. ")
            print(f"   Expected: '{current_year_col}' and '{prev_year_col}'")
            print(f"   Actual headers found: {list(df_headers.columns)}")
            return None

        # Read the actual data using the found column names
        use_cols = ['Currency', current_year_col_actual, prev_year_col_actual]
        
        # --- MODIFIED: Use sheet_name=month_abbr ---
        df_rates = pd.read_excel(currency_file_path, sheet_name=month_abbr, header=1, usecols=use_cols)
        
        # Standardize column names
        df_rates.rename(columns={
            current_year_col_actual: 'Current_Year_Rate',
            prev_year_col_actual: 'Prev_Year_Rate'
        }, inplace=True)

        df_rates = df_rates.dropna(subset=['Currency'])
        df_rates['Currency'] = df_rates['Currency'].str.strip()
        df_rates = df_rates.set_index('Currency')
        
        return df_rates.to_dict('index')

    except FileNotFoundError:
        print(f"❌ ERROR: Currency rates file not found at: {currency_file_path}")
        return None
    except Exception as e:
        # Catch errors if the sheet (e.g., 'Sep') doesn't exist
        if "No sheet named" in str(e):
             print(f"❌ ERROR: Could not find sheet '{month_abbr}' in currency file.")
             return None
        print(f"❌ ERROR: Could not read currency rates file. Error: {e}")
        return None

def get_cross_rates(source_curr, target_curr, rates_dict):
    """
    Calculates the cross rates for current and previous year.
    Returns tuple: (cross_rate_current, cross_rate_prev)
    """
    try:
        # 1. Get rates for Target Currency (e.g., SEK)
        target_rate_current = rates_dict[target_curr]['Current_Year_Rate']
        target_rate_prev = rates_dict[target_curr]['Prev_Year_Rate']
        
        # 2. Get rates for Source Currency (e.g., DKK)
        source_rate_current = rates_dict[source_curr]['Current_Year_Rate']
        source_rate_prev = rates_dict[source_curr]['Prev_Year_Rate']

        # 3. Calculate Cross Rate (Target / Source)
        cross_rate_current = target_rate_current / source_rate_current
        cross_rate_prev = target_rate_prev / source_rate_prev
        
        return cross_rate_current, cross_rate_prev

    except KeyError as e:
        print(f"   ❌ ERROR: Currency '{e.args[0]}' not found in rates file.")
        return None, None
    except ZeroDivisionError:
        print(f"   ❌ ERROR: Cannot divide by zero. Source currency '{source_curr}' has a rate of 0.")
        return None, None
    except Exception as e:
        print(f"   ❌ ERROR: Failed to calculate cross rate. Error: {e}")
        return None, None

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
    match = re.search(r'^Sales_([A-Z0-9]+)_(\d+)_(\d{2})_(\d{4})$', base_name, re.IGNORECASE)
    
    if match:
        parts = {
            "unit_name": match.group(1).upper(),
            "comp_no": match.group(2),
            "month_part": match.group(3),
            "year_part_full": match.group(4),
            "year_part_short": match.group(4)[-2:],
            "month_int": int(match.group(3)),
            "year_int": int(match.group(4))
        }
    else:
        # Fallback for old filename format: ..._UNIT_MM_YYYY
        match_old = re.search(r'_([A-Z0-9]+)_(\d{2})_(\d{4})$', base_name, re.IGNORECASE)
        if match_old:
            parts = {
                "unit_name": match_old.group(1).upper(),
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

def _process_segregated_file(df, output_folder, file_parts, base_headers, types_to_create, cross_rate_current=1.0, cross_rate_prev=1.0):
    """
    NEW LOGIC: Processes a file by segregating into the types specified in `types_to_create`.
    (e.g., ['3RD'], ['IC'], or ['3RD', 'IC'])
    """
    processed_files = []
    fixed_headers = base_headers + [file_parts["sales_col"], file_parts["py_col"]]
    
    for seg_type in types_to_create:
        print(f"   ... segregating for: {seg_type}")
        
        # 1. Filter for rows where Column L (index 11) is an exact match
        # We must keep row 0 (headers) for .iloc logic, then drop it
        df_segregated = df.loc[df.index.isin([0]) | (df[11].astype(str).str.strip() == seg_type)].copy()
        
        if len(df_segregated) <= 1:
            print(f"   ℹ️   No data found for '{seg_type}'.")
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
        
        # --- 4. NEW: Apply Currency Conversion ---
        if cross_rate_current != 1.0 or cross_rate_prev != 1.0:
            print(f"       ... applying currency conversion (Current: *{cross_rate_current:.6f}, PY: *{cross_rate_prev:.6f})")
            # Col 9 = Sales (from index 12=M), Col 10 = PY Sales (from index 14=O)
            df_final.iloc[:, 9] = pd.to_numeric(df_final.iloc[:, 9], errors='coerce').fillna(0) * cross_rate_current
            df_final.iloc[:, 10] = pd.to_numeric(df_final.iloc[:, 10], errors='coerce').fillna(0) * cross_rate_prev
        # --- END NEW ---

        # 5. Assign new headers
        df_final.columns = fixed_headers

        if df_final.empty:
            print(f"   ⚠️   Result for '{seg_type}' is empty after all operations.")
            continue

        # 6. Create new output name and save (NEW NAMING CONVENTION)
        output_filename = f"Sales_Data_Processed_{file_parts['unit_name']}_{file_parts['comp_no']}_{file_parts['month_part']}{file_parts['year_part_short']}_{seg_type}.csv"
        output_path = os.path.join(output_folder, output_filename)
        
        df_final.to_csv(output_path, index=False, encoding='utf-8-sig')
        print(f"   ✅ Successfully created: '{output_filename}'")
        processed_files.append(output_filename)
        
    return processed_files

def _process_3rd_only_file(df, output_folder, file_parts, base_headers, cross_rate_current=1.0, cross_rate_prev=1.0):
    """FALLBACK LOGIC: Processes a file by filtering for 3RD only."""
    processed_files = []
    print("   ... processing with fallback (3RD only) logic.")
    
    # 1. Filter for rows where Column L (index 11) contains '3RD'.
    df_filtered = df.loc[df.index.isin([0]) | df[11].astype(str).str.contains("3RD", na=False)].copy()
    
    if len(df_filtered) <= 1:
        print(f"   ℹ️   No data found for '3RD'.")
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
    
    # --- 4. NEW: Apply Currency Conversion ---
    if cross_rate_current != 1.0 or cross_rate_prev != 1.0:
        print(f"       ... applying currency conversion (Current: *{cross_rate_current:.6f}, PY: *{cross_rate_prev:.6f})")
        # Col 9 = Sales (from index 12=M), Col 10 = PY Sales (from index 14=O)
        df_final.iloc[:, 9] = pd.to_numeric(df_final.iloc[:, 9], errors='coerce').fillna(0) * cross_rate_current
        df_final.iloc[:, 10] = pd.to_numeric(df_final.iloc[:, 10], errors='coerce').fillna(0) * cross_rate_prev
    # --- END NEW ---
    
    # 5. Assign new headers
    fixed_headers = base_headers + [file_parts["sales_col"], file_parts["py_col"]]
    df_final.columns = fixed_headers

    if df_final.empty:
        print(f"   ⚠️   Result is empty after all operations.")
        return []

    # 6. Create old output name and save
    if file_parts["comp_no"]:
        name_parts = [file_parts['unit_name'], file_parts['comp_no'], f"{file_parts['month_part']}{file_parts['year_part_short']}"]
    else:
        name_parts = [file_parts['unit_name'], f"{file_parts['month_part']}{file_parts['year_part_short']}"]
    
    output_filename = f"Sales_Data_Processed_{'_'.join(name_parts)}.csv"
    output_path = os.path.join(output_folder, output_filename)
    
    df_final.to_csv(output_path, index=False, encoding='utf-8-sig')
    print(f"   ✅ Successfully created: '{output_filename}'")
    processed_files.append(output_filename)
    
    return processed_files


def process_files_to_csv(input_folder, output_folder, directory_file_path, currency_rates_file_path):
    """
    Processes all Excel files from input_folder.
    - Applies currency conversion based on directory file.
    - If Comp_No Type is 'MO', creates '3RD' file.
    - If Comp_No Type is 'PO', creates 'IC' file.
    - If Comp_No Type is 'MOPO', creates '3RD' AND 'IC' files.
    - Otherwise, filters for '3RD' only using fallback logic.
    """
    print("Script started...")
    processed_files = []
    rates_cache = {} # Cache loaded currency files

    # --- 1. Get Comp_No Type Map & Currency Map ---
    comp_type_map, currency_map = load_directory_info(directory_file_path)
    if comp_type_map is None:
        print("Script aborted due to error in directory file processing.")
        return []

    # --- 2. Define Headers ---
    base_headers = [
        "Product/Service", "P1-Division", "P2-DPC", "P3-SBU", "P4-SPG",
        "Customer Group", "Holding", "Distribution Channel", "Sales doc. type"
    ]

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
                print(f"⚠️   Skipping '{filename}' because the 'Raw' sheet is empty.")
                continue
            
            if 14 >= df.shape[1]:
                print(f"⚠️   Skipping '{filename}' because it lacks Column O (index 14) for PY Sales.")
                continue
                
            # --- B. Clean Data ---
            df_cleaned = _clean_dataframe(df)

            # --- C. Get Filename Parts ---
            file_parts = _get_headers_and_parts(filename)

            # --- D. Check for Currency Conversion ---
            cross_rate_current, cross_rate_prev = 1.0, 1.0
            comp_no = file_parts["comp_no"]

            if comp_no and comp_no in currency_map:
                curr_info = currency_map[comp_no]
                source_curr = curr_info.get('Original Currency')
                target_curr = curr_info.get('Conversion Currency')

                if pd.notna(source_curr) and pd.notna(target_curr) and source_curr != target_curr:
                    print(f"   Currency conversion required: {source_curr} -> {target_curr}")
                    
                    # Use cache key based on month/year
                    rates_key = f"{file_parts['month_int']}-{file_parts['year_int']}"
                    
                    if rates_key not in rates_cache:
                        print(f"   ... loading currency rates for {rates_key}")
                        rates_cache[rates_key] = load_currency_rates(currency_rates_file_path, file_parts["month_int"], file_parts["year_int"])
                    
                    rates_dict = rates_cache[rates_key]

                    if rates_dict:
                        rates = get_cross_rates(source_curr, target_curr, rates_dict)
                        if rates[0] is not None:
                            cross_rate_current, cross_rate_prev = rates
                        else:
                            print(f"   ❌ ERROR: Could not get cross rates for {source_curr}->{target_curr}. Skipping conversion.")
                    else:
                        print(f"   ❌ ERROR: Could not load currency rates for {rates_key}. Skipping conversion.")
                else:
                    print(f"   No currency conversion needed (Source: {source_curr}, Target: {target_curr}).")
            else:
                 print(f"   No currency info found for Comp_No: {comp_no}.")
            # --- END D ---


            # --- E. NEW: ROUTE TO CORRECT LOGIC ---
            types_to_create = []
            process_with_segregation = False
            comp_type = None

            if comp_no and comp_no in comp_type_map:
                comp_type = comp_type_map[comp_no]
                if comp_type == 'MO':
                    types_to_create = ['3RD']
                    process_with_segregation = True
                elif comp_type == 'PO':
                    types_to_create = ['IC']
                    process_with_segregation = True
                elif comp_type == 'MOPO':
                    types_to_create = ['3RD', 'IC']
                    process_with_segregation = True
                else:
                    # This case should be filtered by load_directory_info, but as a safeguard:
                    print(f"   Unknown type '{comp_type}' for Comp_No {comp_no}. Using fallback.")

            
            if process_with_segregation:
                print(f"   Found matching Comp_No: {comp_no} (Type: {comp_type}). Running new logic.")
                new_files = _process_segregated_file(
                    df_cleaned, output_folder, file_parts, base_headers, 
                    types_to_create,  # <-- Pass the dynamic list
                    cross_rate_current, cross_rate_prev
                )
                processed_files.extend(new_files)
            else:
                if comp_no:
                    print(f"   Comp_No {comp_no} not in 'MO/PO/MOPO' list. Running fallback logic.")
                else:
                    print(f"   No matching Comp_No found in filename. Running fallback logic.")
                new_files = _process_3rd_only_file(
                    df_cleaned, output_folder, file_parts, base_headers, 
                    cross_rate_current, cross_rate_prev
                )
                processed_files.extend(new_files)
            # --- END E ---

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
    currency_file = "path/to/your/curre sep 2025.xlsx" # <-- NEW FILE PATH
    #
    # -----------------

    if input_dir == "path/to/your/input_sales_files":
        print("="*60)
        print("⚠️ PLEASE UPDATE THE 'input_dir', 'output_dir',")
        print("   'directory_file', and 'currency_file' variables")
        print("   in the __main__ block at the bottom of the script.")
        print("="*60)
    else:
        process_files_to_csv(input_dir, output_dir, directory_file, currency_file)