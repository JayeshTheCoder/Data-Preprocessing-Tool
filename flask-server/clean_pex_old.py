import pandas as pd
import shutil
import os
import glob
import re
from pathlib import Path
import calendar  # Added for currency conversion
from datetime import datetime  # Added for currency conversion

# ==============================================================================
# --- SECTION 0: NEW CURRENCY CONVERSION HELPERS ---
# ==============================================================================

def load_directory_info(directory_file_path):
    """
    Reads the Directory_Processed_Output.xlsx file.
    Returns a dictionary mapping Comp_No to its currency info.
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

    required_cols = ["Comp_No", "Original Currency", "Conversion Currency"]
    if not all(col in df_dir.columns for col in required_cols):
        print(f"❌ ERROR: Directory file must contain {required_cols} columns.")
        return None

    try:
        currency_cols = ['Comp_No', 'Original Currency', 'Conversion Currency']
        df_currencies = df_dir[currency_cols].drop_duplicates(subset=['Comp_No']).copy()
        df_currencies['Comp_No'] = df_currencies['Comp_No'].astype(str)
        df_currencies = df_currencies.set_index('Comp_No')
        currency_map = df_currencies.to_dict('index')
        print(f"Loaded currency mapping for {len(currency_map)} Comp_No's.")
        return currency_map
    except Exception as e:
        print(f"❌ ERROR: Failed to build currency map. Error: {e}")
        return None

def load_currency_rates(currency_file_path, month_int, year_int):
    """
    Loads currency rates from the specified file for the given month/year.
    """
    print(f"   ...loading currency rates for {month_int}/{year_int}")
    try:
        month_name = calendar.month_name[month_int]
        current_year_col = f"{month_name} {year_int}"
        prev_year_col = f"{month_name} {year_int - 1}"

        df_headers = pd.read_excel(currency_file_path, header=1, nrows=0)
        
        current_year_col_actual = next((col for col in df_headers.columns if col.strip().lower() == current_year_col.lower()), None)
        prev_year_col_actual = next((col for col in df_headers.columns if col.strip().lower() == prev_year_col.lower()), None)

        if not current_year_col_actual or not prev_year_col_actual:
            print(f"❌ ERROR: Currency file missing required columns. ")
            print(f"   Expected: '{current_year_col}' and '{prev_year_col}'")
            print(f"   Actual headers found: {list(df_headers.columns)}")
            return None

        use_cols = ['Currency', current_year_col_actual, prev_year_col_actual]
        df_rates = pd.read_excel(currency_file_path, header=1, usecols=use_cols)
        
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
        print(f"❌ ERROR: Could not read currency rates file. Error: {e}")
        return None

def get_cross_rates(source_curr, target_curr, rates_dict):
    """
    Calculates the cross rates for current and previous year.
    """
    try:
        target_rate_current = rates_dict[target_curr]['Current_Year_Rate']
        target_rate_prev = rates_dict[target_curr]['Prev_Year_Rate']
        source_rate_current = rates_dict[source_curr]['Current_Year_Rate']
        source_rate_prev = rates_dict[source_curr]['Prev_Year_Rate']

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

def _get_date_from_currency_file(currency_file_path):
    """
    Reads the currency file header to determine the month and year.
    (e.g., from ' September 2025')
    """
    try:
        df_headers = pd.read_excel(currency_file_path, header=1, nrows=0)
        # Get ' September 2025'
        col_cy = df_headers.columns[1].strip() 
        # Parse 'September 2025'
        month_name, year_str = col_cy.split(' ')
        year_int = int(year_str)
        # Convert 'September' to 9
        month_int = datetime.strptime(month_name, '%B').month
        print(f"   ...found date in currency file: {month_name} {year_int}")
        return month_int, year_int
    except Exception as e:
        print(f"❌ ERROR: Could not parse date from currency file header. Error: {e}")
        raise ValueError("Could not parse date from currency file. Check header format.")

# ==============================================================================
# --- SECTION 1: PEX BI & HEADCOUNT PROCESSING (MODIFIED) ---
# ==============================================================================
def generate_pex_output_path(input_path, output_dir):
    try:
        filename, ext = os.path.splitext(os.path.basename(input_path))
        parts = filename.split('_')
        if len(parts) != 5:
            print(f"Warning: PEX filename '{filename}' does not match 'PEX_Unit_Code_MM_YYYY' format.")
            new_filename = f"{filename}_processed{ext}"
        else:
            unit, profit_center, month, year_short = parts[1], parts[2], parts[3], parts[4][-2:]
            new_filename = f"PEX_Data_Processed_{unit}_{profit_center}_{month}{year_short}{ext}"
        return os.path.join(output_dir, new_filename)
    except Exception as e:
        print(f"Error generating PEX filename: {e}")
        base, ext = os.path.splitext(os.path.basename(input_path))
        return os.path.join(output_dir, f"{base}_processed{ext}")

def process_pex_file(input_path, lookup_path, output_folder, currency_map, rates_cache, currency_file_path):
    print(f"--- Starting PEX File Processing for {os.path.basename(input_path)} ---")
    try:
        # --- A. Parse Filename ---
        filename = os.path.basename(input_path)
        parts = os.path.splitext(filename)[0].split('_')
        if len(parts) != 5:
             raise ValueError(f"PEX filename '{filename}' does not match 'PEX_Unit_Code_MM_YYYY' format.")
        
        unit, profit_center, month_num, year_full = parts[1], parts[2], parts[3], parts[4]
        month_int = int(month_num)
        year_int = int(year_full)
        year_prev_full = str(year_int - 1)
        
        output_path = generate_pex_output_path(input_path, output_folder)
        shutil.copy(input_path, output_path)
        print(f"Successfully created a PEX working copy: '{output_path}'")

        # --- B. Get Currency Conversion Rates ---
        cross_rate_current, cross_rate_prev = 1.0, 1.0
        conversion_needed = False
        
        if profit_center in currency_map:
            curr_info = currency_map[profit_center]
            source_curr = curr_info.get('Original Currency')
            target_curr = curr_info.get('Conversion Currency')

            if pd.notna(source_curr) and pd.notna(target_curr) and source_curr != target_curr:
                print(f"   Currency conversion required for {profit_center}: {source_curr} -> {target_curr}")
                
                # Use cache key based on month/year
                rates_key = f"{month_int}-{year_int}"
                if rates_key not in rates_cache:
                    rates_cache[rates_key] = load_currency_rates(currency_file_path, month_int, year_int)
                
                rates_dict = rates_cache[rates_key]

                if rates_dict:
                    rates = get_cross_rates(source_curr, target_curr, rates_dict)
                    if rates[0] is not None:
                        cross_rate_current, cross_rate_prev = rates
                        conversion_needed = True
                        print(f"   Applying rates (Current: *{cross_rate_current:.6f}, PY: *{cross_rate_prev:.6f})")
                    else:
                        print(f"   ❌ ERROR: Could not get cross rates for {source_curr}->{target_curr}.")
                else:
                    print(f"   ❌ ERROR: Could not load currency rates for {rates_key}.")
            else:
                print(f"   No currency conversion needed for {profit_center}.")
        else:
            print(f"   No currency info found for Profit Center: {profit_center}.")


        # --- C. Define Headers ---
        static_part1 = ["Company Code", "Profit Center", "Cost Element", "", "Functional area"]
        df_header_row = pd.read_excel(output_path, sheet_name='Sheet1', nrows=1, header=None)
        dynamic_part1 = list(df_header_row.iloc[0, 5:7]) # e.g., ['Oct 2025', 'Oct 2024']
        static_part2 = ["Actual L3M", "Prior Yr L3M", "Actual YTD", "Prior Yr YTD"]
        dynamic_part2 = list(df_header_row.iloc[0, 11:]) # e.g., ['Sep 2025', 'Aug 2025', ...]
        final_headers = static_part1 + dynamic_part1 + static_part2 + dynamic_part2
        
        # --- D. Read and Clean Data ---
        df = pd.read_excel(output_path, sheet_name='Sheet1', skiprows=2, header=None)
        df.columns = final_headers
        if not df.empty:
            df = df.iloc[:-1]
        
        # --- E. APPLY CURRENCY CONVERSION ---
        if conversion_needed:
            # Identify all CY and PY columns to convert
            cy_cols_to_convert = ["Actual L3M", "Actual YTD"]
            py_cols_to_convert = ["Prior Yr L3M", "Prior Yr YTD"]

            # Add the main dynamic month columns
            cy_cols_to_convert.append(dynamic_part1[0]) # e.g., 'Oct 2025'
            py_cols_to_convert.append(dynamic_part1[1]) # e.g., 'Oct 2024'
            
            # Add all other dynamic month columns
            for header in dynamic_part2:
                if isinstance(header, str):
                    if year_full in header:
                        cy_cols_to_convert.append(header)
                    elif year_prev_full in header:
                        py_cols_to_convert.append(header)
            
            print(f"   Converting {len(cy_cols_to_convert)} CY columns and {len(py_cols_to_convert)} PY columns.")

            # Apply conversion
            for col in cy_cols_to_convert:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0) * cross_rate_current
            
            for col in py_cols_to_convert:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0) * cross_rate_prev
        
        # --- F. Perform Merge/Lookup ---
        df_lookup = pd.read_excel(lookup_path, sheet_name='Sheet4', usecols="B:C", header=None, names=['Cost Element Key', 'Group'])
        df['Cost Element'] = df['Cost Element'].astype(str).str.strip()
        df_lookup['Cost Element Key'] = df_lookup['Cost Element Key'].astype(str).str.strip()
        
        df = pd.merge(df, df_lookup, left_on='Cost Element', right_on='Cost Element Key', how='left')
        df['Group'].fillna('Vehicle Costs', inplace=True)
        df.drop(columns=['Cost Element Key'], inplace=True)
        df.drop_duplicates(inplace=True)
        
        # --- G. Save ---
        with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
            df.to_excel(writer, sheet_name='Sheet1', index=False)
            
        print(f"PEX processing complete! Output saved to: {output_path}")
        
        # Return details for Headcount processing
        return (profit_center, month_num, year_full[-2:], unit), os.path.basename(output_path)
        
    except Exception as e:
        print(f"\n--- An Unexpected Error Occurred during PEX Processing! ---\nError Details: {e}")
        return None, None

def process_headcount_file(headcount_path, output_folder, pex_details):
    # --- This function is unchanged, as Headcount is not currency ---
    print("\n--- Starting Headcount File Processing ---")
    if pex_details is None:
        print("Skipping Headcount processing due to missing PEX details.")
        return None
    profit_center, month_num, year_short, unit = pex_details
    try:
        month_map = {'01':'Jan','02':'Feb','03':'Mar','04':'Apr','05':'May','06':'Jun','07':'Jul','08':'Aug','09':'Sep','10':'Oct','11':'Nov','12':'Dec'}
        month_abbr = month_map.get(month_num, 'Mon')
        sheet_to_read = f"Actual {month_abbr}"
        df = pd.read_excel(headcount_path, sheet_name=sheet_to_read, header=None)
        pc_row = df.iloc[10].astype(str)
        pc_col_series = pc_row[pc_row.str.split('.').str[1] == str(profit_center)]
        if pc_col_series.empty:
            raise ValueError(f"Profit Center '{profit_center}' not found in row 11 of sheet '{sheet_to_read}'.")
        pc_col_index = pc_col_series.index[0]
        custom1_col_index = df.iloc[11][df.iloc[11] == 'Custom1'].index[0]
        df_processed = pd.DataFrame({
            "Account Name": df.iloc[12:, custom1_col_index],
            "Functional Area": df.iloc[12:, 3],
            "CY_Data": df.iloc[12:, pc_col_index],
            "PY_Data": df.iloc[12:, pc_col_index + 1]
        }).reset_index(drop=True)
        df_processed.rename(columns={"CY_Data": f"{year_short}-{month_abbr}", "PY_Data": f"{int(year_short)-1}-{month_abbr}"}, inplace=True)
        output_filename = f"{unit}_{month_num}{year_short}_Headcount_Processed_{profit_center}.xlsx"
        output_path = os.path.join(output_folder, output_filename)
        df_processed.to_excel(output_path, index=False, engine='xlsxwriter')
        print(f"Headcount processing complete! Output saved to: {output_path}")
        return output_filename
    except Exception as e:
        print(f"\n--- An Unexpected Error Occurred during Headcount Processing! ---\nError Details: {e}")
        return None

def process_pex_and_headcount(upload_folder, output_folder, lookup_folder, directory_file_path, currency_file_path):
    """
    Main function for PEX/Headcount, now loads currency info and passes it down.
    """
    try:
        # --- NEW: Load currency info ONCE ---
        print("Loading currency and directory info...")
        currency_map = load_directory_info(directory_file_path)
        if currency_map is None:
             raise FileNotFoundError("Could not load currency directory file.")
        rates_cache = {} # Initialize cache for currency rates
        # --- END NEW ---

        lookup_files = glob.glob(os.path.join(lookup_folder, "PEX validation Template.xlsx"))
        headcount_files = glob.glob(os.path.join(lookup_folder, "HeadcountDatabase.xlsx"))
        if not lookup_files or not headcount_files:
            raise FileNotFoundError("A required lookup file was not found on the server.")
        pex_lookup_file, headcount_input_file = lookup_files[0], headcount_files[0]
        
        pex_input_files = glob.glob(os.path.join(upload_folder, "PEX_*.xlsx"))
        if not pex_input_files:
            pex_input_files = glob.glob(os.path.join(upload_folder, "*", "PEX_*.xlsx"))
            if not pex_input_files: raise FileNotFoundError("No PEX data files found in the upload.")
            
        all_processed_files = []
        for pex_file in pex_input_files:
            # --- MODIFIED: Pass currency info to process_pex_file ---
            pex_details, pex_filename = process_pex_file(
                pex_file, 
                pex_lookup_file, 
                output_folder,
                currency_map,      # <-- NEW
                rates_cache,       # <-- NEW
                currency_file_path # <-- NEW
            )
            if pex_filename: all_processed_files.append(pex_filename)
            
            # Headcount processing is unchanged
            if pex_details:
                headcount_filename = process_headcount_file(headcount_input_file, output_folder, pex_details)
                if headcount_filename: all_processed_files.append(headcount_filename)
                
        return all_processed_files
    except Exception as e:
        print(f"A critical error occurred in process_pex_and_headcount: {e}")
        raise e
    
# ==============================================================================
# --- SECTION 2: PEX VENDOR ANALYSIS PROCESSING (REWRITTEN) ---
# ==============================================================================

def _read_vendor_excel_data(file_path):
    if not file_path: return None
    try:
        print(f"Reading vendor file: {os.path.basename(file_path)}")
        required_cols = ['Cost Element', 'Name of offsetting account', 'Value in Obj. Crcy', 'Offsetting account type']
        df = pd.read_excel(file_path, sheet_name='KSB1', engine='openpyxl')
        df.columns = [str(col).strip() for col in df.columns]
        if not all(col in df.columns for col in required_cols):
            missing = [col for col in required_cols if col not in df.columns]
            print(f"Warning: Missing required columns in {os.path.basename(file_path)}: {missing}")
            return None
        df_filtered = df[df['Offsetting account type'] == 'K']
        df_filtered = df_filtered.dropna(subset=['Cost Element', 'Name of offsetting account', 'Value in Obj. Crcy'], how='all')
        print(f"Extracted {len(df_filtered)} rows from {os.path.basename(file_path)}")
        return df_filtered[required_cols]
    except Exception as e:
        print(f"Error reading {file_path}: {e}")
        return None

def _run_vendor_combination(df_2024, df_2025, output_folder, output_filename,
                            entity_id, currency_map, rates_dict): # <-- NEW ARGS
    if df_2024 is None or df_2025 is None:
        print("Error: Cannot combine data as one of the dataframes is missing.")
        return None
        
    df_2024_agg = df_2024.groupby(['Cost Element', 'Name of offsetting account']).agg({'Value in Obj. Crcy': 'sum'}).reset_index()
    df_2025_agg = df_2025.groupby(['Cost Element', 'Name of offsetting account']).agg({'Value in Obj. Crcy': 'sum'}).reset_index()

    # --- NEW: Apply Currency Conversion ---
    cross_rate_current, cross_rate_prev = 1.0, 1.0
    if entity_id in currency_map:
        curr_info = currency_map[entity_id]
        source_curr = curr_info.get('Original Currency')
        target_curr = curr_info.get('Conversion Currency')

        if pd.notna(source_curr) and pd.notna(target_curr) and source_curr != target_curr:
            rates = get_cross_rates(source_curr, target_curr, rates_dict)
            if rates[0] is not None:
                cross_rate_current = rates[0]
                cross_rate_prev = rates[1]
                print(f"   Applying conversion to group {entity_id} (CY: *{cross_rate_current:.6f}, PY: *{cross_rate_prev:.6f})")
            else:
                print(f"   ❌ ERROR: Could not get cross rates for {source_curr}->{target_curr} for entity {entity_id}.")
    
    # Apply rates (even if 1.0)
    df_2024_agg['Value in Obj. Crcy'] = pd.to_numeric(df_2024_agg['Value in Obj. Crcy'], errors='coerce').fillna(0) * cross_rate_prev
    df_2025_agg['Value in Obj. Crcy'] = pd.to_numeric(df_2025_agg['Value in Obj. Crcy'], errors='coerce').fillna(0) * cross_rate_current
    # --- END NEW ---

    df_2024_renamed = df_2024_agg.rename(columns={'Value in Obj. Crcy': 'Value in Obj. Crcy 2024'})
    df_2025_renamed = df_2025_agg.rename(columns={'Value in Obj. Crcy': 'Value in Obj. Crcy 2025'})
    
    combined_df = pd.merge(df_2024_renamed, df_2025_renamed, on=['Cost Element', 'Name of offsetting account'], how='outer').fillna(0)
    output_path = os.path.join(output_folder, output_filename)
    with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
        combined_df.to_excel(writer, sheet_name='Combined_Vendor_Data', index=False)
    print(f"Successfully saved combined vendor data to {output_path}")
    return os.path.basename(output_path)

def process_pex_vendor(upload_folder, output_folder, bulk_mode, directory_file_path, currency_file_path): # <-- NEW ARGS
    """
    Main entry point for PEX Vendor. In bulk mode, it groups a flat list of files by name.
    """
    
    # --- NEW: Load currency info ONCE ---
    print("Loading currency and directory info for PEX Vendor...")
    currency_map = load_directory_info(directory_file_path)
    if currency_map is None:
         raise FileNotFoundError("Could not load currency directory file.")
    
    # Get date from currency file header, not filename
    month_int, year_int = _get_date_from_currency_file(currency_file_path)
    rates_dict = load_currency_rates(currency_file_path, month_int, year_int)
    if rates_dict is None:
        raise FileNotFoundError("Could not load currency rates.")
    # --- END NEW ---
    
    processed_files = []
    if bulk_mode:
        all_files = glob.glob(os.path.join(upload_folder, "*.xls*"))
        if not all_files:
            raise ValueError("Bulk mode selected, but no Excel files were found in the upload.")

        file_groups = {}
        # Regex captures Unit (UK01), ID (2072), and Year (2025) from filenames
        pattern = re.compile(r'_([A-Z]{2}\d{2})_(\d+)_.*_(\d{4})\.xlsm', re.IGNORECASE)

        for file_path in all_files:
            match = pattern.search(os.path.basename(file_path))
            if match:
                unit, entity_id, year = match.groups()
                group_key = f"{unit}_{entity_id}"
                if group_key not in file_groups:
                    file_groups[group_key] = {}
                file_groups[group_key][year] = file_path

        # Process each complete group (i.e., has both 2024 and 2025 files)
        for group_key, year_files in file_groups.items():
            print(f"\n--- Processing Vendor Group: {group_key} ---")
            file_2024_path = year_files.get('2024')
            file_2025_path = year_files.get('2025')

            if not file_2024_path or not file_2025_path:
                print(f"Warning: Skipping group {group_key} because a 2024 or 2025 file is missing.")
                continue

            df_2024 = _read_vendor_excel_data(file_2024_path)
            df_2025 = _read_vendor_excel_data(file_2025_path)
            
            # --- MODIFIED: Get entity_id and pass to combination function ---
            entity_id = group_key.split('_')[1] # Get '2072' from 'UK01_2072'
            output_filename = f"{group_key}_vendor_analysis_combined.xlsx"
            result_file = _run_vendor_combination(df_2024, df_2025, output_folder, output_filename,
                                                  entity_id, currency_map, rates_dict) # <-- Pass new args
            if result_file:
                processed_files.append(result_file)
    else:
        # Single mode: process exactly two files (2024 and 2025)
        all_files = glob.glob(os.path.join(upload_folder, "*.xls*"))
        if len(all_files) != 2:
            raise ValueError(f"Expected 2 Excel files for vendor analysis, but found {len(all_files)}.")
        
        file_2024 = next((f for f in all_files if "2024" in os.path.basename(f)), None)
        file_2025 = next((f for f in all_files if "2025" in os.path.basename(f)), None)

        if not file_2024 or not file_2025:
            raise FileNotFoundError("Could not identify both a 2024 and a 2025 file in the upload.")
        
        # --- MODIFIED: Get entity_id from filename ---
        pattern = re.compile(r'_([A-Z]{2}\d{2})_(\d+)_.*_(\d{4})\.xlsm', re.IGNORECASE)
        match = pattern.search(os.path.basename(file_2025))
        if not match:
            raise ValueError(f"Could not parse Profit Center/Entity ID from filename: {os.path.basename(file_2025)}")
        entity_id = match.groups()[1] # Get '2072'
        # --- END MODIFIED ---
        
        df_2024 = _read_vendor_excel_data(file_2024)
        df_2025 = _read_vendor_excel_data(file_2025)

        result_file = _run_vendor_combination(df_2024, df_2025, output_folder, "vendor_analysis_combined.xlsx",
                                              entity_id, currency_map, rates_dict) # <-- Pass new args
        if result_file:
            processed_files.append(result_file)

    return processed_files