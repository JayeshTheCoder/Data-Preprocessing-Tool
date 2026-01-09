import pandas as pd
import os
import re
import glob
import hashlib 

# ==============================================================================
# --- NEW: Grouping/PC Lookup Helpers (v4) ---
# ==============================================================================

def _load_group_to_pc_map(directory_file_path, group_col='Grouping Unit', pc_col='Comp_No_for_OE'):
    """
    Loads the directory and creates a REVERSE map from a grouping column
    to a profit center column.
    
    e.g., {'Nordic': '9001', 'MOPO': '9002'}
    
    This assumes the directory file contains a summary row for each group
    that maps the group name to a single summary Profit Center.
    """
    print(f"   - Loading Group-to-PC map from {os.path.basename(directory_file_path)}...")
    try:
        df_dir = pd.read_excel(directory_file_path)
        
        # Ensure we have the columns we need
        if group_col not in df_dir.columns or pc_col not in df_dir.columns:
            print(f"     - ❌ ERROR: Directory missing required columns: '{group_col}' or '{pc_col}'.")
            return {}
            
        # Drop rows where either value is null
        df_dir = df_dir.dropna(subset=[group_col, pc_col])
        
        # Convert PC to string for consistency
        df_dir[pc_col] = df_dir[pc_col].astype(str).str.strip()
        df_dir[group_col] = df_dir[group_col].astype(str).str.strip()

        # Drop duplicates based on the group name to get a 1-to-1 map
        # This takes the *first* PC found for that group name.
        df_map = df_dir.drop_duplicates(subset=[group_col])
        
        group_map = pd.Series(df_map[pc_col].values, index=df_map[group_col]).to_dict()
        print(f"     - Successfully loaded map for {len(group_map)} groups.")
        return group_map
        
    except Exception as e:
        print(f"     - ❌ ERROR loading Group-to-PC map: {e}")
        return {}

def _get_pc_and_group_from_filename(filename, group_to_pc_map, single_file_regex, grouped_file_regex):
    """
    Parses a filename to find the correct Profit Center to check in Hyperion.
    
    - If it's a single-entity file, returns that entity's PC.
    - If it's a grouped file, returns the group's summary PC from the map.
    
    Returns: tuple (pc_to_check, is_grouped_file)
    e.g., ("2175", False) or ("9001", True)
    """
    
    # 1. Try to parse as a single-entity file
    #    e.g., ..._CH01_2175_1025...
    match_single = re.search(single_file_regex, filename, re.IGNORECASE)
    if match_single:
        pc_str = match_single.group(1).strip()
        print(f"   - Parsed as Single-Entity file. Using PC: {pc_str}")
        return (pc_str, False)
        
    # 2. Try to parse as a grouped file
    #    e.g., ..._Processed_Nordic_1025...
    match_grouped = re.search(grouped_file_regex, filename, re.IGNORECASE)
    if match_grouped:
        group_name = match_grouped.group(1).strip()
        
        # 3. Look up the group name in our reverse map
        pc_for_group = group_to_pc_map.get(group_name)
        
        if pc_for_group:
            print(f"   - Parsed as Grouped file '{group_name}'. Using summary PC: {pc_for_group}")
            return (pc_for_group, True)
        else:
            print(f"   - ⚠️ Parsed Grouped file '{group_name}', but no summary PC found in directory map.")
            return (None, True)
            
    # 4. Failed to parse
    print(f"   - ❌ SKIPPING validation: Could not parse filename '{filename}' with known patterns.")
    return (None, False)

# ==============================================================================
# ---                   Sales Hyperion Validation Function                   ---
# ==============================================================================

def generate_sales_validation_data(processed_file_path, hyperion_3rd_file_path, hyperion_IC_file_path, directory_file_path):
    """
    Validates a processed Sales file (3RD or IC) against a Sales Hyperion
    validation template.
    
    (v5 Update):
    - Appends file type (3RD or IC) to the output sheet name.
    - Uses a helper to get the correct PC from the filename (single or group).
    
    Returns a dictionary of {profit_center_TYPE: validation_dataframe}
    """
    
    filename = os.path.basename(processed_file_path)
    print(f"--- Generating Sales Hyperion Validation Data for {filename} ---")
    
    all_validation_sheets = {}
    
    month_num_to_abbr = {
        '01': 'Jan', '02': 'Feb', '03': 'Mar', '04': 'Apr', '05': 'May', '06': 'Jun',
        '07': 'Jul', '08': 'Aug', '09': 'Sep', '10': 'Oct', '11': 'Nov', '12': 'Dec'
    }

    # --- 1. Determine File Type, Comp_No, and Month from Filename ---
    
    # --- A. Load the Group-to-PC map ---
    group_to_pc_map = _load_group_to_pc_map(directory_file_path, 'Grouping Unit', 'Comp_No_for_OE')

    # --- B. Define regex patterns ---
    single_regex = r'_[A-Z0-9]+_(\d+)_\d{4}_(3RD|IC)\.csv'
    grouped_regex = r'_Processed_([A-Za-z\s()-]+)_\d{4}_(3RD|IC)\.csv'

    # --- C. Get the PC to check ---
    (pc_str, is_grouped) = _get_pc_and_group_from_filename(filename, group_to_pc_map, single_regex, grouped_regex)
    
    if not pc_str:
        return all_validation_sheets # Helper function already printed the error
        
    # --- D. Determine which Hyperion file to use ---
    hyperion_file_to_use = None
    date_part = None # e.g., "1025"
    file_type_str = None # <-- To store "3RD" or "IC"

    if '_3RD' in filename.upper():
        hyperion_file_to_use = hyperion_3rd_file_path
        file_type_str = "3RD" # <-- Capture the type
        print(f"   - Detected '3rd Party' (3RD) file.")
        match = re.search(r'_(\d{4})_3RD\.csv', filename, re.IGNORECASE)
        if match: date_part = match.group(1)
        
    elif '_IC' in filename.upper():
        hyperion_file_to_use = hyperion_IC_file_path
        file_type_str = "IC" # <-- Capture the type
        print(f"   - Detected 'Intercompany' (IC) file.")
        match = re.search(r'_(\d{4})_IC\.csv', filename, re.IGNORECASE)
        if match: date_part = match.group(1)
    
    # Check for file_type_str as well
    if not hyperion_file_to_use or not date_part or not file_type_str:
        print(f"   - SKIPPING: Could not determine file type (3RD/IC) or date_part from filename.")
        return all_validation_sheets
        
    month_mm = date_part[0:2]
    month_abbr = month_num_to_abbr.get(month_mm, "Sheet1") # e.g., "Oct"
    print(f"   - Target Sheet: {month_abbr}")

    if not os.path.exists(hyperion_file_to_use):
        print(f"   - SKIPPING validation: Hyperion file not found: {os.path.basename(hyperion_file_to_use)}")
        return all_validation_sheets
        
    print(f"   - Using Hyperion validation file: {os.path.basename(hyperion_file_to_use)}")

    try:
        # --- 2. Get Data from Hyperion (The "Truth") ---
        # (This section is unchanged)
        try:
            df_hyperion = pd.read_excel(hyperion_file_to_use, sheet_name=month_abbr, header=None)
        except Exception as e:
            if month_abbr != "Sheet1":
                print(f"   - WARNING: Could not find sheet '{month_abbr}'. Trying 'Sheet1'... Error: {e}")
                try:
                    df_hyperion = pd.read_excel(hyperion_file_to_use, sheet_name='Sheet1', header=None)
                except Exception as e2:
                    print(f"   - ERROR: Could not read sheet '{month_abbr}' or 'Sheet1'. {e2}")
                    return all_validation_sheets
            else:
                 print(f"   - ERROR: Could not read sheet 'Sheet1'. {e}")
                 return all_validation_sheets

        pc_row_hyperion = df_hyperion.iloc[6] # 7th row
        df_hyperion_data_part = df_hyperion.iloc[11:] # Data from 12th row
        
        target_col_idx_hyperion = None
        for idx, val in pc_row_hyperion.items():
            if isinstance(val, str) and val.strip().endswith(f'.{pc_str}'):
                target_col_idx_hyperion = idx
                break
                
        if target_col_idx_hyperion is None:
            print(f"     - WARNING: Profit Center {pc_str} not found in Hyperion file's 7th row. Skipping.")
            return all_validation_sheets
            
        prior_col_idx_hyperion = target_col_idx_hyperion + 1
        
        df_hyperion_lookup = pd.DataFrame({
            'ProdSvc': df_hyperion_data_part.iloc[:, 1].astype(str),  # Col B
            'Division': df_hyperion_data_part.iloc[:, 2].astype(str), # Col C
            'DPC': df_hyperion_data_part.iloc[:, 3].astype(str),      # Col D
            'Actual': pd.to_numeric(df_hyperion_data_part.iloc[:, target_col_idx_hyperion], errors='coerce').fillna(0),
            'Prior': pd.to_numeric(df_hyperion_data_part.iloc[:, prior_col_idx_hyperion], errors='coerce').fillna(0)
        })
        
        df_hyperion_agg = df_hyperion_lookup.groupby(['ProdSvc', 'Division', 'DPC'])[['Actual', 'Prior']].sum()
        hyperion_map = df_hyperion_agg.apply(lambda row: (row['Actual'], row['Prior']), axis=1).to_dict()
        print(f"     ... Built Hyperion map for PC {pc_str} with {len(hyperion_map)} unique keys.")

        # --- 3. Get Data from BI File (The "Test") ---
        # (This section is unchanged)
        df_bi = pd.read_csv(processed_file_path)
        
        bi_group_cols = ['Product/Service', 'P1-Division', 'P2-DPC']
        if not all(col in df_bi.columns for col in bi_group_cols):
            print(f"   - SKIPPING validation: BI file '{filename}' is missing required grouping columns.")
            return all_validation_sheets
            
        actual_col_bi = next((c for c in df_bi.columns if c.startswith('Sales ') and c.lower() != 'sales doc. type'), None)
        prior_col_bi = next((c for c in df_bi.columns if c.startswith('PY Sales ')), None)
        
        if not actual_col_bi or not prior_col_bi:
            print(f"   - SKIPPING validation: BI file '{filename}' is missing dynamic sales columns.")
            return all_validation_sheets
            
        year_actual_match = re.search(r'(\d{4})$', actual_col_bi)
        year_prior_match = re.search(r'(\d{4})$', prior_col_bi)
        year_actual_str = year_actual_match.group(1) if year_actual_match else "YYYY"
        year_prior_str = year_prior_match.group(1) if year_prior_match else "YYYY-1"
        
        print(f"   - Validating BI columns: '{actual_col_bi}' ({year_actual_str}) and '{prior_col_bi}' ({year_prior_str})")
        
        for col in bi_group_cols:
            df_bi[col] = df_bi[col].astype(str)
            
        df_bi_grouped = df_bi.groupby(bi_group_cols)[[actual_col_bi, prior_col_bi]].sum()
        bi_map = df_bi_grouped.apply(lambda row: (row[actual_col_bi], row[prior_col_bi]), axis=1).to_dict()
        print(f"     ... Built BI (Cleaned File) map with {len(bi_map)} unique keys.")

        # --- 4. Compare the two maps ---
        # (This section is unchanged)
        validation_data = []
        all_groups_bi = set(bi_map.keys())
        all_groups_hyperion = set(hyperion_map.keys())
        all_groups = sorted(list(all_groups_bi | all_groups_hyperion))
        
        print(f"     - Comparing {len(all_groups)} unique 'Group' items for PC {pc_str}.")
        
        for group_key in all_groups:
            if not isinstance(group_key, tuple) or len(group_key) != 3: continue
            
            prod_svc, division, dpc = group_key
            
            bi_vals = bi_map.get(group_key, (0, 0))
            bi_actual = float(bi_vals[0])
            bi_prior = float(bi_vals[1])
            
            hyperion_vals = hyperion_map.get(group_key, (0, 0)) 
            hyperion_actual = float(hyperion_vals[0]) * 1000
            hyperion_prior = float(hyperion_vals[1]) * 1000
            
            difference_actual = hyperion_actual - bi_actual
            difference_prior = hyperion_prior - bi_prior
            variance = 5.0
            status_actual = "Matching" if abs(difference_actual) <= variance else "Not Matching"
            status_prior = "Matching" if abs(difference_prior) <= variance else "Not Matching"
            
            validation_data.append({
                'Product/Service': prod_svc,
                'Division': division,
                'DPC': dpc,
                f'BI {actual_col_bi}': bi_actual,
                f'BI {prior_col_bi}': bi_prior,
                f'Hyperion {year_actual_str}': hyperion_actual,
                f'Hyperion {year_prior_str}': hyperion_prior,
                f'Difference {year_actual_str}': difference_actual,
                f'Difference {year_prior_str}': difference_prior,
                f'Status {year_actual_str}': status_actual,
                f'Status {year_prior_str}': status_prior
            })
        
        df_validation = pd.DataFrame(validation_data)
        
        # --- [MODIFICATION IS HERE] ---
        # Create the new sheet name, e.g., "1024_IC" or "Nordic_3RD"
        new_sheet_name = f"{pc_str}_{file_type_str}" 
        
        # We return the sheet named after the PC/Group and Type
        all_validation_sheets[new_sheet_name] = df_validation
        print(f"     - Successfully generated validation data for sheet: {new_sheet_name}.")
        # --- [END MODIFICATION] ---
        
        print(f"   - Finished generating data for all sheets.")
        return all_validation_sheets

    except Exception as e:
        print(f"   - ❌ ERROR during Sales Hyperion data generation: {e}")
        import traceback
        traceback.print_exc()
        return all_validation_sheets
# ==============================================================================
# ---                   OE Hyperion Validation Function (MODIFIED)           ---
# ==============================================================================

def generate_oe_validation_data(processed_file_path, hyperion_oe_folder_path, directory_file_path):
    """
    Validates a processed Order Entry (OE) file (CSV) against a Hyperion
    validation template.
    
    (v4 Update):
    - Skips adding the 'Adjustment figure' row to the final validation output.
    
    (v3 Update):
    - Implements OE-specific filename parsing.
    - Parses Comp_No from filename (e.g., '5231').
    - Uses a map to find the *actual* Comp_No_for_OE (e.g., '9005').
    - Uses this Comp_No_for_OE to check against Hyperion.
    - For grouped files, finds the group's summary PC.
    """
    
    filename = os.path.basename(processed_file_path)
    print(f"--- Generating OE Hyperion Validation Data for {filename} ---")
    
    all_validation_sheets = {}
    
    month_num_to_abbr = {
        '01': 'Jan', '02': 'Feb', '03': 'Mar', '04': 'Apr', '05': 'May', '06': 'Jun',
        '07': 'Jul', '08': 'Aug', '09': 'Sep', '10': 'Oct', '11': 'Nov', '12': 'Dec'
    }

    try:
        # --- [NEW] Load required maps ---
        # Map 1: For finding the OE_PC from the filename's PC
        comp_no_to_oe_map = _load_comp_no_to_oe_map(directory_file_path)
        if not comp_no_to_oe_map:
            print("   - ❌ SKIPPING: Could not load Comp_No -> Comp_No_for_OE map.")
            return all_validation_sheets

        # Map 2: For finding a group's summary PC (if it's a grouped file)
        group_to_pc_map = _load_group_to_pc_map(directory_file_path, 'Grouping Unit', 'Comp_No_for_OE')
        if not group_to_pc_map:
            print("   - ❌ SKIPPING: Could not load Group -> Comp_No_for_OE map.")
            return all_validation_sheets
        # --- [END NEW] ---


        # --- [MODIFIED] 1. Get the correct Profit Center to check ---
        
        # --- A. Define regex patterns for OE files ---
        # e.g., OE_Data_Processed_DK01_2033_0925.csv
        single_regex = r'_[A-Z0-9]+_(\d+)_\d{4}(\(\d+\))?\.csv'
        # e.g., OE_Data_Processed_Nordic_0925.csv
        grouped_regex = r'_Processed_([A-Za-z\s()-]+)_\d{4}(\(\d+\))?\.csv'

        pc_str = None # This is the final PC to check in Hyperion
        
        # --- B. Try to parse as a single-entity file ---
        match_single = re.search(single_regex, filename, re.IGNORECASE)
        if match_single:
            pc_from_file = match_single.group(1).strip() # e.g., '5231'
            # Now, find the *actual* OE PC
            pc_str = comp_no_to_oe_map.get(pc_from_file)
            if pc_str:
                print(f"   - Parsed as Single-Entity file. Using Comp_No '{pc_from_file}' -> Mapped to OE PC: {pc_str}")
            else:
                print(f"   - ❌ SKIPPING: Parsed Comp_No '{pc_from_file}', but not found in directory's 'Comp_No' column.")
                return all_validation_sheets
        
        # --- C. Try to parse as a grouped file ---
        if not pc_str:
            match_grouped = re.search(grouped_regex, filename, re.IGNORECASE)
            if match_grouped:
                group_name = match_grouped.group(1).strip() # e.g., 'Nordic'
                # Find the group's summary PC
                pc_str = group_to_pc_map.get(group_name)
                if pc_str:
                    print(f"   - Parsed as Grouped file '{group_name}'. Using summary PC: {pc_str}")
                else:
                    print(f"   - ❌ SKIPPING: Parsed Grouped file '{group_name}', but no summary PC found in directory map.")
                    return all_validation_sheets

        if not pc_str:
            print(f"   - ❌ SKIPPING validation: Could not parse filename '{filename}' with OE patterns.")
            return all_validation_sheets
        # --- [END MODIFIED SECTION] ---

        # --- D. Get Month/Year from filename ---
        date_part = None
        match = re.search(r'_(\d{4})(\(\d+\))?\.csv', filename, re.IGNORECASE)
        if match:
            date_part = match.group(1)
        
        if not date_part:
            print(f"   - SKIPPING: Could not determine date_part (e.g., '0925') from filename.")
            return all_validation_sheets

        month_mm = date_part[0:2]
        month_abbr = month_num_to_abbr.get(month_mm, "Sheet1") # e.g., "Sep"
        print(f"   - Target Sheet: {month_abbr}")

        # --- 2. Load Processed BI File (Test Data) ---
        df_bi = pd.read_csv(processed_file_path)
        
        bi_group_col = 'P2-DPC'
        bi_service_col = 'Product/Service'
        bi_mtd_col = 'Bookings MTD Net Sales'
        bi_py_col = 'Bookings PY MTD'
        
        required_cols = [bi_group_col, bi_service_col, bi_mtd_col, bi_py_col]
        if not all(col in df_bi.columns for col in required_cols):
            print(f"   - SKIPPING validation: BI file '{filename}' is missing required columns.")
            return all_validation_sheets

        # Prep BI data
        df_bi[bi_group_col] = df_bi[bi_group_col].astype(str).str.strip()
        df_bi[bi_service_col] = df_bi[bi_service_col].astype(str).str.strip()
        df_bi[bi_mtd_col] = pd.to_numeric(df_bi[bi_mtd_col], errors='coerce').fillna(0)
        df_bi[bi_py_col] = pd.to_numeric(df_bi[bi_py_col], errors='coerce').fillna(0)
        
        # --- A. Get BI DPC-level maps ---
        bi_dpc_sums_mtd = df_bi.groupby(bi_group_col)[bi_mtd_col].sum()
        bi_dpc_sums_py = df_bi.groupby(bi_group_col)[bi_py_col].sum()
        bi_map_mtd = bi_dpc_sums_mtd.to_dict()
        bi_map_py = bi_dpc_sums_py.to_dict()

        # --- B. Get BI SERVICE total ---
        bi_service_total_mtd = df_bi[df_bi[bi_service_col] == 'SERVICE'][bi_mtd_col].sum()
        bi_service_total_py = df_bi[df_bi[bi_service_col] == 'SERVICE'][bi_py_col].sum()

        print(f"     ... Built BI (Cleaned File) map with {len(bi_map_mtd)} DPC keys.")
        print(f"     ... BI SERVICE Total (MTD: {bi_service_total_mtd:.2f}, PY: {bi_service_total_py:.2f})")

        # --- 3. Load Hyperion Validation File (Truth Data) ---
        
        # Use glob to find the file in the folder
        hyperion_files = glob.glob(os.path.join(hyperion_oe_folder_path, "*.xlsx"))
        if not hyperion_files:
            print(f"   - SKIPPING validation: No Hyperion Excel file found in '{hyperion_oe_folder_path}'.")
            return all_validation_sheets
        
        hyperion_oe_file_path = hyperion_files[0]
            
        print(f"   - Using Hyperion validation file: {os.path.basename(hyperion_oe_file_path)}")
        
        try:
            all_sheets = pd.read_excel(hyperion_oe_file_path, sheet_name=None, header=None)
            df_hyperion_sheet = all_sheets.get(month_abbr)
            
            if df_hyperion_sheet is None and month_abbr != "Sheet1":
                 print(f"   - WARNING: Could not find sheet '{month_abbr}'. Trying 'Sheet1'...")
                 df_hyperion_sheet = all_sheets.get('Sheet1')

        except Exception as e:
            print(f"   - ERROR: Could not read Hyperion workbook. {e}")
            return all_validation_sheets
            
        if df_hyperion_sheet is None:
            print(f"   - ERROR: Could not read sheet '{month_abbr}' or 'Sheet1'.")
            return all_validation_sheets

        # --- A. Extract Hyperion data using the helper ---
        (hyperion_dpc_map_mtd, 
         hyperion_dpc_map_py, 
         last_row_mtd, 
         last_row_py) = _extract_dpc_maps_from_sheet(df_hyperion_sheet, pc_str) # <-- pc_str is now the correct OE_PC

        if not hyperion_dpc_map_mtd and last_row_mtd == 0:
            print(f"     - WARNING: Profit Center {pc_str} not found or has no data in Hyperion sheet. Skipping.")
            return all_validation_sheets
            
        print(f"     ... Built Hyperion map for PC {pc_str} with {len(hyperion_dpc_map_mtd)} DPC keys.")
        print(f"     ... Hyperion SERVICE Total (MTD: {last_row_mtd}, PY: {last_row_py})")
            
        # --- 4. Compare the two maps ---
        validation_data = []
        
        # --- A. Compare DPC keys ---
        all_dpc_keys_bi = set(bi_map_mtd.keys()) | set(bi_map_py.keys())
        all_dpc_keys_hyperion = set(hyperion_dpc_map_mtd.keys()) | set(hyperion_dpc_map_py.keys())
        
        mapped_hyperion_keys = {HYPERION_TO_BI_DPC_MAP.get(k, k) for k in all_dpc_keys_hyperion}
        
        all_dpc_keys = sorted(list(all_dpc_keys_bi | mapped_hyperion_keys))
        
        print(f"     - Comparing {len(all_dpc_keys)} unique 'DPC' items for PC {pc_str}.")
        
        for dpc_key in all_dpc_keys:
            if pd.isna(dpc_key) or dpc_key == 'nan': continue
            
            # --- [FINAL CHANGE IS HERE] ---
            # Skip the 'Adjustment figure' row
            if dpc_key == 'Adjustment figure':
                print(f"     - Skipping 'Adjustment figure' row.")
                continue
            # --- [END FINAL CHANGE] ---
            
            # 1. Get BI Values
            bi_val_mtd = bi_map_mtd.get(dpc_key, 0)
            bi_val_py = bi_map_py.get(dpc_key, 0)
            
            # 2. Get Hyperion Values
            hyperion_key = dpc_key
            if dpc_key in HYPERION_TO_BI_DPC_MAP.values():
                hyperion_key = next((k for k, v in HYPERION_TO_BI_DPC_MAP.items() if v == dpc_key), dpc_key)
            
            hyp_val_mtd = float(hyperion_dpc_map_mtd.get(hyperion_key, 0)) * 1000
            hyp_val_py = float(hyperion_dpc_map_py.get(hyperion_key, 0)) * 1000
            
            # 3. Compare
            diff_mtd = hyp_val_mtd - bi_val_mtd
            diff_py = hyp_val_py - bi_val_py
            variance = 5.0
            status_mtd = "Matching" if abs(diff_mtd) <= variance else "Not Matching"
            status_py = "Matching" if abs(diff_py) <= variance else "Not Matching"
            
            validation_data.append({
                'Group': dpc_key,
                'BI MTD': bi_val_mtd,
                'BI PY': bi_val_py,
                'Hyperion MTD': hyp_val_mtd,
                'Hyperion PY': hyp_val_py,
                'Difference MTD': diff_mtd,
                'Difference PY': diff_py,
                'Status MTD': status_mtd,
                'Status PY': status_py
            })

        # --- B. Compare SERVICE Total ---
        print(f"     - Comparing 'SERVICE (Total)' item.")
        
        hyp_val_mtd_svc = float(last_row_mtd) * 1000
        hyp_val_py_svc = float(last_row_py) * 1000

        diff_mtd_svc = hyp_val_mtd_svc - bi_service_total_mtd
        diff_py_svc = hyp_val_py_svc - bi_service_total_py
        
        status_mtd_svc = "Matching" if abs(diff_mtd_svc) <= variance else "Not Matching"
        status_py_svc = "Matching" if abs(diff_py_svc) <= variance else "Not Matching"

        validation_data.append({
            'Group': 'SERVICE (Total)',
            'BI MTD': bi_service_total_mtd,
            'BI PY': bi_service_total_py,
            'Hyperion MTD': hyp_val_mtd_svc,
            'Hyperion PY': hyp_val_py_svc,
            'Difference MTD': diff_mtd_svc,
            'Difference PY': diff_py_svc,
            'Status MTD': status_mtd_svc,
            'Status PY': status_py_svc
        })

        # --- 5. Return results ---
        df_validation = pd.DataFrame(validation_data)
        
        all_validation_sheets[pc_str] = df_validation
        print(f"     - Successfully generated validation data for sheet: {pc_str}.")
        
        print(f"   - Finished generating data for all sheets.")
        return all_validation_sheets

    except Exception as e:
        print(f"   - ❌ ERROR during OE Hyperion data generation: {e}")
        import traceback
        traceback.print_exc()
        return all_validation_sheets
    
# ==============================================================================
# ---                   PEX-BI Hyperion Validation Function                  ---
# ==============================================================================
def generate_pex_validation_data(processed_file_path, hyperion_file_path, directory_file_path):
    """
    Validates a processed PEX-BI file against a PEX Hyperion validation template.
    
    (v3 Update):
    - Dynamically finds the sheet name (e.g., 'Sep') from the filename (e.g., ..._0925.xlsx).
    - Falls back to 'Sheet1' if the month-specific sheet is not found.
    
    (v2 Update):
    - Uses a helper to get the correct PC from the filename.
    - Works for both single-entity files (e.g., ..._2033_...)
    - Works for grouped files (e.g., ..._Nordic_...) by finding the group's
      summary PC (e.g., '9001') from the directory file.
    - Validates the *entire file* against that single PC.
    
    Returns a dictionary of {profit_center_str: validation_dataframe}
    """
    print(f"--- Generating PEX Hyperion Validation Data for {os.path.basename(processed_file_path)} ---")
    
    all_validation_sheets = {}

    # --- [NEW] Add month map ---
    month_num_to_abbr = {
        '01': 'Jan', '02': 'Feb', '03': 'Mar', '04': 'Apr', '05': 'May', '06': 'Jun',
        '07': 'Jul', '08': 'Aug', '09': 'Sep', '10': 'Oct', '11': 'Nov', '12': 'Dec'
    }

    try:
        # --- 1. Get the correct Profit Center to check ---
        filename = os.path.basename(processed_file_path)
        
        # --- A. Load the Group-to-PC map ---
        group_to_pc_map = _load_group_to_pc_map(directory_file_path, 'Grouping Unit', 'Comp_No_for_OE')

        # --- B. Define regex patterns ---
        # Pattern for single files: ..._DK01_2033_0925.xlsx -> Captures "2033"
        single_regex = r'_[A-Z0-9]+_(\d+)_\d{4}\.xlsx'
        # Pattern for grouped files: ..._Processed_Nordic_0925.xlsx -> Captures "Nordic"
        grouped_regex = r'_Processed_([A-Za-z\s()-]+)_\d{4}\.xlsx'

        # --- C. Get the PC to check ---
        (pc_str, is_grouped) = _get_pc_and_group_from_filename(filename, group_to_pc_map, single_regex, grouped_regex)

        if not pc_str:
            return all_validation_sheets # Helper function already printed the error

        # --- [NEW] 1.D. Get Month/Year from filename ---
        date_part = None
        # Regex to find the MMYY part, e.g., "_0925.xlsx"
        match = re.search(r'_(\d{4})\.xlsx', filename, re.IGNORECASE)
        if match:
            date_part = match.group(1)
        
        if not date_part:
            print(f"   - SKIPPING: Could not determine date_part (e.g., '0925') from filename.")
            return all_validation_sheets

        month_mm = date_part[0:2]
        month_abbr = month_num_to_abbr.get(month_mm, "Sheet1") # e.g., "Sep"
        print(f"   - Target Sheet: {month_abbr}")

        # --- 2. Load Processed BI File ---
        df_bi = pd.read_excel(processed_file_path, sheet_name='Sheet1')
        
        if 'Group' not in df_bi.columns:
            print(f"   - SKIPPING validation: BI file '{filename}' is missing 'Group' column.")
            return all_validation_sheets
        
        # (Rest of BI column checks...)
        if len(df_bi.columns) < 7:
            print(f"   - SKIPPING validation: BI file '{filename}' has fewer than 7 columns.")
            return all_validation_sheets
            
        actual_col_bi = df_bi.columns[5] # e.g., 'Actual SEP 2025'
        prior_col_bi = df_bi.columns[6]  # e.g., 'Prior Yr SEP 2024'
        
        year_actual_match = re.search(r'(\d{4})$', actual_col_bi)
        year_prior_match = re.search(r'(\d{4})$', prior_col_bi)
        year_actual_str = year_actual_match.group(1) if year_actual_match else "YYYY"
        year_prior_str = year_prior_match.group(1) if year_prior_match else "YYYY-1"
        
        print(f"   - Validating BI columns: '{actual_col_bi}' and '{prior_col_bi}'")

        # --- [MODIFIED] 3. Load Hyperion Validation File (Dynamic Sheet) ---
        if not os.path.exists(hyperion_file_path):
            print(f"   - SKIPPING validation: Hyperion file not found: {os.path.basename(hyperion_file_path)}")
            return all_validation_sheets
            
        print(f"   - Using Hyperion validation file: {os.path.basename(hyperion_file_path)}")
        
        try:
            # Try to read the dynamic month sheet
            df_hyperion = pd.read_excel(hyperion_file_path, sheet_name=month_abbr, header=None)
        except Exception as e:
            if month_abbr != "Sheet1":
                # If it failed, and it wasn't 'Sheet1' already, try 'Sheet1'
                print(f"   - WARNING: Could not find sheet '{month_abbr}'. Trying 'Sheet1'... Error: {e}")
                try:
                    df_hyperion = pd.read_excel(hyperion_file_path, sheet_name='Sheet1', header=None)
                except Exception as e2:
                    print(f"   - ERROR: Could not read sheet '{month_abbr}' or 'Sheet1'. {e2}")
                    return all_validation_sheets
            else:
                # It failed even on 'Sheet1'
                print(f"   - ERROR: Could not read sheet 'Sheet1'. {e}")
                return all_validation_sheets

        # Now that df_hyperion is loaded, get the data parts
        pc_row_hyperion = df_hyperion.iloc[10] # 11th row (for PC lookup)
        df_hyperion_data_part = df_hyperion.iloc[12:] # Start from 13th row

        # --- 4. Process the single Profit Center ---
        # (This block REPLACES the old 'for pc in profit_centers:' loop)
        print(f"   - Processing Profit Center: {pc_str}")
        
        # --- A. Find matching columns in Hyperion for this PC ---
        target_col_idx_hyperion = None
        for idx, val in pc_row_hyperion.items():
            if isinstance(val, str) and val.strip().endswith(f'.{pc_str}'):
                target_col_idx_hyperion = idx
                break
                
        if target_col_idx_hyperion is None:
            print(f"   - WARNING: Profit Center {pc_str} not found in Hyperion file's 11th row. Skipping.")
            return all_validation_sheets
            
        prior_col_idx_hyperion = target_col_idx_hyperion + 1
        
        # --- B. Get Hyperion data for these columns (Hyperion Map) ---
        account_names_hyperion = df_hyperion_data_part.iloc[:, 3] # Col D
        actual_vals_hyperion = pd.to_numeric(df_hyperion_data_part.iloc[:, target_col_idx_hyperion], errors='coerce').fillna(0)
        prior_vals_hyperion = pd.to_numeric(df_hyperion_data_part.iloc[:, prior_col_idx_hyperion], errors='coerce').fillna(0)
        
        df_hyperion_lookup = pd.DataFrame({
            'Group': account_names_hyperion.astype(str), 
            'Actual': actual_vals_hyperion, 
            'Prior': prior_vals_hyperion
        })
        
        df_hyperion_agg = df_hyperion_lookup.groupby('Group')[['Actual', 'Prior']].sum()
        hyperion_map = df_hyperion_agg.apply(lambda row: (row['Actual'], row['Prior']), axis=1).to_dict()
        print(f"   - Built Hyperion map, aggregated {len(df_hyperion_lookup)} rows into {len(hyperion_map)} unique groups.")

        # --- C. Get BI data for this PC (BI Map) ---
        df_bi_pc = None
        if is_grouped:
            # For a grouped file, we validate the *entire* DataFrame
            df_bi_pc = df_bi.copy()
            print(f"   - Using entire BI file (grouped file detected).")
        else:
            # For a single file, we filter by the 'Profit Center' column
            if 'Profit Center' not in df_bi.columns:
                print(f"   - SKIPPING validation: BI file '{filename}' is missing 'Profit Center' column needed for single-file validation.")
                return all_validation_sheets
            df_bi_pc = df_bi[df_bi['Profit Center'].astype(str).str.strip() == pc_str]
            print(f"   - Filtering BI file for Profit Center == {pc_str}.")
            
        df_bi_grouped = df_bi_pc.groupby('Group')[[actual_col_bi, prior_col_bi]].sum().reset_index()
        df_bi_grouped['Group'] = df_bi_grouped['Group'].astype(str)

        bi_total_actual_sum = df_bi_grouped[actual_col_bi].sum()
        bi_total_prior_sum = df_bi_grouped[prior_col_bi].sum()

        # --- D. Compare and build validation sheet ---
        validation_data = []
        all_groups_bi = set(df_bi_grouped['Group'].dropna().unique())
        all_groups_hyperion = set(hyperion_map.keys())
        all_groups = sorted(list(all_groups_bi | all_groups_hyperion))
        
        print(f"   - Comparing {len(all_groups)} unique 'Group' items for PC {pc_str}.")
        
        for group in all_groups:
            if pd.isna(group) or group == 'nan': continue
            
            if group == "Total Period Expense":
                bi_actual = bi_total_actual_sum
                bi_prior = bi_total_prior_sum
            else:
                bi_row = df_bi_grouped[df_bi_grouped['Group'] == group] 
                bi_actual = bi_row.iloc[0][actual_col_bi] if not bi_row.empty else 0
                bi_prior = bi_row.iloc[0][prior_col_bi] if not bi_row.empty else 0
            
            hyperion_vals = hyperion_map.get(group, (0, 0)) 
            hyperion_actual = float(hyperion_vals[0]) * 1000
            hyperion_prior = float(hyperion_vals[1]) * 1000
            
            difference_actual = hyperion_actual - bi_actual
            difference_prior = hyperion_prior - bi_prior
            variance = 5.0
            status_actual = "Matching" if abs(difference_actual) <= variance else "Not Matching"
            status_prior = "Matching" if abs(difference_prior) <= variance else "Not Matching"
            
            validation_data.append({
                'Group': group,
                f'BI {actual_col_bi}': bi_actual,
                f'BI {prior_col_bi}': bi_prior,
                f'Hyperion {actual_col_bi}': hyperion_actual,
                f'Hyperion {prior_col_bi}': hyperion_prior,
                f'Difference {year_actual_str}': difference_actual,
                f'Difference {year_prior_str}': difference_prior,
                f'Status {year_actual_str}': status_actual,
                f'Status {year_prior_str}': status_prior
            })
        
        df_validation = pd.DataFrame(validation_data)
        all_validation_sheets[pc_str] = df_validation
        print(f"   - Successfully generated validation data for {pc_str}.")
            
        print(f"   - Finished generating data for all sheets.")
        return all_validation_sheets

    except Exception as e:
        print(f"   - ❌ ERROR during PEX Hyperion data generation: {e}")
        import traceback
        traceback.print_exc()
        # Return empty dict on failure
        return all_validation_sheets
        
# ==============================================================================
# --- SECTION 0: Hyperion Adjustment Logic (Moved from clean_oe.py) ---
# ==============================================================================
DPC_TO_DIVISION_MAP = {
        'Labtec': 'Lab', 'ANA': 'Lab', 'Ohaus': 'Lab', 'Pipettes': 'Lab', 'Biotix':'Lab', 'AutoChem':'Lab',
        'OEM': 'Industrial', 'Standard Industrial': 'Industrial', 'T&L': 'Industrial', 'Vehicle': 'Industrial', 'AS':'Industrial',
        'Miscellaneous': 'Misc', 'PI': 'PI', 'Retail': 'Retail', 'PRO': 'PRO',
        'Adjustment figure': 'SERVICE'  
    }

HYPERION_TO_BI_DPC_MAP = {
        'OEM_SI': 'OEM', 'SI_S': 'Standard Industrial', 'TL': 'T&L', 'Misc':'Miscellaneous', 'Pro':'PRO', 'AC':'AutoChem', 'AS_S':'AS'
    }

def _extract_dpc_maps_from_sheet(df_sheet, profit_center):
    """
    Helper function to extract the DPC-to-value mapping from a single Hyperion sheet
    for both the current year (MTD) and prior year (PY), AND extract the final total row.
    """
    # Initialize returns for failure cases
    mtd_map, py_map = {}, {}
    last_row_mtd, last_row_py = 0, 0

    # Basic validation: Incorporating your change to check for at least 27 rows.
    if df_sheet is None or len(df_sheet) < 27:
        print(f"         - Sheet data is invalid or has fewer than 27 rows. Cannot extract DPC map.")
        return mtd_map, py_map, last_row_mtd, last_row_py

    header_row_7 = df_sheet.iloc[6]
    target_col_idx = None
    
    for idx, val in enumerate(header_row_7):
        if idx < 4:
            continue
        if isinstance(val, str) and val.strip().endswith(f'.{profit_center}'):
            target_col_idx = idx
            break
            
    if target_col_idx is None:
        print(f"         - Profit Center '{profit_center}' not found in this sheet's 7th row.")
        return mtd_map, py_map, last_row_mtd, last_row_py

    prior_year_col_idx = target_col_idx + 1
    
    if len(df_sheet) <= 12: # Check if there are any data rows at all
        print(f"         - No data rows found after the header. Cannot extract DPC map.")
        return mtd_map, py_map, last_row_mtd, last_row_py
        
    hyperion_data_full = df_sheet.iloc[11:] # Slices from index 11 (row 12) to the end

    if hyperion_data_full.empty:
        return mtd_map, py_map, last_row_mtd, last_row_py

    hyperion_dpc_data = hyperion_data_full.iloc[:-1]
    last_row_data = hyperion_data_full.iloc[-1]

    if not hyperion_dpc_data.empty:
        dpc_column = hyperion_dpc_data.iloc[:, 1]
        mtd_values = pd.to_numeric(hyperion_dpc_data.iloc[:, target_col_idx], errors='coerce').fillna(0)
        mtd_map = pd.Series(mtd_values.values, index=dpc_column).to_dict()
        
        if prior_year_col_idx < len(df_sheet.columns):
            py_values = pd.to_numeric(hyperion_dpc_data.iloc[:, prior_year_col_idx], errors='coerce').fillna(0)
            py_map = pd.Series(py_values.values, index=dpc_column).to_dict()
    
    mtd_val = pd.to_numeric(last_row_data.iloc[target_col_idx], errors='coerce')
    last_row_mtd = 0 if pd.isna(mtd_val) else mtd_val

    if prior_year_col_idx < len(df_sheet.columns):
        py_val = pd.to_numeric(last_row_data.iloc[prior_year_col_idx], errors='coerce')
        last_row_py = 0 if pd.isna(py_val) else py_val
    
    return mtd_map, py_map, last_row_mtd, last_row_py


def add_hyperion_adjustments(bi_df, profit_center, month_abbr, year_str, hyperion_folder_path, p2_dpc_col_name):
    """
    Compares the BI data against a Hyperion file for both MTD and PY MTD, 
    calculates differences, and adds them back as adjustment lines.
    """
    BOOKINGS_MTD_COL = 'Bookings MTD Net Sales' 
    BOOKINGS_PY_COL = 'Bookings PY MTD'
    ADJUSTMENT_DESC_COL = 'Sales doc. type'

    required_cols = [p2_dpc_col_name, BOOKINGS_MTD_COL, BOOKINGS_PY_COL]
    if not all(col in bi_df.columns for col in required_cols):
        print(f"         ⚠️  Skipping Hyperion adjustment: BI data is missing one or more required columns.")
        return bi_df

    # --- MODIFIED: Use passed-in arguments ---
    sheet_name = month_abbr
    print(f"   -> Starting Hyperion check for Profit Center: '{profit_center}'")
    print(f"         - Target Hyperion Sheet: '{sheet_name}' (Year: {year_str})")
    # --- END MODIFICATION ---

    hyperion_files = glob.glob(os.path.join(hyperion_folder_path, "*.xlsx"))
    if not hyperion_files:
        print(f"         ⚠️  Skipping Hyperion adjustment: No Hyperion Excel file found in '{hyperion_folder_path}'.")
        return bi_df
    
    try:
        all_sheets = pd.read_excel(hyperion_files[0], sheet_name=None, header=None)
        df_hyperion_sheet = all_sheets.get(sheet_name)
    except Exception as e:
        print(f"         ⚠️  Skipping Hyperion adjustment: Could not read workbook. Error: {e}.")
        return bi_df

    if df_hyperion_sheet is None:
        print(f"         ⚠️  Skipping Hyperion adjustment: Sheet '{sheet_name}' not found in Hyperion file.")
        return bi_df
        
    hyperion_dpc_map_mtd, hyperion_dpc_map_py, last_row_mtd, last_row_py = _extract_dpc_maps_from_sheet(df_hyperion_sheet, profit_center)
    
    if not hyperion_dpc_map_mtd:
        print(f"         ⚠️  Warning: No DPC data extracted from sheet '{sheet_name}'. Proceeding to check for SERVICE total.")

    bi_df[BOOKINGS_MTD_COL] = pd.to_numeric(bi_df[BOOKINGS_MTD_COL], errors='coerce').fillna(0)
    bi_df[BOOKINGS_PY_COL] = pd.to_numeric(bi_df[BOOKINGS_PY_COL], errors='coerce').fillna(0)
    bi_dpc_sums_mtd = bi_df.groupby(p2_dpc_col_name)[BOOKINGS_MTD_COL].sum()
    bi_dpc_sums_py = bi_df.groupby(p2_dpc_col_name)[BOOKINGS_PY_COL].sum()

    
    adjustments_to_add = []
    
    for hyperion_dpc_str, hyperion_value_mtd in hyperion_dpc_map_mtd.items():
        if pd.isna(hyperion_dpc_str) or "Total" in hyperion_dpc_str: continue
        
        bi_dpc_str = HYPERION_TO_BI_DPC_MAP.get(hyperion_dpc_str, hyperion_dpc_str)
        
        bi_value_mtd = bi_dpc_sums_mtd.get(bi_dpc_str, 0)
        bi_value_py = bi_dpc_sums_py.get(bi_dpc_str, 0)

        hyperion_value_mtd_scaled = float(hyperion_value_mtd) * 1000
        difference_mtd = hyperion_value_mtd_scaled - float(bi_value_mtd)
        
        hyperion_value_py = hyperion_dpc_map_py.get(hyperion_dpc_str, 0)
        hyperion_value_py_scaled = float(hyperion_value_py) * 1000
        difference_py = hyperion_value_py_scaled - float(bi_value_py)
        
        if abs(difference_mtd) > 0.001 or abs(difference_py) > 0.001:
            print(f"           -> Difference found for DPC '{bi_dpc_str}' (from Hyperion's '{hyperion_dpc_str}'):")
            print(f"                 - MTD: Hyperion={hyperion_value_mtd_scaled:.2f}, BI={bi_value_mtd:.2f}, Adjustment={difference_mtd:.2f}")
            print(f"                 - PY:  Hyperion={hyperion_value_py_scaled:.2f}, BI={bi_value_py:.2f}, Adjustment={difference_py:.2f}")

            new_row = {col: 'Adjustment figure' for col in bi_df.columns}
            new_row[p2_dpc_col_name] = bi_dpc_str
            new_row[BOOKINGS_MTD_COL] = difference_mtd
            new_row[BOOKINGS_PY_COL] = difference_py
            if 'Product/Service' in new_row:
                new_row['Product/Service'] = 'PRODUCT'
            
            if 'P1-Division' in new_row:
                new_row['P1-Division'] = DPC_TO_DIVISION_MAP.get(bi_dpc_str, 'Adjustment figure')

            adjustments_to_add.append(new_row)

    service_mtd_val = float(last_row_mtd) * 1000
    service_py_val = float(last_row_py) * 1000

    if abs(service_mtd_val) > 0.001 or abs(service_py_val) > 0.001:
        print(f"           -> Adding 'SERVICE' adjustment row from Hyperion total (MTD: {service_mtd_val:.2f}, PY: {service_py_val:.2f})")
        service_row = {col: 'Adjustment figure' for col in bi_df.columns}
        
        if 'Product/Service' in service_row:
            service_row['Product/Service'] = 'SERVICE'
            
        if 'P1-Division' in service_row:
            service_row['P1-Division'] = DPC_TO_DIVISION_MAP.get(service_row['P1-Division'], 'SERVICE')
            
        service_row[BOOKINGS_MTD_COL] = service_mtd_val
        service_row[BOOKINGS_PY_COL] = service_py_val
        adjustments_to_add.append(service_row)

    if adjustments_to_add:
        print("           -> Adding adjustment rows to the BI data.")
        df_adjustments = pd.DataFrame(adjustments_to_add)
        bi_df = pd.concat([bi_df, df_adjustments], ignore_index=True)
    else:
        print("   -> No differences found between Hyperion and BI data, and no SERVICE total to add.")
        
    return bi_df

# ==============================================================================
# --- SECTION 1: Grouping Logic ---
# ==============================================================================
def _load_comp_no_to_oe_map(directory_file_path):
    """
    Loads the Comp_No to Comp_No_for_OE map from the directory file.
    Returns a dictionary: {'5231': '9005', '5232': '9005', ...}
    """
    print("Loading Comp_No to Comp_No_for_OE map...")
    try:
        df_dir = pd.read_excel(directory_file_path)
        
        # Ensure we have the columns we need
        if 'Comp_No' not in df_dir.columns or 'Comp_No_for_OE' not in df_dir.columns:
            print(f"     - ❌ ERROR: Directory missing required columns: 'Comp_No' or 'Comp_No_for_OE'.")
            return {}
            
        df_dir = df_dir.dropna(subset=['Comp_No', 'Comp_No_for_OE'])
        
        # Convert to string for reliable matching
        df_dir['Comp_No'] = df_dir['Comp_No'].astype(str).str.strip()
        df_dir['Comp_No_for_OE'] = df_dir['Comp_No_for_OE'].astype(str).str.strip()
        
        # Drop duplicates based on Comp_No
        df_dir = df_dir.drop_duplicates(subset=['Comp_No'])
        
        oe_map = pd.Series(df_dir['Comp_No_for_OE'].values, index=df_dir['Comp_No']).to_dict()
        print(f"Loaded {len(oe_map)} Comp_No -> OE Comp_No entries.")
        return oe_map
        
    except Exception as e:
        print(f"❌ Error loading Comp_No -> OE map: {e}")
        return {}

def _load_grouping_map(directory_file_path):
    """
    Loads the Comp_No to Grouping Unit map from the directory file.
    Returns a dictionary: {'2031': 'Nordic', '2032': 'Nordic', ...}
    """
    print("Loading grouping map...")
    try:
        df_dir = pd.read_excel(directory_file_path)
        # Drop rows where Grouping Unit is null
        df_dir = df_dir.dropna(subset=['Grouping Unit'])
        # Convert Comp_No to string for reliable matching
        df_dir['Comp_No'] = df_dir['Comp_No'].astype(str)
        # Create the map
        group_map = pd.Series(df_dir['Grouping Unit'].values, index=df_dir['Comp_No']).to_dict()
        print(f"Loaded {len(group_map)} grouping entries.")
        return group_map
    except Exception as e:
        print(f"❌ Error loading grouping map: {e}")
        return {}

def _load_grouping_map_oe(directory_file_path):
    """
    Loads the Comp_No_for_OE to Grouping Unit map from the directory file.
    Returns a dictionary: {'2031': 'Nordic', '2032': 'Nordic', ...}
    """
    print("Loading grouping map...")
    try:
        df_dir = pd.read_excel(directory_file_path)
        # Drop rows where Grouping Unit is null
        df_dir = df_dir.dropna(subset=['Grouping Unit'])
        # Convert Comp_No_for_OE to string for reliable matching
        df_dir['Comp_No_for_OE'] = df_dir['Comp_No_for_OE'].astype(str)
        # Create the map
        group_map = pd.Series(df_dir['Grouping Unit'].values, index=df_dir['Comp_No_for_OE']).to_dict()
        print(f"Loaded {len(group_map)} grouping entries.")
        return group_map
    except Exception as e:
        print(f"❌ Error loading grouping map: {e}")
        return {}


def _parse_sales_filename(filename):
    """
    Parses the processed sales filename to get Comp_No and other parts.
    Format: Sales_Data_Processed_UNIT_COMPNO_MMYY_TYPE.csv
    """
    # Format: Sales_Data_Processed_UNIT_COMPNO_MMYY_TYPE.csv
    match = re.search(r'Sales_Data_Processed_([A-Z0-9]+)_(\d+)_(\d{4})_([A-Z0-9]+)\.csv', filename)
    if match:
        return {
            "comp_no": match.group(2),
            "unit": match.group(1),
            "date_part": match.group(3), # "MMYY"
            "type": match.group(4)      # "3RD" or "IC"
        }
    
    # Fallback Format: Sales_Data_Processed_UNIT_COMPNO_MMYY.csv (if no TYPE)
    match_no_type = re.search(r'Sales_Data_Processed_([A-Z0-9]+)_(\d+)_(\d{4})\.csv', filename)
    if match_no_type:
         return {
            "comp_no": match_no_type.group(2),
            "unit": match_no_type.group(1),
            "date_part": match_no_type.group(3),
            "type": None
        }

    return None # Not a file we can parse

# --- NEW PARSING FUNCTIONS ---

def _parse_oe_filename(filename):
    """
    Parses: OE_Data_Processed_DK01_2033_0925.csv
    """
    # --- UPDATED REGEX ---
    match = re.search(r'OE_Data_Processed_([A-Z0-9]+)_(\d+)_(\d{4})(\(\d+\))?\.csv', filename, re.IGNORECASE)
    # --- END UPDATE ---
    
    if match:
        return {
            "unit": match.group(1),
            "comp_no": match.group(2),
            "date_part": match.group(3) # "MMYY"
        }
    return None

def _parse_pex_bi_filename(filename):
    """
    Parses: PEX_Data_Processed_DK01_2033_0925.xlsx
    """
    match = re.search(r'PEX_Data_Processed_([A-Z0-9]+)_(\d+)_(\d{4})\.xlsx', filename, re.IGNORECASE)
    if match:
        return {
            "unit": match.group(1),
            "comp_no": match.group(2),
            "date_part": match.group(3) # "MMYY"
        }
    return None

def _parse_headcount_filename(filename):
    """
    Parses: DK01_0925_Headcount_Processed_2033.xlsx
    """
    # Catches Unit, MMYY, and CompNo
    match = re.search(r'([A-Z0-9]+)_(\d{4})_Headcount_Processed_(\d+)\.xlsx', filename, re.IGNORECASE)
    if match:
        return {
            "unit": match.group(1),
            "date_part": match.group(2), # "MMYY"
            "comp_no": match.group(3)
        }
    return None

def _parse_pex_vendor_filename(filename):
    """
    Parses: DK01_2033_vendor_analysis_combined.xlsx
    """
    match = re.search(r'([A-Z0-9]+)_(\d+)_vendor_analysis_combined\.xlsx', filename, re.IGNORECASE)
    if match:
        return {
            "unit": match.group(1),
            "comp_no": match.group(2)
        }
    return None

# --- END NEW PARSING FUNCTIONS ---
def group_sales_files(output_folder, processed_files, directory_file_path):
    """
    Post-processes files in the output folder to group them by 'Grouping Unit'.
    Merges Sales (CSV) files based on group, date, and (if available) type.
    """
    print("--- Starting Sales file grouping post-process ---")
    group_map = _load_grouping_map(directory_file_path)
    if not group_map:
        print("No grouping map found. Aborting grouping.")
        return processed_files

    files_to_group = {}
    final_file_list = [] 

    for filename in processed_files:
        file_info = _parse_sales_filename(filename)
        
        if not file_info:
            print(f"   - Could not parse '{filename}', keeping as is.")
            final_file_list.append(filename)
            continue
        
        comp_no = file_info.get("comp_no")
        group_name = group_map.get(comp_no) if comp_no else None
        
        # --- MODIFIED LOGIC ---
        if group_name:
            file_type = file_info.get("type")
            
            if file_type:
                # SCENARIO 1: Group by Name, Date, and Type
                group_key = (group_name, file_info["date_part"], file_type)
                print(f"   - Queued '{filename}' for grouping under '{group_name}' (Type: {file_type}).")
            else:
                # SCENARIO 2 (FALLBACK): Group by Name and Date only
                group_key = (group_name, file_info["date_part"])
                print(f"   - Queued '{filename}' for fallback grouping under '{group_name}' (No Type).")

            if group_key not in files_to_group:
                files_to_group[group_key] = []
            
            file_path = os.path.join(output_folder, filename)
            files_to_group[group_key].append(file_path)
            
        else:
            # SCENARIO 3: No group found
            print(f"   - No group for Comp_No '{comp_no}' in '{filename}', keeping as is.")
            final_file_list.append(filename)
    
    # --- END MODIFIED LOGIC ---

    # Now, process the grouped files
    for group_key, file_paths in files_to_group.items():
        
        group_desc = "" # For logging
        
        # --- MODIFIED LOGIC ---
        # Check key length to determine filename format
        if len(group_key) == 3:
            # Key is (group_name, date_part, file_type)
            group_name, date_part, file_type = group_key
            new_filename = f"Sales_Data_Processed_{group_name}_{date_part}_{file_type}.csv"
            group_desc = f"'{group_name}' (Type: {file_type})"
            
        elif len(group_key) == 2:
            # Key is (group_name, date_part) - FALLBACK
            group_name, date_part = group_key
            new_filename = f"Sales_Data_Processed_{group_name}_{date_part}.csv"
            group_desc = f"'{group_name}' (Fallback)"
            
        else:
            # Safety check, should not happen
            print(f"     - Skipping malformed group key: {group_key}")
            continue
        # --- END MODIFIED LOGIC ---
            
        new_file_path = os.path.join(output_folder, new_filename)
        
        print(f"   - Merging {len(file_paths)} files into '{new_filename}'")
        
        df_list = []
        for f_path in file_paths:
            try:
                df_list.append(pd.read_csv(f_path))
            except Exception as e:
                print(f"     - Warning: Could not read {os.path.basename(f_path)}. Error: {e}")
        
        if not df_list:
            print(f"     - No files could be read for group {group_desc}. Skipping.")
            continue
            
        try:
            merged_df = pd.concat(df_list, ignore_index=True)
            merged_df.to_csv(new_file_path, index=False, encoding='utf-8-sig')
            
            final_file_list.append(new_filename)
            print(f"   ✅ Successfully created '{new_filename}'")
            
            for f_path in file_paths:
                try:
                    os.remove(f_path)
                except Exception as e:
                    print(f"     - Warning: Could not delete old file {os.path.basename(f_path)}. Error: {e}")
                    
        except Exception as e:
            print(f"     - ❌ Error merging data for group {group_desc}. Error: {e}")
            final_file_list.extend([os.path.basename(p) for p in file_paths])

    print(f"--- Grouping finished. Final files: {final_file_list} ---")
    return final_file_list
# ### MODIFIED FUNCTION ###
def group_oe_files(output_folder, processed_files, directory_file_path, hyperion_folder_path):
    """
    Post-processes files in the output folder to group them by 'Grouping Unit'.
    Merges OE (CSV) files based on group and date, then applies Hyperion adjustments.
    
    (v3 Update):
    - Parses Comp_No from filename (e.g., '5231').
    - Uses a new map to find the *actual* Comp_No_for_OE (e.g., '9005').
    - Uses this Comp_No_for_OE for grouping and for Hyperion adjustments.
    """
    print("--- Starting OE file grouping post-process ---")
    
    # Map 1: Used to find which files to group (Comp_No_for_OE -> Group)
    group_map = _load_grouping_map_oe(directory_file_path)
    if not group_map:
        print("No grouping map (Comp_No_for_OE -> Group) found. Aborting grouping.")
        return processed_files
        
    # Map 2: Used to find summary PC for adjustments (Group -> Summary PC)
    reverse_group_map = _load_group_to_pc_map(directory_file_path, 'Grouping Unit', 'Comp_No_for_OE')
    if not reverse_group_map:
        print("No reverse grouping map (Group -> Summary PC) found. Aborting grouping.")
        return processed_files

    # --- [NEW] Map 3: Used to find the correct OE PC from the filename PC ---
    comp_no_to_oe_map = _load_comp_no_to_oe_map(directory_file_path)
    if not comp_no_to_oe_map:
        print("No Comp_No -> Comp_No_for_OE map found. Aborting grouping.")
        return processed_files
    # --- [END NEW] ---

    files_to_group = {}
    final_file_list = [] 
    
    month_num_to_abbr = {
        '01': 'Jan', '02': 'Feb', '03': 'Mar', '04': 'Apr', '05': 'May', '06': 'Jun',
        '07': 'Jul', '08': 'Aug', '09': 'Sep', '10': 'Oct', '11': 'Nov', '12': 'Dec'
    }

    for filename in processed_files:
        file_info = _parse_oe_filename(filename)
        
        if not file_info:
            print(f"   - Could not parse '{filename}', keeping as is.")
            final_file_list.append(filename)
            continue
        
        # --- [MODIFIED LOGIC] ---
        pc_from_file = file_info.get("comp_no") # e.g., '5231'
        
        # Find the actual OE Profit Center to use for grouping and adjustments
        comp_no_for_oe = comp_no_to_oe_map.get(pc_from_file)
        
        if not comp_no_for_oe:
            print(f"   - ❌ Skipping '{filename}': Comp_No '{pc_from_file}' not found in directory's 'Comp_No' column.")
            final_file_list.append(filename) # Keep it as is
            continue

        # Now, use the correct Comp_No_for_OE to find the group
        group_name = group_map.get(comp_no_for_oe) 
        # --- [END MODIFIED LOGIC] ---
        
        if group_name:
            # --- FILE WILL BE GROUPED ---
            date_part = file_info["date_part"] # e.g., "0925"
            group_key = (group_name, date_part)
            
            if group_key not in files_to_group:
                files_to_group[group_key] = []
            
            file_path = os.path.join(output_folder, filename)
            files_to_group[group_key].append(file_path)
            
            print(f"   - Queued '{filename}' (PC {pc_from_file} -> OE_PC {comp_no_for_oe}) for grouping under '{group_name}'.")
        else:
            # --- FILE IS NOT GROUPED, APPLY ADJUSTMENTS NOW ---
            
            # --- [MODIFIED LOGIC FOR STANDALONE FILES] ---
            print(f"   - No group for Comp_No_for_OE '{comp_no_for_oe}' (from file PC '{pc_from_file}'). Applying adjustments directly.")
            
            file_path = os.path.join(output_folder, filename)
            
            try:
                # --- 1. Get metadata for this standalone file ---
                date_part = file_info["date_part"] # e.g., "0925"
                profit_center_to_use = comp_no_for_oe # <-- Use the mapped OE_PC
                
                month_mm = date_part[0:2]
                year_yy = date_part[2:4]
                month_abbr = month_num_to_abbr[month_mm]
                year_str = f"20{year_yy}" # Assumes 21st century

                # --- 2. Load the file ---
                standalone_df = pd.read_csv(file_path)

                # --- 3. Apply Hyperion Adjustment ---
                print(f"     - Running adjustment for standalone Profit Center: {profit_center_to_use} ({month_abbr} {year_str})")
                standalone_df = add_hyperion_adjustments(
                    standalone_df, 
                    profit_center_to_use, # <-- Use the mapped OE_PC
                    month_abbr, 
                    year_str, 
                    hyperion_folder_path, 
                    p2_dpc_col_name='P2-DPC'
                )

                # --- 4. Save (overwrite) the file ---
                standalone_df.to_csv(file_path, index=False, encoding='utf-8-sig')
                print(f"   ✅ Successfully applied adjustments to standalone file '{filename}'")
                
            except Exception as e:
                print(f"     - ❌ FAILED to apply adjustment to standalone file {filename}. Error: {e}")
                traceback.print_exc()
            
            # --- 5. Add to final list ---
            final_file_list.append(filename)
            # --- [END MODIFIED LOGIC] ---

    # Now, process the grouped files
    for group_key, file_paths in files_to_group.items():
        group_name, date_part = group_key
        
        new_filename = f"OE_Data_Processed_{group_name}_{date_part}.csv"
        new_file_path = os.path.join(output_folder, new_filename)
        
        print(f"   - Merging {len(file_paths)} files into '{new_filename}'")
        
        df_list = []
        for f_path in file_paths:
            try:
                df_list.append(pd.read_csv(f_path))
            except Exception as e:
                print(f"     - Warning: Could not read {os.path.basename(f_path)}. Error: {e}")
        
        if not df_list:
            print(f"     - No files could be read for group '{group_name}'. Skipping.")
            continue
            
        try:
            merged_df = pd.concat(df_list, ignore_index=True)
            
            # --- Apply Hyperion Adjustments AFTER Merging ---
            print(f"   -> Applying Hyperion adjustments to merged file: {new_filename}")
            
            # 1. Find the summary PC for this group
            summary_pc = reverse_group_map.get(group_name)
            
            if not summary_pc:
                print(f"     - ❌ ERROR: Cannot find summary PC for group '{group_name}'. Skipping adjustments.")
            
            else:
                # 2. Get date info
                try:
                    month_mm = date_part[0:2]
                    year_yy = date_part[2:4]
                    month_abbr = month_num_to_abbr[month_mm]
                    year_str = f"20{year_yy}"

                    # 3. Run adjustment ONCE
                    print(f"     - Running adjustment for Group '{group_name}' using Summary PC: {summary_pc} ({month_abbr} {year_str})")
                    merged_df = add_hyperion_adjustments(
                        merged_df, 
                        summary_pc, # <-- Use summary PC
                        month_abbr, 
                        year_str, 
                        hyperion_folder_path, 
                        p2_dpc_col_name='P2-DPC'
                    )
                except Exception as e:
                    print(f"     - ❌ FAILED to apply adjustment for {summary_pc}. Error: {e}")
                    traceback.print_exc()
            
            print(f"   -> Finished adjustments for {new_filename}.")
            # --- END ADJUSTMENT BLOCK ---

            merged_df.to_csv(new_file_path, index=False, encoding='utf-8-sig')
            
            final_file_list.append(new_filename)
            print(f"   ✅ Successfully created '{new_filename}'")
            
            for f_path in file_paths:
                try:
                    os.remove(f_path)
                except Exception as e:
                    print(f"     - Warning: Could not delete old file {os.path.basename(f_path)}. Error: {e}")
                    
        except Exception as e:
            print(f"     - ❌ Error merging/adjusting data for '{group_name}'. Error: {e}")
            import traceback
            traceback.print_exc()
            final_file_list.extend([os.path.basename(p) for p in file_paths])

    print(f"--- Grouping finished. Final files: {final_file_list} ---")
    return final_file_list

def group_pex_bi_and_headcount_files(output_folder, processed_files, directory_file_path):
    """
    Post-processes files in the output folder to group them by 'Grouping Unit'.
    Merges PEX-BI (Excel) files based on group and date.
    
    (v3 Update):
    - For Headcount, selects *one* file, renames it, and deletes other duplicates.
    - Uses Comp_No_for_OE for the grouped Headcount filename.
    """
    print("--- Starting PEX-BI & Headcount file grouping post-process ---")
    
    # Map 1: Comp_No -> Group (for finding the group)
    group_map = _load_grouping_map(directory_file_path)
    if not group_map:
        print("No grouping map (Comp_No -> Group) found. Aborting grouping.")
        return processed_files

    # Map 2: Group -> Summary PC (for the new filename)
    reverse_group_map = _load_group_to_pc_map(directory_file_path, 'Grouping Unit', 'Comp_No_for_OE')
    if not reverse_group_map:
        print("No reverse grouping map (Group -> Summary PC) found. Aborting grouping.")
        return processed_files

    pex_files_to_group = {}
    headcount_files_to_group = {}
    final_file_list = [] 

    for filename in processed_files:
        # Try parsing as PEX-BI first
        file_info_pex = _parse_pex_bi_filename(filename)
        
        if file_info_pex:
            comp_no = file_info_pex.get("comp_no")
            group_name = group_map.get(comp_no) if comp_no else None
            
            if group_name:
                group_key = (group_name, file_info_pex["date_part"])
                if group_key not in pex_files_to_group:
                    pex_files_to_group[group_key] = []
                
                file_path = os.path.join(output_folder, filename)
                pex_files_to_group[group_key].append(file_path)
                print(f"   - Queued PEX-BI file '{filename}' for grouping under '{group_name}'.")
            else:
                print(f"   - No group for PEX-BI Comp_No '{comp_no}' in '{filename}', keeping as is.")
                final_file_list.append(filename)
            
            continue # Go to the next file

        # Try parsing as Headcount
        file_info_hc = _parse_headcount_filename(filename)
        if file_info_hc:
            comp_no = file_info_hc.get("comp_no")
            group_name = group_map.get(comp_no) if comp_no else None
            
            if group_name:
                group_key = (group_name, file_info_hc["date_part"])
                if group_key not in headcount_files_to_group:
                    headcount_files_to_group[group_key] = []
                
                file_path = os.path.join(output_folder, filename)
                headcount_files_to_group[group_key].append(file_path)
                print(f"   - Queued Headcount file '{filename}' for grouping under '{group_name}'.")
            else:
                print(f"   - No group for Headcount Comp_No '{comp_no}' in '{filename}', keeping as is.")
                final_file_list.append(filename)

            continue # Go to the next file

        # If it's neither, just keep it
        if filename not in final_file_list:
            print(f"   - File '{filename}' is not a PEX-BI or Headcount file, keeping as is.")
            final_file_list.append(filename)

    # --- Process PEX-BI Groups (Unchanged) ---
    for group_key, file_paths in pex_files_to_group.items():
        group_name, date_part = group_key
        
        new_filename = f"PEX_Data_Processed_{group_name}_{date_part}.xlsx"
        new_file_path = os.path.join(output_folder, new_filename)
        
        print(f"   - Merging {len(file_paths)} PEX-BI files into '{new_filename}'")
        
        df_list = []
        for f_path in file_paths:
            try:
                df = pd.read_excel(f_path, sheet_name='Sheet1')
                if 'Unnamed: 3' in df.columns:
                    df.rename(columns={'Unnamed: 3': ''}, inplace=True)
                df_list.append(df)
            except Exception as e:
                print(f"     - Warning: Could not read {os.path.basename(f_path)}. Error: {e}")
        
        if not df_list:
            print(f"     - No PEX-BI files could be read for group '{group_name}'. Skipping.")
            continue
            
        try:
            merged_df = pd.concat(df_list, ignore_index=True)
            merged_df.to_excel(new_file_path, index=False, sheet_name='Sheet1', engine='xlsxwriter')
            final_file_list.append(new_filename)
            print(f"   ✅ Successfully created '{new_filename}'")
            
            for f_path in file_paths:
                try: os.remove(f_path)
                except Exception as e: print(f"     - Warning: Could not delete old file {os.path.basename(f_path)}. Error: {e}")
                    
        except Exception as e:
            print(f"     - ❌ Error merging PEX-BI data for '{group_name}'. Error: {e}")
            import traceback
            traceback.print_exc()
            final_file_list.extend([os.path.basename(p) for p in file_paths])

    # --- [MODIFIED] Process Headcount Groups ---
    for group_key, file_paths in headcount_files_to_group.items():
        group_name, date_part = group_key
        
        # Get the summary PC for the filename
        summary_pc = reverse_group_map.get(group_name, "UnknownPC")
        
        # Format: {Group}_{MMYY}_Headcount_Processed_{Summary_PC}.xlsx
        new_filename = f"{group_name}_{date_part}_Headcount_Processed_{summary_pc}.xlsx"
        new_file_path = os.path.join(output_folder, new_filename)
        
        # --- [NEW LOGIC START] ---
        # 1. Check if there are any files at all
        if not file_paths:
            print(f"     - No Headcount files found for group '{group_name}'. Skipping.")
            continue
            
        # 2. Take the first file as the "representative" file
        source_file_path = file_paths[0]
        
        print(f"   - Selecting '{os.path.basename(source_file_path)}' as representative for Headcount group '{group_name}'.")
        print(f"   - Renaming to '{new_filename}'.")
        
        try:
            # 3. Rename the first file to the new grouped name
            os.rename(source_file_path, new_file_path)
            final_file_list.append(new_filename)
            print(f"   ✅ Successfully created '{new_filename}'.")
            
            # 4. Delete all *other* files in the group (from index 1 onwards)
            other_files_to_delete = file_paths[1:]
            if other_files_to_delete:
                print(f"   - Deleting {len(other_files_to_delete)} other duplicate Headcount files...")
                for f_path in other_files_to_delete:
                    try: 
                        os.remove(f_path)
                        print(f"     - Deleted '{os.path.basename(f_path)}'.")
                    except Exception as e: 
                        print(f"     - Warning: Could not delete old file {os.path.basename(f_path)}. Error: {e}")
        
        except Exception as e:
            # This catch block is for the os.rename or the deletion loop
            print(f"     - ❌ Error processing Headcount group '{group_name}'. Error: {e}")
            import traceback
            traceback.print_exc()
            # If it failed, just add back the original files to the list so they aren't lost
            final_file_list.extend([os.path.basename(p) for p in file_paths])
        # --- [NEW LOGIC END] ---

    print(f"--- Grouping finished. Final files: {final_file_list} ---")
    return final_file_list

def group_pex_vendor_files(output_folder, processed_files, directory_file_path):
    """
    Post-processes files in the output folder to group them by 'Grouping Unit'.
    Merges PEX-Vendor (Excel) files based on group.
    """
    print("--- Starting PEX-Vendor file grouping post-process ---")
    group_map = _load_grouping_map(directory_file_path)
    if not group_map:
        print("No grouping map found. Aborting grouping.")
        return processed_files

    files_to_group = {}
    final_file_list = [] 

    for filename in processed_files:
        file_info = _parse_pex_vendor_filename(filename)
        
        if not file_info:
            print(f"   - Could not parse '{filename}', keeping as is.")
            final_file_list.append(filename)
            continue
        
        comp_no = file_info.get("comp_no")
        group_name = group_map.get(comp_no) if comp_no else None
        
        if group_name:
            # No date part, group only by name
            group_key = (group_name,)
            if group_key not in files_to_group:
                files_to_group[group_key] = []
            
            file_path = os.path.join(output_folder, filename)
            files_to_group[group_key].append(file_path)
            print(f"   - Queued '{filename}' for grouping under '{group_name}'.")
        else:
            print(f"   - No group for Comp_No '{comp_no}' in '{filename}', keeping as is.")
            final_file_list.append(filename)

    # Now, process the grouped files
    for group_key, file_paths in files_to_group.items():
        group_name = group_key[0] # Unpack the 1-item tuple
        
        new_filename = f"{group_name}_vendor_analysis_combined.xlsx"
        new_file_path = os.path.join(output_folder, new_filename)
        
        print(f"   - Merging {len(file_paths)} files into '{new_filename}'")
        
        df_list = []
        for f_path in file_paths:
            try:
                # PEX-Vendor files have data on 'Combined_Vendor_Data' (from clean_pex.py)
                df_list.append(pd.read_excel(f_path, sheet_name='Combined_Vendor_Data'))
            except Exception as e:
                print(f"     - Warning: Could not read {os.path.basename(f_path)}. Error: {e}")
        
        if not df_list:
            print(f"     - No files could be read for group '{group_name}'. Skipping.")
            continue
            
        try:
            merged_df = pd.concat(df_list, ignore_index=True)
            # Save as Excel
            merged_df.to_excel(new_file_path, index=False, sheet_name='Combined_Vendor_Data', engine='xlsxwriter')
            
            final_file_list.append(new_filename)
            print(f"   ✅ Successfully created '{new_filename}'")
            
            for f_path in file_paths:
                try:
                    os.remove(f_path)
                except Exception as e:
                    print(f"     - Warning: Could not delete old file {os.path.basename(f_path)}. Error: {e}")
                    
        except Exception as e:
            print(f"     - ❌ Error merging data for '{group_name}'. Error: {e}")
            final_file_list.extend([os.path.basename(p) for p in file_paths])

    print(f"--- Grouping finished. Final files: {final_file_list} ---")
    return final_file_list

# --- NEW FUNCTION FOR DUPLICATE REMOVAL ---

def _get_file_content_hash(file_path):
    """
    Reads a CSV or Excel file and returns a hash of its sorted content.
    """
    try:
        data_string = ""
        if file_path.endswith('.csv'):
            df = pd.read_csv(file_path)
            # Sort columns by name
            df = df.reindex(sorted(df.columns), axis=1)
            # Sort rows by all column values
            df = df.sort_values(by=list(df.columns)).reset_index(drop=True)
            data_string = df.to_string()
            
        elif file_path.endswith('.xlsx'):
            # Read all sheets
            xls = pd.read_excel(file_path, sheet_name=None)
            all_sheets_string = []
            # Process each sheet individually
            for sheet_name in sorted(xls.keys()):
                df = xls[sheet_name]
                # Sort columns by name
                df = df.reindex(sorted(df.columns), axis=1)
                # Sort rows by all column values
                df = df.sort_values(by=list(df.columns)).reset_index(drop=True)
                # Add sheet name and content to representation
                all_sheets_string.append(f"---{sheet_name}---" + df.to_string())
            data_string = "".join(all_sheets_string)
            
        else:
            # Not a file type we can process, return its path as hash to be unique
            return file_path 

        # Return a hash of the canonical string representation
        return hashlib.md5(data_string.encode('utf-8')).hexdigest()

    except Exception as e:
        print(f"     - Warning: Could not read or hash {os.path.basename(file_path)}. Error: {e}")
        # Return a unique hash (the path itself) to prevent it from being marked as a duplicate
        return file_path 

def remove_duplicate_files(output_folder, processed_files):
    """
    Identifies and removes files with duplicate content from the output folder.
    Returns a new list of processed filenames (without the duplicates).
    """
    print("--- Starting duplicate file removal ---")
    content_map = {} # Stores {hash: filename}
    files_to_keep = []
    files_to_remove = []

    for filename in processed_files:
        file_path = os.path.join(output_folder, filename)
        if not os.path.exists(file_path):
            print(f"   - Skipping {filename}: File not found.")
            continue
            
        print(f"   - Analyzing {filename}...")
        file_hash = _get_file_content_hash(file_path)
        
        if file_hash in content_map:
            # This is a duplicate
            original_file = content_map[file_hash]
            print(f"     - Found duplicate: '{filename}' is identical to '{original_file}'.")
            files_to_remove.append(filename)
        else:
            # This is the first time seeing this content
            content_map[file_hash] = filename
            files_to_keep.append(filename)

    if not files_to_remove:
        print("   - No duplicate files found.")
        print("--- Duplicate removal finished ---")
        return processed_files

    print(f"   - Found {len(files_to_remove)} duplicate file(s) to remove.")
    
    # Remove the duplicate files
    for filename in files_to_remove:
        try:
            file_path = os.path.join(output_folder, filename)
            os.remove(file_path)
            print(f"     - Removed '{filename}'.")
        except Exception as e:
            print(f"     -Warning: Could not remove duplicate file {filename}. Error: {e}")
            # If removal fails, it's already excluded from files_to_keep.

    print(f"--- Duplicate removal finished. Kept {len(files_to_keep)} unique files. ---")
    return files_to_keep
