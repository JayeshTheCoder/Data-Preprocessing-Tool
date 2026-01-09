import os
import pandas as pd
import glob
import re  # Added for new filename parsing
import calendar  # Added for currency conversion
from datetime import datetime  # Added for currency conversion

# ==============================================================================
# --- SECTION 0: NEW CURRENCY CONVERSION HELPERS ---
# ==============================================================================

# ### MODIFIED FUNCTION ###
def load_directory_info(directory_file_path):
    """
    Reads the Directory_Processed_Output.xlsx file.
    Returns:
    1. A set of 'Comp_No_for_OE' values where Type is 'MO' or 'MOPO'.
    2. A dictionary mapping (Comp_No_for_OE, SAP_Comp_Code) to its currency info.
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
    
    # Add 'SAP_Comp_Code' to the required columns
    required_cols = ["Comp_No_for_OE", "Type", "SAP_Comp_Code", "Original Currency", "Conversion Currency"]
    if not all(col in df_dir.columns for col in required_cols):
        print(f"❌ ERROR: Directory file must contain {required_cols} columns.")
        return None, None

    # --- 1. Get Comp_No's for 3RD filter logic ---
    comp_numbers_to_process = set() # Initialize as empty set
    try:
        filtered_dir = df_dir[df_dir['Type'].astype(str).str.contains("MO|MOPO", na=False)]
        # Convert all comp numbers to string for reliable matching
        comp_numbers_to_process = set(filtered_dir['Comp_No_for_OE'].astype(str).unique())
        
        if not comp_numbers_to_process:
            print("⚠️ Warning: No 'MO' or 'MOPO' types found in directory file.")
        else:
            print(f"Found {len(comp_numbers_to_process)} relevant Comp_No_for_OE's for (MO/MOPO) logic: {comp_numbers_to_process}")
            
    except Exception as e:
        print(f"❌ ERROR: Failed to filter directory file for MO/MOPO types. Error: {e}")
        return None, None
    
    # --- 2. Build Currency Map ---
    try:
        key_cols = ['Comp_No_for_OE', 'SAP_Comp_Code']
        currency_data_cols = ['Original Currency', 'Conversion Currency']
        
        # We need all columns to be valid
        df_currencies = df_dir[key_cols + currency_data_cols].copy()
        
        # 1. Drop rows where any key or currency info is missing
        df_currencies = df_currencies.dropna(subset=key_cols + currency_data_cols)

        # 2. Convert key columns to string for reliable mapping
        for col in key_cols:
            df_currencies[col] = df_currencies[col].astype(str)

        # 3. Drop fully identical duplicate rows
        df_currencies = df_currencies.drop_duplicates()

        # 4. Check for conflicts:
        duplicates = df_currencies.duplicated(subset=key_cols, keep=False)
        
        if duplicates.any():
            conflicting_data = df_currencies[duplicates].sort_values(by=key_cols)
            print("❌ ERROR: Conflicting currency data found. The same (Comp_No_for_OE, SAP_Comp_Code) pair points to different currencies.")
            print("Please fix these in the Directory_Processed_Output.xlsx file:")
            print(conflicting_data.to_string(index=False))
            return None, None # Stop execution

        # 5. Build the map using the two-column key
        df_currencies = df_currencies.set_index(key_cols)
        currency_map = df_currencies.to_dict('index')
        
        print(f"Loaded currency mapping for {len(currency_map)} (Comp_No, SAP_Code) pairs.")
        
        # Return both the set and the map
        return comp_numbers_to_process, currency_map

    except Exception as e:
        print(f"❌ ERROR: Failed to build currency map. Error: {e}")
        return None, None


def load_currency_rates(currency_file_path, month_int, year_int):
    """
    Loads currency rates from the specified file for the given month/year.
    """
    print(f"    ...loading currency rates for {month_int}/{year_int}")
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

# ==============================================================================
# --- SECTION 1: ORIGINAL SCRIPT (MODIFIED) ---
# ==============================================================================

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

# ### MODIFIED FUNCTION ###
# Simplified to accept parsed filename arguments
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

    dpc_to_division_map = {
        'Labtec': 'Lab', 'ANA': 'Lab', 'Ohaus': 'Lab', 'Pipettes': 'Lab', 'Biotix':'Lab', 'AutoChem':'Lab',
        'OEM': 'Industrial', 'Standard Industrial': 'Industrial', 'T&L': 'Industrial', 'Vehicle': 'Industrial', 'AS':'Industrial',
        'Miscellaneous': 'Misc', 'PI': 'PI', 'Retail': 'Retail', 'PRO': 'PRO',
        'Adjustment figure': 'SERVICE'  
    }
    hyperion_to_bi_dpc_map = {
        'OEM_SI': 'OEM', 'SI_S': 'Standard Industrial', 'TL': 'T&L', 'Misc':'Miscellaneous', 'Pro':'PRO', 'AC':'AutoChem'
    }
    
    adjustments_to_add = []
    
    for hyperion_dpc_str, hyperion_value_mtd in hyperion_dpc_map_mtd.items():
        if pd.isna(hyperion_dpc_str) or "Total" in hyperion_dpc_str: continue
        
        bi_dpc_str = hyperion_to_bi_dpc_map.get(hyperion_dpc_str, hyperion_dpc_str)
        
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
                new_row['P1-Division'] = dpc_to_division_map.get(bi_dpc_str, 'Adjustment figure')

            adjustments_to_add.append(new_row)

    service_mtd_val = float(last_row_mtd) * 1000
    service_py_val = float(last_row_py) * 1000

    if abs(service_mtd_val) > 0.001 or abs(service_py_val) > 0.001:
        print(f"           -> Adding 'SERVICE' adjustment row from Hyperion total (MTD: {service_mtd_val:.2f}, PY: {service_py_val:.2f})")
        service_row = {col: 'Adjustment figure' for col in bi_df.columns}
        
        if 'Product/Service' in service_row:
            service_row['Product/Service'] = 'SERVICE'
            
        if 'P1-Division' in service_row:
            service_row['P1-Division'] = dpc_to_division_map.get(service_row['P1-Division'], 'SERVICE')
            
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


def process_excel_files(folder_path, output_folder, hyperion_folder_path, directory_file_path, currency_file_path):
    os.makedirs(output_folder, exist_ok=True)
    excel_files = glob.glob(os.path.join(folder_path, "*.xlsx"))
    processed_files = []

    month_map = {
        'Jan': '01', 'Feb': '02', 'Mar': '03', 'Apr': '04', 'May': '05', 'Jun': '06',
        'Jul': '07', 'Aug': '08', 'Sep': '09', 'Oct': '10', 'Nov': '11', 'Dec': '12'
    }
    
    if not excel_files:
        print(f"No Excel files found in {folder_path}")
        return processed_files
        
    # --- NEW: Load currency info ONCE ---
    print("Loading currency and directory info...")
    # --- MODIFIED: Unpack two return values ---
    mo_comp_numbers, currency_map = load_directory_info(directory_file_path)
    
    if mo_comp_numbers is None or currency_map is None:
          print("❌ CRITICAL: Could not load currency directory file. Aborting.")
          return []
    rates_cache = {} # Initialize cache for currency rates
    # --- END NEW ---
    
    for file_path in excel_files:
        filename = os.path.basename(file_path)
        try:
            # --- MODIFIED: New filename parsing ---
            # Format: Order Entry_Sep_2025_DK01_2033_MO.xlsx
            match = re.search(r'Order Entry_([A-Za-z]{3})_(\d{4})_([A-Z0-9]+)_(\d+)_([A-Z0-9]+)\.xlsx', filename, re.IGNORECASE)
            if not match:
                print(f"❌ Skipping {filename}: Filename does not match 'Order Entry_MMM_YYYY_Unit_CompNo_Type.xlsx' format.")
                continue

            month_abbr, year_str, unit, profit_center, type_mo = match.groups()
            month_int = datetime.strptime(month_abbr, '%b').month
            year_int = int(year_str)
            print(f"Processing '{filename}'... (PC: {profit_center}, Unit: {unit}, Date: {month_abbr}-{year_str})")
            # --- END MODIFICATION ---

            # --- NEW: Get Currency Conversion Rates ---
            cross_rate_current, cross_rate_prev = 1.0, 1.0
            conversion_needed = False
            
            # --- MODIFIED: Use (profit_center, unit) as the key ---
            # profit_center = Comp_No_for_OE
            # unit = SAP_Comp_Code
            currency_key = (str(profit_center), str(unit))
            
            if currency_key in currency_map:
                curr_info = currency_map[currency_key]
                source_curr = curr_info.get('Original Currency')
                target_curr = curr_info.get('Conversion Currency')

                if pd.notna(source_curr) and pd.notna(target_curr) and source_curr != target_curr:
                    print(f"   Currency conversion required for {currency_key}: {source_curr} -> {target_curr}")
                    rates_key = f"{month_int}-{year_int}"
                    if rates_key not in rates_cache:
                        print(f"   ...loading currency rates for {rates_key}")
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
                    print(f"   No currency conversion needed for {currency_key}.")
            else:
                print(f"   No currency info found for key (Comp_No/PC, SAP_Code/Unit): {currency_key}.")
            # --- END CURRENCY GET ---

            df = pd.read_excel(file_path, sheet_name='Sheet1', header=None)

            # --- NEW: Check for "No applicable data found" ---
            if not df.empty and isinstance(df.iloc[0, 0], str) and "no applicable data found" in df.iloc[0, 0].lower():
                print(f"   -> Skipping '{filename}': File contains 'No applicable data found'.\n")
                continue # Skip to the next file
            # --- END NEW ---

            if df.shape[1] >= 12:
                df.iloc[:, 0:12] = df.iloc[:, 0:12].shift(-1)
            
            if df.shape[1] >= 14 and len(df) > 1:
                df.iloc[1:, 12:14] = df.iloc[1:, 12:14].shift(-1)

            if 7 < df.shape[1]:
                df[7] = df[7].astype(str).str.replace('#', 'non-holding', regex=False)

            if not df.empty:
                new_headers = df.iloc[0].astype(str).str.replace(r'\s+', ' ', regex=True).str.strip()
                df.columns = new_headers
                df = df.iloc[1:].reset_index(drop=True)
            
            # --- [START] NEW CONDITIONAL 3RD FILTER ---
            # Check if the profit center (Comp_No_for_OE) is in the MO/MOPO list
            if str(profit_center) in mo_comp_numbers:
                print(f"   -> File is for MO/MOPO Comp_No '{profit_center}'. Applying '3RD' only filter.")
                
                # Column K in the original Excel file is at index 10
                COL_K_INDEX = 10
                if not df.empty and COL_K_INDEX < len(df.columns):
                    # Get the name of the column that was originally Column K
                    col_k_name = df.columns[COL_K_INDEX] 
                    
                    print(f"   -> Filtering for '3RD' on column: '{col_k_name}' (Original Col K)")
                    
                    initial_row_count_k = len(df)
                    
                    # Filter to keep rows where the value is *exactly* '3RD'
                    # .astype(str).str.strip() handles NaNs and whitespace
                    df = df[df[col_k_name].astype(str).str.strip() == '3RD'].copy()
                    
                    rows_removed_k = initial_row_count_k - len(df)
                    
                    if rows_removed_k > 0:
                        print(f"   -> Removed {rows_removed_k} rows that did not match '3RD' in '{col_k_name}'.")
                    else:
                        print(f"   -> No rows removed by '3RD' filter.")

                    if df.empty:
                        print(f"   -> No data remaining after '3RD' filter. Skipping file.\n")
                        continue # Skip to the next file
                        
                elif not df.empty:
                    # This case means df is not empty, but it has fewer than 11 columns
                    print(f"   ⚠️ Warning: File has fewer than 11 columns. Cannot apply '3RD' filter on Column K.")
            
            else:
                print(f"   -> Skipping '3RD' filter: Comp_No '{profit_center}' is not in the MO/MOPO list. Keeping all data (3RD and IC).")
            # --- [END] NEW CONDITIONAL 3RD FILTER ---

            if 'Type - Sales Document' in df.columns:
                df.rename(columns={'Type - Sales Document': 'Sales doc. type'}, inplace=True)
            doc_type_col = 'Sales doc. type'

            product_service_col = 'Product/Service'
            dist_channel_col = 'Distribution Channel'

            for col in df.select_dtypes(include=['object']).columns:
                df[col] = df[col].apply(lambda x: x.replace(',', '_').replace('/', '_') if isinstance(x, str) else x)
            
            classification_map = {
                'mt arm return': 'PRODUCT', 'mt credit memo req': 'PRODUCT',
                'mt debit memo req': 'PRODUCT', 'mt eco order hybris': 'PRODUCT',
                'mt epro order b2b': 'PRODUCT', 'mt standard order': 'PRODUCT',
                'mt rental deb req': 'SERVICE', 'mt svc conf dmr': 'SERVICE',
                'mt svc contract dmr': 'SERVICE', 'pipette svc order': 'SERVICE',
                'mt free of charge': 'SERVICE', 'mt int cred memo req': 'PRODUCT'
            }

            if doc_type_col in df.columns and product_service_col in df.columns:
                source_series = df[doc_type_col].astype(str).str.strip().str.lower()
                df[product_service_col] = source_series.map(classification_map).fillna('Product')

            if product_service_col in df.columns:
                initial_row_count = len(df)
                df = df[df[product_service_col] != 'SERVICE']
                rows_removed = initial_row_count - len(df)
                if rows_removed > 0:
                    print(f"   -> Removed {rows_removed} rows where '{product_service_col}' was 'SERVICE'.")

            if dist_channel_col in df.columns:
                df[dist_channel_col] = df[dist_channel_col].astype(str).str.replace('#', 'non-holding', regex=False)
            
            # if not df.empty:
            #     df.drop(df.tail(1).index, inplace=True)

            if not df.empty:
                df.drop(columns=df.columns[0], inplace=True)
            
            p2_dpc_column_name = 'P2-DPC'

            if p2_dpc_column_name in df.columns:
                df[p2_dpc_column_name] = df[p2_dpc_column_name].replace('Std Industrial', 'Standard Industrial')

            print(f"   -> Using '{p2_dpc_column_name}' as the DPC column for adjustments.")
            
            # ==================================================================
            # --- LOGIC RE-ORDERED ---
            # ==================================================================

            # --- [MOVED UP] STEP 1: Apply Currency Conversion to the main 'df' ---
            mtd_col_name = 'Bookings MTD Net Sales'
            py_col_name = 'Bookings PY MTD'
            
            if conversion_needed and mtd_col_name in df.columns and py_col_name in df.columns:
                print(f"   -> Applying currency conversion to main BI data...")
                df[mtd_col_name] = pd.to_numeric(df[mtd_col_name], errors='coerce').fillna(0) * cross_rate_current
                df[py_col_name] = pd.to_numeric(df[py_col_name], errors='coerce').fillna(0) * cross_rate_prev
                print(f"   -> Conversion applied to '{mtd_col_name}' and '{py_col_name}' in main dataframe.")
            elif conversion_needed:
                print(f"         ⚠️  Warning: Could not find '{mtd_col_name}' or '{py_col_name}' in main df. Skipping conversion.")
            # --- END CURRENCY CONVERSION ---
            
            # --- [MOVED DOWN] STEP 2: Add Hyperion adjustments *after* conversion ---
            df = add_hyperion_adjustments(df, profit_center, month_abbr, year_str, hyperion_folder_path, p2_dpc_column_name)
            
            # --- MODIFIED: Use parsed arguments for output filename ---
            base_output_filename = f"OE_Data_Processed_{unit}_{profit_center}_{month_map.get(month_abbr, 'MM')}{year_str[-2:]}.csv"
            output_path = os.path.join(output_folder, base_output_filename)
            
            # --- NEW: Check for existing file and append (1), (2), etc. ---
            counter = 1
            final_output_filename = base_output_filename
            
            while os.path.exists(output_path):
                # Split the base name and extension
                base_name, extension = os.path.splitext(base_output_filename)
                # Create new filename
                final_output_filename = f"{base_name}({counter}){extension}"
                # Create new full path
                output_path = os.path.join(output_folder, final_output_filename)
                # Increment counter for the next check
                counter += 1
            
            if final_output_filename != base_output_filename:
                print(f"   ⚠️ File '{base_output_filename}' already exists. Saving as '{final_output_filename}'")
            # --- END NEW ---
            
            # --- STEP 3: Select specific columns for the output file ---
            output_cols = []
            all_cols = list(df.columns)
            
            # Corresponds to original columns B-J (9 columns)
            if len(all_cols) >= 9:
                output_cols.extend(all_cols[0:9])

            # Corresponds to original columns M-N (indices 11 and 12 of the processed dataframe)
            if len(all_cols) >= 13:
                output_cols.extend(all_cols[11:13])
            
            if not output_cols:
                print(f"    ⚠️ Warning: DataFrame has too few columns to select B-J and M-N. Writing all available columns.")
                df_output = df
            else:
                print(f"   -> Selecting columns corresponding to original B-J and M-N for the output file.")
                df_output = df[output_cols].copy() # Use .copy() to avoid SettingWithCopyWarning
            # --- MODIFICATION END ---
            
            # --- [This block was moved up] ---

            df_output.to_csv(output_path, index=False, encoding='utf-8-sig')
            
            processed_files.append(final_output_filename) # <-- Use the final, unique filename
            print(f"✅ Successfully processed '{filename}' -> '{final_output_filename}'\n") # <-- Use the final, unique filename
            
        except Exception as e:
            print(f"❌ Error processing {filename}: {str(e)}\n")
            
    return processed_files