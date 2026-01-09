import os
import pandas as pd
import glob
import re  # Added for new filename parsing
import calendar  # Added for currency conversion
from datetime import datetime  # Added for currency conversion

# --- IMPORT MOVED FUNCTIONS ---
try:
    # This function is now only called when grouping is OFF
    from cleaning_configurations import add_hyperion_adjustments
except ImportError:
    print("Warning: Could not import 'add_hyperion_adjustments'. Adjustments will not work.")
    # Create a fallback function
    def add_hyperion_adjustments(bi_df, *args):
        print("         ❌ ERROR: add_hyperion_adjustments function not found!")
        return bi_df
# --- END IMPORT ---


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
    3. A dictionary mapping 'Comp_No' to 'Comp_No_for_OE'.
    """
    print(f"Reading directory file from: {directory_file_path}")
    try:
        df_dir = pd.read_excel(directory_file_path)
    except FileNotFoundError:
        print(f"❌ ERROR: Directory file not found at: {directory_file_path}")
        return None, None, None
    except Exception as e:
        print(f"❌ ERROR: Could not read directory file. Error: {e}")
        return None, None, None
    
    # Add 'Comp_No' to the required columns
    required_cols = ["Comp_No", "Comp_No_for_OE", "Type", "SAP_Comp_Code", "Original Currency", "Conversion Currency"]
    if not all(col in df_dir.columns for col in required_cols):
        print(f"❌ ERROR: Directory file must contain {required_cols} columns.")
        return None, None, None

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
        return None, None, None
    
    # --- 2. Build Currency Map ---
    try:
        key_cols = ['Comp_No_for_OE', 'SAP_Comp_Code']
        currency_data_cols = ['Original Currency', 'Conversion Currency']
        
        df_currencies = df_dir[key_cols + currency_data_cols].copy()
        df_currencies = df_currencies.dropna(subset=key_cols + currency_data_cols)

        for col in key_cols:
            df_currencies[col] = df_currencies[col].astype(str)

        df_currencies = df_currencies.drop_duplicates()
        duplicates = df_currencies.duplicated(subset=key_cols, keep=False)
        
        if duplicates.any():
            conflicting_data = df_currencies[duplicates].sort_values(by=key_cols)
            print("❌ ERROR: Conflicting currency data found. The same (Comp_No_for_OE, SAP_Comp_Code) pair points to different currencies.")
            print("Please fix these in the Directory_Processed_Output.xlsx file:")
            print(conflicting_data.to_string(index=False))
            return None, None, None # Stop execution

        df_currencies = df_currencies.set_index(key_cols)
        currency_map = df_currencies.to_dict('index')
        
        print(f"Loaded currency mapping for {len(currency_map)} (Comp_No, SAP_Code) pairs.")
        
    except Exception as e:
        print(f"❌ ERROR: Failed to build currency map. Error: {e}")
        return None, None, None

    # --- 3. Build Comp_No -> Comp_No_for_OE Map (NEW) ---
    try:
        df_map = df_dir.dropna(subset=['Comp_No', 'Comp_No_for_OE'])
        
        # Convert to string for reliable matching
        df_map['Comp_No'] = df_map['Comp_No'].astype(str).str.strip()
        df_map['Comp_No_for_OE'] = df_map['Comp_No_for_OE'].astype(str).str.strip()
        
        # Drop duplicates based on Comp_No
        df_map = df_map.drop_duplicates(subset=['Comp_No'])
        
        comp_no_to_oe_map = pd.Series(df_map['Comp_No_for_OE'].values, index=df_map['Comp_No']).to_dict()
        print(f"Loaded {len(comp_no_to_oe_map)} Comp_No -> Comp_No_for_OE entries.")

    except Exception as e:
        print(f"❌ Error loading Comp_No -> OE map: {e}")
        return None, None, None

    # Return all three items
    return comp_numbers_to_process, currency_map, comp_no_to_oe_map

def load_currency_rates(currency_file_path, month_int, year_int):
    """
    Loads currency rates from the specified file for the given month/year.
    
    --- UPDATED ---
    Dynamically reads from the correct sheet based on the month (e.g., 'Sep', 'Oct').
    """
    print(f"    ...loading currency rates for {month_int}/{year_int}")
    try:
        # Get the 3-letter month abbreviation (e.g., 9 -> 'Sep')
        month_abbr = calendar.month_abbr[month_int]
        
        print(f"       -> Reading sheet: '{month_abbr}'")

        month_name = calendar.month_name[month_int]
        current_year_col = f"{month_name} {year_int}"
        prev_year_col = f"{month_name} {year_int - 1}"

        # --- MODIFIED: Use sheet_name=month_abbr ---
        df_headers = pd.read_excel(currency_file_path, sheet_name=month_abbr, header=1, nrows=0)
        
        current_year_col_actual = next((col for col in df_headers.columns if col.strip().lower() == current_year_col.lower()), None)
        prev_year_col_actual = next((col for col in df_headers.columns if col.strip().lower() == prev_year_col.lower()), None)

        if not current_year_col_actual or not prev_year_col_actual:
            print(f"❌ ERROR: Currency file (sheet '{month_abbr}') missing required columns. ")
            print(f"   Expected: '{current_year_col}' and '{prev_year_col}'")
            print(f"   Actual headers found: {list(df_headers.columns)}")
            return None

        use_cols = ['Currency', current_year_col_actual, prev_year_col_actual]
        
        # --- MODIFIED: Use sheet_name=month_abbr ---
        df_rates = pd.read_excel(currency_file_path, sheet_name=month_abbr, header=1, usecols=use_cols)
        
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

# --- FUNCTIONS _extract_dpc_maps_from_sheet AND add_hyperion_adjustments
# --- have been MOVED to cleaning_configurations.py

# ### MODIFIED FUNCTION ###
def process_excel_files(folder_path, output_folder, hyperion_folder_path, directory_file_path, currency_file_path, group_units=False):
    """
    Processes all Excel files in the given folder.
    If group_units is True, it skips Hyperion adjustments, which will be done
    by the grouping function after merging.
    """
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
        
    print("Loading currency and directory info...")
    # --- MODIFIED: Unpack three return values ---
    mo_comp_numbers, currency_map, comp_no_to_oe_map = load_directory_info(directory_file_path)
    
    if mo_comp_numbers is None or currency_map is None or comp_no_to_oe_map is None:
          print("❌ CRITICAL: Could not load currency directory file. Aborting.")
          return []
    rates_cache = {} # Initialize cache for currency rates
    
    for file_path in excel_files:
        filename = os.path.basename(file_path)
        try:
            # Format: Order Entry_Sep_2025_DK01_2033_MO.xlsx
            match = re.search(r'Order Entry_([A-Za-z]{3})_(\d{4})_([A-Z0-9]+)_(\d+)_([A-Z0-9]+)\.xlsx', filename, re.IGNORECASE)
            if not match:
                print(f"❌ Skipping {filename}: Filename does not match 'Order Entry_MMM_YYYY_Unit_CompNo_Type.xlsx' format.")
                continue

            # profit_center variable now holds the Comp_No from the filename (e.g., '5231')
            month_abbr, year_str, unit, profit_center, type_mo = match.groups()
            month_int = datetime.strptime(month_abbr, '%b').month
            year_int = int(year_str)
            print(f"Processing '{filename}'... (PC/Comp_No: {profit_center}, Unit: {unit}, Date: {month_abbr}-{year_str})")

            # --- [NEW] Map Comp_No from file to Comp_No_for_OE ---
            pc_from_file = profit_center # e.g., '5231'
            comp_no_for_oe = comp_no_to_oe_map.get(str(pc_from_file))
            
            if not comp_no_for_oe:
                print(f"   ❌ Skipping '{filename}': Comp_No '{pc_from_file}' not found in directory's 'Comp_No' column.")
                continue 
            
            print(f"   -> File Comp_No '{pc_from_file}' mapped to Comp_No_for_OE '{comp_no_for_oe}'.")
            # --- [END NEW] ---

            # --- Get Currency Conversion Rates ---
            cross_rate_current, cross_rate_prev = 1.0, 1.0
            conversion_needed = False
            
            # --- MODIFIED: Use (comp_no_for_oe, unit) as the key ---
            currency_key = (str(comp_no_for_oe), str(unit))
            
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
                print(f"   No currency info found for key (Comp_No_for_OE, SAP_Code/Unit): {currency_key}.")
            # --- END CURRENCY GET ---

            df = pd.read_excel(file_path, sheet_name='Sheet1', header=None)

            if not df.empty and isinstance(df.iloc[0, 0], str) and "no applicable data found" in df.iloc[0, 0].lower():
                print(f"   -> Skipping '{filename}': File contains 'No applicable data found'.\n")
                continue 

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
            
            # --- [START] MODIFIED CONDITIONAL 3RD FILTER ---
            # Check if the *mapped* Comp_No_for_OE is in the MO/MOPO list
            if str(comp_no_for_oe) in mo_comp_numbers:
                print(f"   -> File is for MO/MOPO Comp_No_for_OE '{comp_no_for_oe}'. Applying '3RD' only filter.")
                
                COL_K_INDEX = 10
                if not df.empty and COL_K_INDEX < len(df.columns):
                    col_k_name = df.columns[COL_K_INDEX] 
                    print(f"   -> Filtering for '3RD' on column: '{col_k_name}' (Original Col K)")
                    initial_row_count_k = len(df)
                    df = df[df[col_k_name].astype(str).str.strip() == '3RD'].copy()
                    rows_removed_k = initial_row_count_k - len(df)
                    
                    if rows_removed_k > 0:
                        print(f"   -> Removed {rows_removed_k} rows that did not match '3RD' in '{col_k_name}'.")
                    else:
                        print(f"   -> No rows removed by '3RD' filter.")

                    if df.empty:
                        print(f"   -> No data remaining after '3RD' filter. Skipping file.\n")
                        continue 
                        
                elif not df.empty:
                    print(f"   ⚠️ Warning: File has fewer than 11 columns. Cannot apply '3RD' filter on Column K.")
            
            else:
                print(f"   -> Skipping '3RD' filter: Comp_No_for_OE '{comp_no_for_oe}' is not in the MO/MOPO list. Keeping all data (3RD and IC).")
            # --- [END] MODIFIED CONDITIONAL 3RD FILTER ---

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
                source_series = df[doc_type_col].astype(str).str.strip().str.lower()
                mapped_series = source_series.map(classification_map)
                df[product_service_col] = mapped_series.fillna('Product')
                unique_doc_types = pd.Series(source_series.unique()).dropna().astype(str)
                normalized_map_keys = set(classification_map.keys())
                unmapped_types = [t for t in unique_doc_types if t and t not in normalized_map_keys]

                if unmapped_types:
                    print(f"   ⚠️ Unmapped sales doc types found ({len(unmapped_types)} distinct): {unmapped_types}")
                else:
                    print("   -> All sales doc types in file are mapped by classification_map.")
                unmapped_mask = mapped_series.isna()
                unmapped_count = int(unmapped_mask.sum())
                if unmapped_count > 0:
                    print(f"   -> {unmapped_count} rows defaulted to 'Product' because doc type was unmapped or missing.")
                    sample_cols = [doc_type_col]
                    for c in ['Bookings MTD Net Sales', 'Bookings PY MTD']:
                        if c in df.columns:
                            sample_cols.append(c)

                    sample_preview = df.loc[unmapped_mask, sample_cols].head(20)
                    if not sample_preview.empty:
                        print("   -> Sample unmapped rows (showing doc type and numeric columns):")
                        print(sample_preview.to_string(index=False))
                    if 'none' in [t.lower() for t in unmapped_types]:
                        none_mask = source_series == 'none'
                        none_count = int(none_mask.sum())
                        if none_count > 0:
                            print(f"   -> Removing {none_count} rows where doc type is 'none'.")
                            df = df.loc[~none_mask].copy()
                            source_series = df[doc_type_col].astype(str).str.strip().str.lower()
                            mapped_series = source_series.map(classification_map)
                            df[product_service_col] = mapped_series.fillna('Product')
                else:
                    print("   -> No rows defaulted to 'Product'.")

            if product_service_col in df.columns:
                initial_row_count = len(df)
                df = df[df[product_service_col] != 'SERVICE']
                rows_removed = initial_row_count - len(df)
                if rows_removed > 0:
                    print(f"   -> Removed {rows_removed} rows where '{product_service_col}' was 'SERVICE'.")

            if dist_channel_col in df.columns:
                df[dist_channel_col] = df[dist_channel_col].astype(str).str.replace('#', 'non-holding', regex=False)
            
            if not df.empty:
                df.drop(columns=df.columns[0], inplace=True)
            
            p2_dpc_column_name = 'P2-DPC'

            if p2_dpc_column_name in df.columns:
                df[p2_dpc_column_name] = df[p2_dpc_column_name].replace('Std Industrial', 'Standard Industrial')

            print(f"   -> Using '{p2_dpc_column_name}' as the DPC column for adjustments.")
            
            mtd_col_name = 'Bookings MTD Net Sales'
            py_col_name = 'Bookings PY MTD'
            
            if conversion_needed and mtd_col_name in df.columns and py_col_name in df.columns:
                print(f"   -> Applying currency conversion to main BI data...")
                df[mtd_col_name] = pd.to_numeric(df[mtd_col_name], errors='coerce').fillna(0) * cross_rate_current
                df[py_col_name] = pd.to_numeric(df[py_col_name], errors='coerce').fillna(0) * cross_rate_prev
                print(f"   -> Conversion applied to '{mtd_col_name}' and '{py_col_name}' in main dataframe.")
            elif conversion_needed:
                print(f"         ⚠️  Warning: Could not find '{mtd_col_name}' or '{py_col_name}' in main df. Skipping conversion.")
            
            if not group_units:
                print(f"   -> Grouping is OFF. Adding Hyperion adjustments before saving.")
                # --- MODIFIED: Pass the correct comp_no_for_oe to adjustments ---
                df = add_hyperion_adjustments(df, comp_no_for_oe, month_abbr, year_str, hyperion_folder_path, p2_dpc_column_name)
            else:
                print(f"   -> Grouping is ON. Skipping Hyperion adjustments (will be done after merge).")
            
            month_mm = month_map.get(month_abbr, 'MM')
            year_yy = year_str[-2:]
            date_part = f"{month_mm}{year_yy}"
            
            # --- Filename remains based on the *original* profit_center from the file ---
            base_output_filename = f"OE_Data_Processed_{unit}_{profit_center}_{date_part}.csv"
            output_path = os.path.join(output_folder, base_output_filename)
            
            counter = 1
            final_output_filename = base_output_filename
            
            while os.path.exists(output_path):
                base_name, extension = os.path.splitext(base_output_filename)
                final_output_filename = f"{base_name}({counter}){extension}"
                output_path = os.path.join(output_folder, final_output_filename)
                counter += 1
            
            if final_output_filename != base_output_filename:
                print(f"   ⚠️ File '{base_output_filename}' already exists. Saving as '{final_output_filename}'")
            
            output_cols = []
            all_cols = list(df.columns)
            
            if len(all_cols) >= 9:
                output_cols.extend(all_cols[0:9])

            if len(all_cols) >= 13:
                output_cols.extend(all_cols[11:13])
            
            if not output_cols:
                print(f"    ⚠️ Warning: DataFrame has too few columns to select B-J and M-N. Writing all available columns.")
                df_output = df
            else:
                print(f"   -> Selecting columns corresponding to original B-J and M-N for the output file.")
                df_output = df[output_cols].copy() 
            
            df_output.to_csv(output_path, index=False, encoding='utf-8-sig')
            
            processed_files.append(final_output_filename) 
            print(f"✅ Successfully processed '{filename}' -> '{final_output_filename}'\n") 
            
        except Exception as e:
            print(f"❌ Error processing {filename}: {str(e)}\n")
            import traceback
            traceback.print_exc()
            
    return processed_files