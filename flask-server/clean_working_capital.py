# clean_working_capital.py

import pandas as pd
import sys
import os
from datetime import datetime
from dateutil.relativedelta import relativedelta

# --- Configuration for Overhead Calculation ---
OVERHEAD_KEYS_TO_EXTRACT = [
    "Variable Production", "Inventory Revaluation", "Fixed Manufacturing Costs",
    "Logistics & Purchasing Costs", "Total Sales - Net", "Contribution 1"
]
OVERHEAD_KEY_MAPPING = {
    "Variable Production": "LTM Variable production costs (account \"VPC\")",
    "Inventory Revaluation": "LTM Inventory revaluation (account 7053)",
    "Fixed Manufacturing Costs": "LTM Fixed Manufacturing costs (account \"FMC\")",
    "Logistics & Purchasing Costs": "LTM Logistics & Purchasing costs (account \"LPC\")",
    "Total Sales - Net": "Sales",
    "Contribution 1": "C1"
}

# --- Configuration for DSO Calculation ---
DSO_KEYS_TO_EXTRACT = ['Sales to third']


# --- Helper Functions ---

def _convert_month_to_int(month_input):
    """Safely converts a month string (e.g., 'Aug', '8') or integer to an integer."""
    if isinstance(month_input, int):
        if 1 <= month_input <= 12:
            return month_input
        else:
            raise ValueError(f"Invalid month number: {month_input}")

    if isinstance(month_input, str):
        month_str = month_input.strip()
        if month_str.isdigit():
            month_int = int(month_str)
            if 1 <= month_int <= 12:
                return month_int
            else:
                raise ValueError(f"Invalid month number: {month_str}")
        else:
            month_map = {
                'jan': 1, 'january': 1, 'feb': 2, 'february': 2, 'mar': 3, 'march': 3,
                'apr': 4, 'april': 4, 'may': 5, 'jun': 6, 'june': 6, 'jul': 7, 'july': 7,
                'aug': 8, 'august': 8, 'sep': 9, 'september': 9, 'oct': 10, 'october': 10,
                'nov': 11, 'november': 11, 'dec': 12, 'december': 12
            }
            month_num = month_map.get(month_str.lower())
            if month_num:
                return month_num
            else:
                raise ValueError(f"Invalid month name: {month_input}")
    
    raise TypeError(f"Unsupported month format: {month_input}")


def get_ltm_months(end_month, end_year):
    """Generates a list of the last 12 months (LTM)."""
    # --- FIXED: Use the new helper function for safe conversion ---
    month_num = _convert_month_to_int(end_month)
    end_date = datetime(int(end_year), month_num, 1)
    return [(end_date - relativedelta(months=i)).strftime('%b, %Y') for i in range(11, -1, -1)]

def get_l3m_months(end_month, end_year):
    """Generates a list of the last 3 months (L3M)."""
    # --- FIXED: Use the new helper function for safe conversion ---
    month_num = _convert_month_to_int(end_month)
    end_date = datetime(int(end_year), month_num, 1)
    return [(end_date - relativedelta(months=i)).strftime('%b, %Y') for i in range(2, -1, -1)]


# --- Overhead Calculation Logic ---

def extract_ke30_data_for_overhead(file_path):
    """Extracts data from column M of a KE30 file for Overhead calculations."""
    df = pd.read_excel(file_path, sheet_name='Sheet1', header=None)
    if len(df.columns) < 13:
        raise ValueError(f"File {os.path.basename(file_path)} has fewer than 13 columns.")
    col_a, col_m = df.iloc[:, 0], df.iloc[:, 12]
    data_dict = {}
    for key in OVERHEAD_KEYS_TO_EXTRACT:
        mask = col_a.fillna('').astype(str).str.strip() == key.strip()
        if mask.any():
            value = col_m[mask].iloc[0]
            data_dict[key] = 0 if pd.isna(value) else value
        else:
            print(f"Warning: '{key}' not found in {os.path.basename(file_path)} - setting to 0")
            data_dict[key] = 0
    return data_dict

def calculate_inventory_ytd(balance_file, end_month, end_year):
    """Calculates LTM values for Inventory from the Balance sheet."""
    df = pd.read_excel(balance_file, sheet_name='MT-A', header=14)
    ltm_current_months = get_ltm_months(end_month, end_year)
    ltm_previous_months = get_ltm_months(end_month, end_year - 1)
    
    ltm_current = sum(df[month].sum() for month in ltm_current_months if month in df.columns and pd.notna(df[month].sum()))
    ltm_previous = sum(df[month].sum() for month in ltm_previous_months if month in df.columns and pd.notna(df[month].sum()))
    
    return ltm_current, ltm_previous

def calculate_overhead_summary(upload_folder, output_folder, end_month='Aug', end_year=2025):
    """Main function to process overhead data from 6 input files."""
    # Define required files
    files = {
        'ke30_month_cy': os.path.join(upload_folder, "KE30 Month CY.xlsx"),
        'ke30_month_py': os.path.join(upload_folder, "KE30 Month PY.xlsx"),
        'ke30_month_py1': os.path.join(upload_folder, "KE30 Month PY-1.xlsx"),
        'ke30_ytd_cy': os.path.join(upload_folder, "KE30 YTD CY.xlsx"),
        'ke30_ytd_py': os.path.join(upload_folder, "KE30 YTD PY.xlsx"),
        'balance': os.path.join(upload_folder, "Balance.xlsx"),
    }
    for f_path in files.values():
        if not os.path.exists(f_path):
            raise FileNotFoundError(f"Required file not found for Overhead calculation: {os.path.basename(f_path)}")

    # Extract data
    data_month_cy = extract_ke30_data_for_overhead(files['ke30_month_cy'])
    data_month_py = extract_ke30_data_for_overhead(files['ke30_month_py'])
    data_month_py1 = extract_ke30_data_for_overhead(files['ke30_month_py1'])
    data_ytd_cy = extract_ke30_data_for_overhead(files['ke30_ytd_cy'])
    data_ytd_py = extract_ke30_data_for_overhead(files['ke30_ytd_py'])
    inventory_current, inventory_previous = calculate_inventory_ytd(files['balance'], end_month, end_year)

    # Calculate YTD values
    ytd_current_values, ytd_previous_values, overhead_names = [], [], []
    for key in OVERHEAD_KEYS_TO_EXTRACT:
        overhead_names.append(OVERHEAD_KEY_MAPPING.get(key, key))
        ytd_cy = data_month_cy.get(key, 0) + data_ytd_cy.get(key, 0) - data_month_py.get(key, 0)
        ytd_py = data_month_py.get(key, 0) + data_ytd_py.get(key, 0) - data_month_py1.get(key, 0)
        ytd_current_values.append(ytd_cy)
        ytd_previous_values.append(ytd_py)
    
    overhead_names.append('LTM Inventory net (account "Inv")')
    ytd_current_values.append(inventory_current)
    ytd_previous_values.append(inventory_previous)
    
    # Create DataFrame
    variance = [c - p for c, p in zip(ytd_current_values, ytd_previous_values)]
    variance_pct = [round((v / p) * 100, 2) if p != 0 else (0 if v == 0 else float('inf')) for v, p in zip(variance, ytd_previous_values)]
    
    # --- FIXED: Use helper function for month display ---
    month_num = _convert_month_to_int(end_month)
    month_display = datetime(2000, month_num, 1).strftime('%b').upper()
    
    df_output = pd.DataFrame({
        'Overhead Name': overhead_names,
        f'YTD {month_display} {end_year}': ytd_current_values,
        f'YTD {month_display} {end_year-1}': ytd_previous_values,
        'Variance': variance, 'Variance %': variance_pct
    })
    
    # Save to Excel
    output_filename = "overhead_comparison.xlsx"
    output_path = os.path.join(output_folder, output_filename)
    df_output.to_excel(output_path, index=False, sheet_name='Overhead Data')
    return [output_filename]


# --- DSO Calculation Logic ---

def extract_ke30_data_for_dso(file_path):
    """Extracts data from column H of a KE30 file for DSO calculations."""
    df = pd.read_excel(file_path, sheet_name='Sheet1', header=None)
    if len(df.columns) < 8:
        raise ValueError(f"File {os.path.basename(file_path)} has fewer than 8 columns.")
    col_a, col_h = df.iloc[:, 0], df.iloc[:, 7]
    data_dict = {}
    for key in DSO_KEYS_TO_EXTRACT:
        mask = col_a.fillna('').astype(str).str.strip() == key.strip()
        if mask.any():
            value = col_h[mask].iloc[0]
            data_dict[key] = 0 if pd.isna(value) else value
        else:
            print(f"Warning: '{key}' not found in {os.path.basename(file_path)} - setting to 0")
            data_dict[key] = 0
    return data_dict

def calculate_dso_summary(upload_folder, output_folder, end_month='Aug', end_year=2025):
    """Processes financial data to create a detailed L3M summary for DSO."""
    # Define required files
    files = {
        'balance': os.path.join(upload_folder, "Balance.xlsx"),
        'ke30_cy': os.path.join(upload_folder, "KE30 Month CY.xlsx"),
        'ke30_py': os.path.join(upload_folder, "KE30 Month PY.xlsx"),
    }
    for f_path in files.values():
        if not os.path.exists(f_path):
            raise FileNotFoundError(f"Required file not found for DSO calculation: {os.path.basename(f_path)}")

    # 1. Balance Sheet L3M Calculations
    balance_df = pd.read_excel(files['balance'], sheet_name='MT-A', header=14)
    balance_df.rename(columns={balance_df.columns[0]: 'Account'}, inplace=True)
    balance_df['Account'] = balance_df['Account'].astype(str).str.strip()
    
    l3m_cy_cols = get_l3m_months(end_month, end_year)
    l3m_py_cols = get_l3m_months(end_month, end_year - 1)

    tar_avg_cy = balance_df[balance_df['Account'].str.startswith('110')][l3m_cy_cols].sum().mean()
    prepayment_avg_cy = balance_df[balance_df['Account'].str.startswith('360')][l3m_cy_cols].sum().mean()
    tar_avg_py = balance_df[balance_df['Account'].str.startswith('110')][l3m_py_cols].sum().mean()
    prepayment_avg_py = balance_df[balance_df['Account'].str.startswith('360')][l3m_py_cols].sum().mean()
    
    # 2. Extract Data from KE30 Files
    ke30_data_cy = extract_ke30_data_for_dso(files['ke30_cy'])
    ke30_data_py = extract_ke30_data_for_dso(files['ke30_py'])
    
    # 3. Combine and Structure Data for Output
    sales_cy = ke30_data_cy.get('Sales to third', 0)
    sales_py = ke30_data_py.get('Sales to third', 0)
    
    rows = [
        ['Sales to third', sales_cy, sales_py],
        ['TAR AVERAGE (L3M)', tar_avg_cy * 1000, tar_avg_py * 1000],
        ['Customer Prepayment AVERAGE (L3M)', prepayment_avg_cy * -1000, prepayment_avg_py * -1000],
        ['VAT', '', ''],
    ]
    summary_df = pd.DataFrame(rows, columns=['Description', 'Current Year', 'Previous Year'])

    # 4. Write to Excel with formulas
    output_filename = "dso_financial_summary.xlsx"
    output_path = os.path.join(output_folder, output_filename)
    with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
        summary_df.to_excel(writer, sheet_name='Summary', index=False, startrow=0)
        workbook  = writer.book
        worksheet = writer.sheets['Summary']
        
        # Add formula rows
        worksheet.write('A6', 'TAR L3M - customer prepayments L3M) * (100/(100+VAT)) x 360 days')
        worksheet.write_formula('B6', '=(B3-B4)*(100/(100+B5))*360')
        worksheet.write_formula('C6', '=(C3-C4)*(100/(100+C5))*360')
        
        worksheet.write('A7', 'Sales third party L3M * 4')
        worksheet.write_formula('B7', '=B2*4')
        worksheet.write_formula('C7', '=C2*4')
        
        worksheet.write('A8', 'DSO')
        num_format = workbook.add_format({'num_format': '0.00'})
        worksheet.write_formula('B8', '=IF(B7<>0, B6/B7, 0)', num_format)
        worksheet.write_formula('C8', '=IF(C7<>0, C6/C7, 0)', num_format)

    return [output_filename]


# --- Main Entry Point ---

def process_working_capital(upload_folder, output_folder, metric):
    """
    Router function to trigger the correct working capital calculation.
    
    Parameters:
    - upload_folder: Path to the folder with uploaded source files.
    - output_folder: Path where the resulting Excel file will be saved.
    - metric: The type of calculation to perform ('dso' or 'overhead').
    
    Returns:
    - A list containing the filename(s) of the generated report(s).
    """
    # For now, we use hardcoded month/year, but this could be passed from the UI
    end_month = 'Aug'
    end_year = 2025
    
    if metric == 'dso':
        print("Starting DSO calculation...")
        return calculate_dso_summary(upload_folder, output_folder, end_month, end_year)
    elif metric == 'overhead':
        print("Starting Overhead calculation...")
        return calculate_overhead_summary(upload_folder, output_folder, end_month, end_year)
    else:
        raise ValueError(f"Unknown working capital metric: '{metric}'")