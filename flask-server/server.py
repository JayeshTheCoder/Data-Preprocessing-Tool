from flask import Flask, jsonify, request, send_from_directory
from flask_cors import CORS
import os
import uuid
import io
import zipfile
import pandas as pd
import glob
import numpy as np
from dotenv import load_dotenv
import base64
import shutil
load_dotenv()

import inference_service 
import cleaning_configurations 

# --- [NEW] IMPORT DYNAMIC PIPELINE ---
import process_all_complete

# -------------------------------------

try:
    from clean_sales import process_files_to_csv
except ImportError:
    print("Warning: Could not import 'process_files_to_csv'. Sales endpoint will not work.")
    process_files_to_csv = lambda u, o, d, c: (print("Sales processor not found!"), [])[1] 

try:
    from clean_oe import process_excel_files
except ImportError:
    print("Warning: Could not import 'process_excel_files'. OE endpoint will not work.")
    process_excel_files = lambda u, o, h, d, c, g=False: (print("OE processor not found!"), [])[1]

try:
    from clean_working_capital import process_working_capital
except ImportError:
    print("Warning: Could not import 'process_working_capital'. WC endpoints will not work.")
    process_working_capital = lambda u, o, m: (print("WC processor not found!"), [])[1]

try:
    from clean_pex import process_pex_and_headcount, process_pex_vendor
except ImportError:
    print("Warning: Could not import from 'clean_pex'. PEX endpoint will not work.")
    process_pex_and_headcount = lambda u, o, l: (print("PEX BI processor not found!"), [])[1]
    process_pex_vendor = lambda u, o, b: (print("PEX Vendor processor not found!"), [])[1]

app = Flask(__name__)
CORS(app, origins='*')

UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'output'

# --- ENV VARIABLES ---
HYPERION_FOLDER = os.environ.get("hyperion_folder_path")
PEX_LOOKUP_FOLDER = os.environ.get("pex_folder_path")
SALES_DIRECTORY_FILE = os.environ.get("sales_directory_file_path")
SALES_CURRENCY_FILE = os.environ.get("sales_currency_file_path")
PEX_HYPERION_FILE = os.environ.get("pex_hyperion_file_path") 
SALES_3RD_HYPERION_FILE = os.environ.get("sales_3rd_dpc")
SALES_IC_HYPERION_FILE = os.environ.get("sales_IC_dpc")  

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)
if HYPERION_FOLDER: os.makedirs(HYPERION_FOLDER, exist_ok=True)
if PEX_LOOKUP_FOLDER: os.makedirs(PEX_LOOKUP_FOLDER, exist_ok=True)

# --- CONFIG CHECKS ---
if not SALES_DIRECTORY_FILE or not os.path.exists(SALES_DIRECTORY_FILE):
    print("="*60)
    print("⚠️ WARNING: SALES_DIRECTORY_FILE is not set correctly.")
    print(f"   Current path: {SALES_DIRECTORY_FILE}")
    print("   Please set 'sales_directory_file_path' in your .env file")
    print("="*60)

if not SALES_CURRENCY_FILE or not os.path.exists(SALES_CURRENCY_FILE):
    print("="*60)
    print("⚠️ WARNING: SALES_CURRENCY_FILE is not set correctly.")
    print(f"   Current path: {SALES_CURRENCY_FILE}")
    print("   Please set 'sales_currency_file_path' in your .env file")
    print("="*60)

if not SALES_3RD_HYPERION_FILE or not os.path.exists(SALES_3RD_HYPERION_FILE):
    print("="*60)
    print("⚠️ WARNING: SALES_3RD_HYPERION_FILE is not set correctly.")
    print(f"   Current path: {SALES_3RD_HYPERION_FILE}")
    print("   Please set 'sales_3rd_dpc' in your .env file")
    print("="*60)

if not SALES_IC_HYPERION_FILE or not os.path.exists(SALES_IC_HYPERION_FILE):
    print("="*60)
    print("⚠️ WARNING: SALES_IC_HYPERION_FILE is not set correctly.")
    print(f"   Current path: {SALES_IC_HYPERION_FILE}")
    print("   Please set 'sales_IC_dpc' in your .env file")
    print("="*60)


# --- ROUTES ---

@app.route("/inference", methods=["POST"])
def run_inference():
    if 'file' not in request.files: return jsonify({"error": "No file part"}), 400
    file = request.files['file']
    if file.filename == '': return jsonify({"error": "No selected file"}), 400
    
    prompt = request.form.get('prompt') or inference_service.DEFAULT_PROMPT
    try:
        input_content = file.read().decode('utf-8')
        response_data = inference_service.process_markdown_file(input_content, prompt)
        final_response = { "filename": file.filename, **response_data }
        return jsonify(final_response)
    except Exception as e:
        return jsonify({"error": "An unexpected error occurred on the server.", "logs": f"Fatal error: {str(e)}"}), 500

@app.route("/inference/bulk", methods=["POST"])
def run_bulk_inference():
    if 'files' not in request.files: return jsonify({"error": "No file part"}), 400
    files = request.files.getlist('files')
    if not files: return jsonify({"error": "No files selected"}), 400
    prompt = request.form.get('prompt') or inference_service.DEFAULT_PROMPT
    all_results = []
    for file in files:
        if file and file.filename:
            try:
                input_content = file.stream.read().decode('utf-8')
                response_data = inference_service.process_markdown_file(input_content, prompt)
                all_results.append({ "filename": file.filename, **response_data })
            except Exception as e:
                all_results.append({ "filename": file.filename, "result": {"success": False, "error": str(e)}, "logs": f"Error: {str(e)}"})
    return jsonify({"bulk_results": all_results})

@app.route("/upload", methods=["POST"])
def upload_files():
    if 'files' not in request.files: return jsonify({"error": "No file part"}), 400
    files = request.files.getlist('files')
    if not files or files[0].filename == '': return jsonify({"error": "No selected file"}), 400
    session_id = str(uuid.uuid4())
    session_upload_folder = os.path.join(UPLOAD_FOLDER, session_id)
    os.makedirs(session_upload_folder)
    for file in files:
        if file:
            file.save(os.path.join(session_upload_folder, file.filename))
    return jsonify({"session_id": session_id})
# --- [UPDATED] ROUTE: PROCESSING PIPELINE ---

def encode_file_to_base64(file_path):
    """Helper to read a file and return its base64 string."""
    try:
        with open(file_path, "rb") as f:
            return base64.b64encode(f.read()).decode('utf-8')
    except Exception as e:
        print(f"Error encoding file {file_path}: {e}")
        return None

def find_output_file(output_folder, input_filename):
    """
    Tries to find the corresponding output DOCX file for a given input MD file.
    Assumes the pipeline maintains the base filename.
    """
    base_name = os.path.splitext(input_filename)[0]
    # Look for any .docx file that starts with the base name
    candidates = glob.glob(os.path.join(output_folder, f"*{base_name}*.docx"))
    
    if candidates:
        # Return the most likely match (shortest name usually implies exact match if prefixes exist)
        return max(candidates, key=os.path.getctime) # Return the most recently created match
    return None

@app.route("/run_pipeline", methods=["POST"])
def run_pipeline_single():
    if 'file' not in request.files: return jsonify({"error": "No file part"}), 400
    file = request.files['file']
    if file.filename == '': return jsonify({"error": "No selected file"}), 400

    # Create Session
    session_id = str(uuid.uuid4())
    session_upload_folder = os.path.join(UPLOAD_FOLDER, session_id)
    session_output_folder = os.path.join(OUTPUT_FOLDER, session_id)
    os.makedirs(session_upload_folder, exist_ok=True)
    os.makedirs(session_output_folder, exist_ok=True)

    try:
        # Save Input
        input_path = os.path.join(session_upload_folder, file.filename)
        file.save(input_path)

        # Run Pipeline
        print(f"🚀 Starting Pipeline for Single File Session: {session_id}")
        process_all_complete.run_pipeline(input_path=session_upload_folder, output_dir=session_output_folder)

        # Find Output
        output_file_path = find_output_file(session_output_folder, file.filename)
        
        if output_file_path and os.path.exists(output_file_path):
            docx_base64 = encode_file_to_base64(output_file_path)
            output_filename = os.path.basename(output_file_path)
            
            return jsonify({
                "filename": file.filename,
                "result": {
                    "success": True,
                    "docx_filename": output_filename,
                    "docx_base64": docx_base64
                },
                "logs": f"Successfully processed {file.filename} to {output_filename}"
            })
        else:
             return jsonify({
                "filename": file.filename,
                "result": { "success": False, "error": "Pipeline ran, but no matching DOCX output was found." },
                "logs": "Output file generation failed."
            })

    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({
            "filename": file.filename, 
            "result": { "success": False, "error": str(e) }, 
            "logs": f"Fatal Exception: {str(e)}"
        }), 500

@app.route("/run_pipeline/bulk", methods=["POST"])
def run_pipeline_bulk():
    if 'files' not in request.files: return jsonify({"error": "No file part"}), 400
    files = request.files.getlist('files')
    if not files: return jsonify({"error": "No files selected"}), 400

    # Create Session
    session_id = str(uuid.uuid4())
    session_upload_folder = os.path.join(UPLOAD_FOLDER, session_id)
    session_output_folder = os.path.join(OUTPUT_FOLDER, session_id)
    os.makedirs(session_upload_folder, exist_ok=True)
    os.makedirs(session_output_folder, exist_ok=True)

    saved_filenames = []

    try:
        # Save All Inputs
        for file in files:
            if file and file.filename:
                file.save(os.path.join(session_upload_folder, file.filename))
                saved_filenames.append(file.filename)

        # Run Pipeline (Runs on the whole folder)
        print(f"🚀 Starting Pipeline for Bulk Session: {session_id}")
        process_all_complete.run_pipeline(input_path=session_upload_folder, output_dir=session_output_folder)

        # Gather Results
        all_results = []
        for filename in saved_filenames:
            output_file_path = find_output_file(session_output_folder, filename)
            
            if output_file_path and os.path.exists(output_file_path):
                docx_base64 = encode_file_to_base64(output_file_path)
                output_filename = os.path.basename(output_file_path)
                all_results.append({
                    "filename": filename,
                    "result": {
                        "success": True,
                        "docx_filename": output_filename,
                        "docx_base64": docx_base64
                    },
                    "logs": "Success"
                })
            else:
                all_results.append({
                    "filename": filename,
                    "result": { "success": False, "error": "Output file not found" },
                    "logs": "Pipeline executed, but specific output file missing."
                })

        return jsonify({"bulk_results": all_results})

    except Exception as e:
        import traceback
        traceback.print_exc()
        # Return error structure for all files if the pipeline crashes entirely
        error_results = [{
            "filename": f, 
            "result": { "success": False, "error": str(e) }, 
            "logs": "Pipeline Crash"
        } for f in saved_filenames]
        return jsonify({"bulk_results": error_results})

# -------------------------------------

@app.route("/clean_sales/<session_id>", methods=["POST"])
def run_sales_processing(session_id):
    session_upload_folder = os.path.join(UPLOAD_FOLDER, session_id)
    session_output_folder = os.path.join(OUTPUT_FOLDER, session_id)
    if not os.path.exists(session_upload_folder): return jsonify({"error": "Invalid session ID"}), 400
    os.makedirs(session_output_folder, exist_ok=True)
    
    req_data = request.get_json() or {}
    group_units = req_data.get('groupUnits', False)
    
    if not SALES_DIRECTORY_FILE or not os.path.exists(SALES_DIRECTORY_FILE):
        print(f"❌ ERROR: Sales Directory File not found at: {SALES_DIRECTORY_FILE}")
        return jsonify({"error": f"Server configuration error: Sales Directory File not found."}), 500
        
    if not SALES_CURRENCY_FILE or not os.path.exists(SALES_CURRENCY_FILE):
        print(f"❌ ERROR: Sales Currency File not found at: {SALES_CURRENCY_FILE}")
        return jsonify({"error": f"Server configuration error: Sales Currency File not found."}), 500
        
    processing_results = process_files_to_csv(
        session_upload_folder, 
        session_output_folder, 
        SALES_DIRECTORY_FILE,
        SALES_CURRENCY_FILE
    )
    
    if not processing_results: return jsonify({"message": "Sales processing failed or no files were processed."}), 500

    if processing_results and group_units:
        print(f"Grouping flag is ON for session {session_id}. Running sales grouper...")
        try:
            processing_results = cleaning_configurations.group_sales_files(
                session_output_folder, 
                processing_results, 
                SALES_DIRECTORY_FILE
            )
        except Exception as e:
            print(f"❌ FATAL ERROR during grouping: {e}")
            return jsonify({
                "cleaned_files": processing_results, 
                "session_id": session_id, 
                "type": "sales",
                "logs": f"Grouping step failed: {e}. Returning individual files."
            })

    if processing_results and req_data.get('validateFormats', False):
        print(f"Hyperion validation flag is ON for session {session_id}. Running SALES validation...")
        if not SALES_3RD_HYPERION_FILE or not os.path.exists(SALES_3RD_HYPERION_FILE) or \
           not SALES_IC_HYPERION_FILE or not os.path.exists(SALES_IC_HYPERION_FILE) or \
           not SALES_DIRECTORY_FILE or not os.path.exists(SALES_DIRECTORY_FILE):
            print(f"   - ❌ SKIPPING validation: One or more required files are not set in .env or not found.")
        else:
            val_filename = "Sales_Hyperion_Validation_Output.xlsx"
            val_output_path = os.path.join(session_output_folder, val_filename)
            try:
                sheets_were_added = False
                with pd.ExcelWriter(val_output_path, engine='xlsxwriter') as writer:
                    for processed_file in list(processing_results):
                        if not processed_file.startswith("Sales_Data_Processed"): continue
                        processed_file_path = os.path.join(session_output_folder, processed_file)
                        validation_sheets_data = cleaning_configurations.generate_sales_validation_data(
                            processed_file_path, 
                            SALES_3RD_HYPERION_FILE,
                            SALES_IC_HYPERION_FILE,
                            SALES_DIRECTORY_FILE
                        )
                        for sheet_name, df_sheet in validation_sheets_data.items():
                            df_sheet.to_excel(writer, sheet_name=sheet_name, index=False)
                            sheets_were_added = True
                if sheets_were_added:
                    processing_results.append(val_filename)
                elif os.path.exists(val_output_path):
                    os.remove(val_output_path)
            except Exception as e:
                print(f"❌ FATAL ERROR during Sales Hyperion validation: {e}")
                import traceback
                traceback.print_exc() 
                return jsonify({
                    "cleaned_files": processing_results, 
                    "session_id": session_id, 
                    "type": "sales",
                    "logs": f"Sales Hyperion Validation step failed: {e}. Returning partial files."
                })

    return jsonify({"cleaned_files": processing_results, "session_id": session_id, "type": "sales"})

@app.route("/clean_oe/<session_id>", methods=["POST"])
def run_oe_processing(session_id):
    session_upload_folder = os.path.join(UPLOAD_FOLDER, session_id)
    session_output_folder = os.path.join(OUTPUT_FOLDER, session_id)
    if not os.path.exists(session_upload_folder): return jsonify({"error": "Invalid session ID"}), 400
    os.makedirs(session_output_folder, exist_ok=True)

    req_data = request.get_json() or {}
    group_units = req_data.get('groupUnits', False)

    if not HYPERION_FOLDER or not os.path.exists(HYPERION_FOLDER):
        print(f"❌ ERROR: Hyperion folder not found at: {HYPERION_FOLDER}")
        return jsonify({"error": f"Server configuration error: Hyperion folder not found."}), 500
        
    processing_results = process_excel_files(
        session_upload_folder, 
        session_output_folder, 
        HYPERION_FOLDER, 
        SALES_DIRECTORY_FILE, 
        SALES_CURRENCY_FILE,
        group_units=group_units
    )
    
    if not processing_results: return jsonify({"message": "OE processing failed or no Excel files were processed."}), 500
    
    if processing_results and group_units:
        print(f"Grouping flag is ON for session {session_id}. Running OE grouper...")
        try:
            processing_results = cleaning_configurations.group_oe_files(
                session_output_folder, 
                processing_results, 
                SALES_DIRECTORY_FILE,
                HYPERION_FOLDER
            )
        except Exception as e:
            print(f"❌ FATAL ERROR during OE grouping: {e}")
            import traceback
            traceback.print_exc()
            return jsonify({
                "cleaned_files": processing_results, 
                "session_id": session_id, 
                "type": "oe",
                "logs": f"Grouping step failed: {e}. Returning individual files."
            })

    if processing_results and req_data.get('validateFormats', False):
        print(f"Hyperion validation flag is ON for session {session_id}. Running OE validation...")
        if not HYPERION_FOLDER or not os.path.exists(HYPERION_FOLDER) or \
           not SALES_DIRECTORY_FILE or not os.path.exists(SALES_DIRECTORY_FILE):
            print(f"   - ❌ SKIPPING: Config files missing.")
        else:
            val_filename = "OE_Hyperion_Validation_Output.xlsx"
            val_output_path = os.path.join(session_output_folder, val_filename)
            try:
                sheets_were_added = False
                with pd.ExcelWriter(val_output_path, engine='xlsxwriter') as writer:
                    for processed_file in list(processing_results):
                        if not processed_file.startswith("OE_Data_Processed"): continue
                        processed_file_path = os.path.join(session_output_folder, processed_file)
                        validation_sheets_data = cleaning_configurations.generate_oe_validation_data(
                            processed_file_path,
                            HYPERION_FOLDER,
                            SALES_DIRECTORY_FILE
                        )
                        for sheet_name, df_sheet in validation_sheets_data.items():
                            df_sheet.to_excel(writer, sheet_name=sheet_name, index=False)
                            sheets_were_added = True
                if sheets_were_added:
                    processing_results.append(val_filename)
                elif os.path.exists(val_output_path):
                    os.remove(val_output_path)
            except Exception as e:
                print(f"❌ FATAL ERROR during OE Hyperion validation: {e}")
                import traceback
                traceback.print_exc() 
                return jsonify({
                    "cleaned_files": processing_results, 
                    "session_id": session_id, 
                    "type": "oe",
                    "logs": f"OE Hyperion Validation step failed: {e}. Returning partial files."
                })

    return jsonify({"cleaned_files": processing_results, "session_id": session_id, "type": "oe"})

@app.route("/clean_wc/<session_id>", methods=["POST"])
def run_wc_processing(session_id):
    session_upload_folder = os.path.join(UPLOAD_FOLDER, session_id)
    session_output_folder = os.path.join(OUTPUT_FOLDER, session_id)
    if not os.path.exists(session_upload_folder): return jsonify({"error": "Invalid session ID"}), 400
    os.makedirs(session_output_folder, exist_ok=True)
    req_data = request.get_json()
    if not req_data or 'metric' not in req_data: return jsonify({"error": "Metric ('dso' or 'overhead') not specified."}), 400
    metric = req_data['metric']
    try:
        processing_results = process_working_capital(session_upload_folder, session_output_folder, metric)
        if not processing_results: return jsonify({"message": f"Working Capital ({metric}) processing failed."}), 500
        return jsonify({"cleaned_files": processing_results, "session_id": session_id, "type": metric})
    except (FileNotFoundError, ValueError) as e:
        return jsonify({"error": str(e)}), 400
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({"error": "An unexpected error occurred."}), 500

@app.route("/clean_pex/<session_id>", methods=["POST"])
def run_pex_processing(session_id):
    session_upload_folder = os.path.join(UPLOAD_FOLDER, session_id)
    session_output_folder = os.path.join(OUTPUT_FOLDER, session_id)

    if not os.path.exists(session_upload_folder): return jsonify({"error": "Invalid session ID"}), 400
    os.makedirs(session_output_folder, exist_ok=True)
    
    req_data = request.get_json()
    if not req_data or 'sub_metric' not in req_data: return jsonify({"error": "PEX 'sub_metric' not specified."}), 400
    
    sub_metric = req_data['sub_metric']
    bulk_mode = req_data.get('bulk_mode', False)
    group_units = req_data.get('groupUnits', False)
    analysis_type = req_data.get('vendorAnalysisType', 'mom')
    
    try:
        processing_results = []
        if sub_metric == 'pex-bi':
            if not PEX_LOOKUP_FOLDER or not os.path.exists(PEX_LOOKUP_FOLDER):
                print(f"❌ ERROR: PEX Lookup folder not found at: {PEX_LOOKUP_FOLDER}")
                return jsonify({"error": f"Server configuration error: PEX Lookup folder not found."}), 500
            processing_results = process_pex_and_headcount(session_upload_folder, session_output_folder, PEX_LOOKUP_FOLDER, SALES_DIRECTORY_FILE, SALES_CURRENCY_FILE)
        
        elif sub_metric == 'pex-vendor':
            processing_results = process_pex_vendor(
                session_upload_folder, 
                session_output_folder, 
                bulk_mode, 
                SALES_DIRECTORY_FILE, 
                SALES_CURRENCY_FILE,
                analysis_type
            )
            if isinstance(processing_results, str): processing_results = [processing_results]
        else:
            return jsonify({"error": f"Unknown PEX sub-metric: '{sub_metric}'"}), 400
            
        if not processing_results: return jsonify({"message": f"PEX processing ({sub_metric}) failed."}), 500
        
        if processing_results and group_units:
            print(f"Grouping flag is ON for session {session_id}. Running PEX grouper...")
            try:
                if sub_metric == 'pex-bi':
                    processing_results = cleaning_configurations.group_pex_bi_and_headcount_files(session_output_folder, processing_results, SALES_DIRECTORY_FILE)
                elif sub_metric == 'pex-vendor':
                    processing_results = cleaning_configurations.group_pex_vendor_files(session_output_folder, processing_results, SALES_DIRECTORY_FILE)
            except Exception as e:
                print(f"❌ FATAL ERROR during PEX grouping: {e}")
                return jsonify({"cleaned_files": processing_results, "session_id": session_id, "type": sub_metric, "logs": f"Grouping step failed: {e}."})

        if processing_results and req_data.get('validateFormats', False):
            print(f"Hyperion validation flag is ON for session {session_id}. Running PEX validation...")
            if sub_metric == 'pex-bi':
                try:
                    if not PEX_HYPERION_FILE or not os.path.exists(PEX_HYPERION_FILE) or \
                       not SALES_DIRECTORY_FILE or not os.path.exists(SALES_DIRECTORY_FILE):
                        print(f"   - ❌ SKIPPING: Config files missing.")
                    else:
                        val_filename = "PEX_Hyperion_Validation_Output.xlsx"
                        val_output_path = os.path.join(session_output_folder, val_filename)
                        sheets_were_added = False
                        with pd.ExcelWriter(val_output_path, engine='xlsxwriter') as writer:
                            for processed_file in list(processing_results):
                                if not processed_file.startswith("PEX_Data_Processed"): continue
                                processed_file_path = os.path.join(session_output_folder, processed_file)
                                validation_sheets_data = cleaning_configurations.generate_pex_validation_data(
                                    processed_file_path,
                                    PEX_HYPERION_FILE,
                                    SALES_DIRECTORY_FILE
                                )
                                for sheet_name, df_sheet in validation_sheets_data.items():
                                    df_sheet.to_excel(writer, sheet_name=sheet_name, index=False)
                                    sheets_were_added = True
                        if sheets_were_added:
                            processing_results.append(val_filename)
                        elif os.path.exists(val_output_path):
                            os.remove(val_output_path)
                except Exception as e:
                    print(f"❌ FATAL ERROR during PEX Hyperion validation: {e}")
                    return jsonify({"cleaned_files": processing_results, "session_id": session_id, "type": sub_metric, "logs": f"PEX Validation failed: {e}."})
        
        return jsonify({"cleaned_files": processing_results, "session_id": session_id, "type": sub_metric})

    except (FileNotFoundError, ValueError) as e:
        return jsonify({"error": str(e)}), 400
    except Exception as e:
        print("="*50)
        print(f"❌ A FATAL unhandled exception occurred in /clean_pex:")
        import traceback
        traceback.print_exc()
        print("="*50)
        return jsonify({"error": f"An unexpected server error occurred: {str(e)}"}), 500

@app.route("/remove_duplicates/<session_id>", methods=["POST"])
def run_duplicate_removal(session_id):
    session_output_folder = os.path.join(OUTPUT_FOLDER, session_id)
    if not os.path.exists(session_output_folder): return jsonify({"error": "Invalid session ID"}), 400
    try:
        current_files = [f for f in os.listdir(session_output_folder) if f.endswith(('.csv', '.xlsx'))]
        if not current_files: return jsonify({"message": "No files found."}), 200
        kept_files = cleaning_configurations.remove_duplicate_files(session_output_folder, current_files)
        return jsonify({"message": "Duplicate removal complete.", "cleaned_files": kept_files, "session_id": session_id, "type": "duplicates_removed"})
    except Exception as e:
        print(f"❌ FATAL ERROR during duplicate removal: {e}")
        return jsonify({"error": "Error during duplicate removal."}), 500

@app.route("/preview/<session_id>", methods=["GET"])
def get_data_preview(session_id):
    preview_type = request.args.get('type', 'raw')
    processing_type = request.args.get('processing_type', 'oe') 
    source_folder = os.path.join(UPLOAD_FOLDER if preview_type == 'raw' else OUTPUT_FOLDER, session_id)
    if not os.path.exists(source_folder): return jsonify({"error": "Session ID not found."}), 404
    files = glob.glob(os.path.join(source_folder, "*.xlsx")) + glob.glob(os.path.join(source_folder, "*.csv"))
    if not files: return jsonify({"error": "No files found."}), 404
    file_path, filename = files[0], os.path.basename(files[0])
    try:
        if file_path.endswith('.csv'):
            df_preview = pd.read_csv(file_path, nrows=5)
            with open(file_path, 'r', encoding='utf-8') as f: total_rows = sum(1 for line in f) - 1
        elif file_path.endswith('.xlsx'):
            sheet_name = 'Raw' if processing_type == 'sales' else 0
            try:
                df_preview = pd.read_excel(file_path, nrows=5, sheet_name=sheet_name)
                total_rows = len(pd.read_excel(file_path, sheet_name=sheet_name))
            except Exception:
                sheet_name = 0
                df_preview = pd.read_excel(file_path, nrows=5, sheet_name=sheet_name)
                total_rows = len(pd.read_excel(file_path, sheet_name=sheet_name))
        else:
            return jsonify({"error": "Unsupported file."}), 400
        df_preview = df_preview.replace([np.nan, np.inf, -np.inf], None).where(pd.notna(df_preview), None)
        preview_data = df_preview.to_dict('records')
        def clean_value(v):
            if pd.isna(v): return None
            return v.item() if isinstance(v, (np.integer, np.floating)) else v
        cleaned_preview_data = [{k: clean_value(v) for k, v in row.items()} for row in preview_data]
        return jsonify({ "filename": filename, "data": cleaned_preview_data, "total_rows": int(total_rows), "preview_type": preview_type, "processing_type": processing_type, "sheet_name": str(sheet_name) })
    except Exception as e:
        return jsonify({"error": f"Preview error: {str(e)}"}), 500

@app.route("/download/<session_id>/<string:filename>", methods=["GET"])
def download_file(session_id, filename):
    session_output_folder = os.path.join(OUTPUT_FOLDER, session_id)
    try:
        return send_from_directory(directory=session_output_folder, path=filename, as_attachment=True)
    except FileNotFoundError:
        return jsonify({"error": "File not found!"}), 404

@app.route("/download/zip/<session_id>", methods=["GET"])
def download_zip(session_id):
    session_output_folder = os.path.join(OUTPUT_FOLDER, session_id)
    if not os.path.exists(session_output_folder) or not os.listdir(session_output_folder):
        return jsonify({"error": "No processed files found."}), 404
    memory_file = io.BytesIO()
    with zipfile.ZipFile(memory_file, 'w', zipfile.ZIP_DEFLATED) as zf:
        for file in os.listdir(session_output_folder):
            zf.write(os.path.join(session_output_folder, file), file)
    memory_file.seek(0)
    return app.response_class( memory_file.read(), mimetype='application/zip', headers={'Content-Disposition': f'attachment; filename=processed_data_{session_id}.zip'} )

if __name__ == "__main__":
    # use_reloader=False stops the server from restarting when files are created
    print("🚀 Starting Server on Port 8080...")
    app.run(debug=True, use_reloader=False, port=8080)