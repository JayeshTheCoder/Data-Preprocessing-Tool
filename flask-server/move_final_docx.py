"""
Move Final DOCX Files (Dynamic - Fixed)
"""

import os
import sys
import shutil
from datetime import datetime

def generate_destination_filename(filename, subfolder_name=""):
    # Clean up filename
    base_name = os.path.splitext(filename)[0]
    base_name = base_name.replace('_C1_Analysis', '').replace('_MOM', '')
    
    current_year = datetime.now().year
    new_filename = f"{base_name}_Commentary_{current_year}.docx"
    
    return new_filename

def process_input(input_path, output_dir, dry_run=False, backup=True):
    if not os.path.exists(output_dir) and not dry_run:
        os.makedirs(output_dir)

    files_to_process = []
    
    if os.path.isfile(input_path):
        if input_path.endswith('.docx'):
            files_to_process.append(input_path)
    elif os.path.isdir(input_path):
        for root, dirs, files in os.walk(input_path):
            for file in files:
                # [FIX] Removed check for 'C1_Analysis' so ALL docx files are moved
                if file.endswith('.docx'):
                    files_to_process.append(os.path.join(root, file))

    print(f"Found {len(files_to_process)} file(s) to move/copy.")
    
    results = []
    for source_path in files_to_process:
        filename = os.path.basename(source_path)
        
        # Don't try to move files that are already in the output directory
        if os.path.abspath(os.path.dirname(source_path)) == os.path.abspath(output_dir):
            continue

        parent_folder = os.path.basename(os.path.dirname(source_path))
        
        new_filename = generate_destination_filename(filename, parent_folder)
        dest_path = os.path.join(output_dir, new_filename)
        
        if dry_run:
            print(f"[DRY RUN] Would copy {filename} -> {dest_path}")
            results.append({'file': filename, 'status': 'dry_run'})
        else:
            try:
                if backup:
                    shutil.copy2(source_path, dest_path)
                    action = "Copied"
                else:
                    shutil.move(source_path, dest_path)
                    action = "Moved"
                
                print(f"{action}: {new_filename}")
                results.append({'file': new_filename, 'status': 'success'})
            except Exception as e:
                print(f"Error processing {filename}: {e}")
                results.append({'file': filename, 'status': 'error'})

    return results

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("Usage: python move_final_docx.py <input_path> <output_dir> [--dry-run]")
        sys.exit(1)
        
    input_p = sys.argv[1]
    output_d = sys.argv[2]
    dry_run = '--dry-run' in sys.argv
    process_input(input_p, output_d, dry_run=dry_run)