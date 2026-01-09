"""
Batch MD to DOCX Converter (Dynamic)
"""

import os
import sys
from convert_clean import convert_md_to_docx

def process_input(input_path, dry_run=False):
    """
    Process a specific file or folder (recursively).
    """
    files_to_process = []
    
    # Determine files to process
    if os.path.isfile(input_path):
        if input_path.endswith('.md'):
            files_to_process.append(input_path)
    elif os.path.isdir(input_path):
        for root, dirs, files in os.walk(input_path):
            for file in files:
                if file.endswith('.md'):
                    files_to_process.append(os.path.join(root, file))
    
    print(f"Found {len(files_to_process)} MD file(s) to process.")

    results = []
    
    for md_file_path in files_to_process:
        docx_file_path = md_file_path.replace('.md', '.docx')
        
        if os.path.exists(docx_file_path):
            print(f"Skipped (DOCX exists): {os.path.basename(md_file_path)}")
            continue

        if dry_run:
            print(f"[DRY RUN] Would convert: {os.path.basename(md_file_path)}")
            results.append({'file': md_file_path, 'status': 'dry_run'})
        else:
            result = convert_md_to_docx(md_file_path)
            print(result)
            status = 'success' if "Successfully" in result else 'failed'
            results.append({'file': md_file_path, 'status': status, 'msg': result})

    return {'processed': len(files_to_process), 'results': results}

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python convert_all_folders.py <input_path> [--dry-run]")
        sys.exit(1)
    
    path = sys.argv[1]
    dry_run = '--dry-run' in sys.argv
    process_input(path, dry_run=dry_run)