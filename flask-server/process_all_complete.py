"""
Complete MD to DOCX Processing Pipeline

This is the main orchestration script that runs the entire workflow:
1. Analyze current status
2. Convert MD files to DOCX
3. Apply currency conversions
4. Move final files to destination
5. Generate final report

Usage:
    python Script/process_all_complete.py              # Run full pipeline
    python Script/process_all_complete.py --dry-run    # Preview what would happen
    python Script/process_all_complete.py --analyze    # Only run analysis
"""
"""
Dynamic MD to DOCX Pipeline Orchestrator
"""

import sys
import os
import argparse
from datetime import datetime

# Import modified modules
import convert_all_folders
import currency_converter_enhanced
import add_header_to_docx
import move_final_docx
import analyze_md_status

def print_step_header(step_num, title):
    print(f"\n{'='*60}")
    print(f"STEP {step_num}: {title}")
    print(f"{'='*60}")

def run_pipeline(input_path, output_dir=None, dry_run=False):
    start_time = datetime.now()
    
    print(f"\n🚀 STARTING PIPELINE")
    print(f"   Target: {input_path}")
    print(f"   Mode:   {'DRY RUN' if dry_run else 'LIVE ACTION'}")

    # --- Step 1: Analysis ---
    print_step_header(1, "Analyzing Current Status")
    # This script prints its own details
    analyze_md_status.analyze_path(input_path)

    # --- Step 2: Convert ---
    print_step_header(2, "Converting MD to DOCX")
    # This script prints its own details
    convert_all_folders.process_input(input_path, dry_run=dry_run)

    # --- Step 3: Add Headers ---
    print_step_header(3, "Adding Headers to DOCX")
    add_header_to_docx.process_input(input_path, dry_run=dry_run)

    # --- Step 4: Currency Conversion ---
    print_step_header(4, "Applying Currency Conversions")
    currency_converter_enhanced.process_input(input_path, dry_run=dry_run)

    # --- Step 5: Move (Optional) ---
    print_step_header(5, "Final File Organization")
    if output_dir:
        print(f"Moving files to: {output_dir}")
        move_final_docx.process_input(input_path, output_dir, dry_run=dry_run)
    else:
        print("Skipped (No --output directory provided).")
        print("Files remain in their source folders.")

    # --- Final Summary ---
    duration = datetime.now() - start_time
    print(f"\n{'='*60}")
    print(f"✅ PIPELINE COMPLETE")
    print(f"   Time taken: {duration.total_seconds():.1f} seconds")
    print(f"{'='*60}\n")

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Dynamic MD to DOCX Pipeline")
    parser.add_argument("input_path", help="Path to input file or folder")
    parser.add_argument("--output", "-o", help="Optional output folder for final files")
    parser.add_argument("--dry-run", "-d", action="store_true", help="Preview changes without executing")
    
    args = parser.parse_args()
    
    if not os.path.exists(args.input_path):
        print(f"Error: Path not found: {args.input_path}")
        sys.exit(1)
        
    run_pipeline(args.input_path, args.output, args.dry_run)