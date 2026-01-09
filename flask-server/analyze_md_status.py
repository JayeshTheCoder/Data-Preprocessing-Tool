"""
Status Analyzer (Dynamic)
"""

import os
import sys

def analyze_path(input_path):
    """
    Analyze a specific path (file or folder) for conversion status.
    """
    md_files = []
    if os.path.isfile(input_path):
        if input_path.endswith('.md'):
            md_files.append(input_path)
    elif os.path.isdir(input_path):
        for root, dirs, files in os.walk(input_path):
            for file in files:
                if file.endswith('.md'):
                    md_files.append(os.path.join(root, file))
    
    stats = {'total': len(md_files), 'converted': 0, 'pending': 0}
    
    print(f"Analysis for: {input_path}")
    print("-" * 60)
    
    for md_file in md_files:
        docx_path = md_file.replace('.md', '.docx')
        has_docx = os.path.exists(docx_path)
        
        if has_docx:
            status = "✓ Converted"
            stats['converted'] += 1
        else:
            status = "⚠ Pending"
            stats['pending'] += 1
        
        print(f"  {status}: {os.path.basename(md_file)}")
        
    print("-" * 60)
    print(f"Summary: {stats['total']} files found | {stats['converted']} converted | {stats['pending']} pending.")
    return stats

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python analyze_md_status.py <input_path>")
        sys.exit(1)
    analyze_path(sys.argv[1])