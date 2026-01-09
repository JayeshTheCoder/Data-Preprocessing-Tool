import os
import sys

# 1. Check Imports
print("1. Checking libraries...")
try:
    import markdown
    import docx
    import bs4
    import requests
    print("   ✅ Libraries found.")
except ImportError as e:
    print(f"   ❌ MISSING LIBRARY: {e}")
    print("   Run: pip install markdown python-docx beautifulsoup4 requests")
    sys.exit(1)

# 2. Check Pipeline Script
print("2. Checking pipeline script...")
try:
    from process_all_complete import run_pipeline
    print("   ✅ Pipeline script imported successfully.")
except ImportError as e:
    print(f"   ❌ FAILED to import pipeline: {e}")
    sys.exit(1)

# 3. Create Dummy Data
print("3. Creating test data...")
os.makedirs("test_input", exist_ok=True)
os.makedirs("test_output", exist_ok=True)
with open("test_input/test.md", "w") as f:
    f.write("# Test Document\n\nThis is a test.")

# 4. Run Pipeline
print("4. Running pipeline...")
try:
    # Call the function exactly how server.py calls it
    run_pipeline(input_path="test_input", output_dir="test_output")
    
    if os.path.exists("test_output/test.docx"):
        print("\n✅ SUCCESS: test.docx was created.")
    else:
        print("\n❌ FAILED: Pipeline ran but no file was created.")
except Exception as e:
    print(f"\n❌ CRASHED: {e}")
    import traceback
    traceback.print_exc()