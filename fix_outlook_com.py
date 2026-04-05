"""Rebuild pywin32 COM cache for Outlook"""
import shutil
import os
from pathlib import Path

print("Rebuilding pywin32 COM cache...")
print("=" * 60)

# Find and clear the gen_py cache
try:
    import win32com
    gen_py_path = Path(win32com.__gen_path__)
    
    if gen_py_path.exists():
        print(f"Found cache directory: {gen_py_path}")
        print("Clearing cache...")
        
        # Delete all files in gen_py
        for item in gen_py_path.iterdir():
            try:
                if item.is_file():
                    item.unlink()
                elif item.is_dir():
                    shutil.rmtree(item)
                print(f"  Deleted: {item.name}")
            except Exception as e:
                print(f"  Warning: Could not delete {item.name}: {e}")
        
        print("\n✓ Cache cleared successfully!")
    else:
        print("Cache directory not found (this is OK)")
    
except Exception as e:
    print(f"Error: {e}")

print("\nNow try running the script again with: python main.py")
print("=" * 60)
