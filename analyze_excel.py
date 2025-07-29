import pandas as pd
import sys
from pathlib import Path

def analyze_excel_file(file_path):
    """Analyze Excel file structure and content"""
    try:
        # Read all sheets from the Excel file
        excel_file = pd.ExcelFile(file_path)
        print(f"Excel file: {file_path}")
        print(f"Sheet names: {excel_file.sheet_names}")
        print(f"Number of sheets: {len(excel_file.sheet_names)}")
        print("-" * 50)
        
        # Analyze each sheet
        for sheet_name in excel_file.sheet_names:
            print(f"\nSheet: {sheet_name}")
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            
            print(f"Shape: {df.shape} (rows, columns)")
            print(f"Columns: {list(df.columns)}")
            print(f"Data types:")
            for col in df.columns:
                print(f"  {col}: {df[col].dtype}")
            
            print(f"\nFirst 5 rows:")
            print(df.head())
            
            print(f"\nSummary statistics:")
            print(df.describe())
            
            print(f"\nMissing values:")
            missing = df.isnull().sum()
            if missing.sum() > 0:
                print(missing[missing > 0])
            else:
                print("No missing values")
            
            print("-" * 50)
            
    except Exception as e:
        print(f"Error analyzing file: {e}")
        return None

if __name__ == "__main__":
    # Look for Excel files in current directory
    excel_files = list(Path(".").glob("*.xlsx"))
    
    if not excel_files:
        print("No Excel files found in current directory")
        sys.exit(1)
    
    # Analyze the first Excel file found
    file_path = excel_files[0]
    print(f"Analyzing: {file_path}")
    analyze_excel_file(file_path) 