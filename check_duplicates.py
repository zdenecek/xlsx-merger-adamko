#!/usr/bin/env python3
"""
Script to check for duplicate values in column B for each Excel file in a folder.
"""

import os
import pandas as pd
from pathlib import Path
import argparse


def check_duplicates_in_column_b(file_path):
    """
    Check for duplicate values in column B of an Excel file.
    
    Args:
        file_path (str): Path to the Excel file
        
    Returns:
        tuple: (has_duplicates, duplicate_values, total_rows)
    """
    try:
        # Read the Excel file
        df = pd.read_excel(file_path)
        
        # Check if column B exists (index 1)
        if df.shape[1] < 2:
            return None, f"File has less than 2 columns", 0
        
        # Get column B (index 1)
        column_b = df.iloc[:, 1]
        
        # Remove NaN values for duplicate checking
        column_b_clean = column_b.dropna()
        
        # Find duplicates
        duplicates = column_b_clean[column_b_clean.duplicated(keep=False)]
        
        if len(duplicates) > 0:
            # Get unique duplicate values
            unique_duplicates = duplicates.unique()
            return True, unique_duplicates.tolist(), len(column_b_clean)
        else:
            return False, [], len(column_b_clean)
            
    except Exception as e:
        return None, f"Error reading file: {str(e)}", 0


def check_folder_for_duplicates(folder_path):
    """
    Check all Excel files in a folder for duplicates in column B.
    
    Args:
        folder_path (str): Path to the folder containing Excel files
    """
    folder = Path(folder_path)
    
    if not folder.exists():
        print(f"Error: Folder '{folder_path}' does not exist.")
        return
    
    # Find all Excel files
    xlsx_files = list(folder.glob("*.xlsx")) + list(folder.glob("*.xls"))
    
    if not xlsx_files:
        print(f"No Excel files found in '{folder_path}'")
        return
    
    print(f"Checking {len(xlsx_files)} Excel files in '{folder_path}'...\n")
    
    files_with_duplicates = 0
    total_files_checked = 0
    
    for file_path in sorted(xlsx_files):
        print(f"Checking: {file_path.name}")
        
        has_duplicates, duplicate_info, total_rows = check_duplicates_in_column_b(file_path)
        
        if has_duplicates is None:
            print(f"  ❌ {duplicate_info}")
        elif has_duplicates:
            files_with_duplicates += 1
            print(f"  ⚠️  DUPLICATES FOUND!")
            print(f"     Total rows in column B: {total_rows}")
            print(f"     Duplicate values: {duplicate_info}")
        else:
            print(f"  ✅ No duplicates found (checked {total_rows} rows)")
        
        total_files_checked += 1
        print()
    
    # Summary
    print("="*50)
    print("SUMMARY:")
    print(f"Total files checked: {total_files_checked}")
    print(f"Files with duplicates: {files_with_duplicates}")
    print(f"Files without duplicates: {total_files_checked - files_with_duplicates}")


def main():
    parser = argparse.ArgumentParser(description="Check for duplicate values in column B of Excel files")
    parser.add_argument("folder", nargs="?", default=".", 
                       help="Path to folder containing Excel files (default: current directory)")
    
    args = parser.parse_args()
    
    print("Excel Duplicate Checker - Column B")
    print("="*40)
    
    check_folder_for_duplicates(args.folder)


if __name__ == "__main__":
    main() 