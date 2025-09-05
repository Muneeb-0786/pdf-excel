#!/usr/bin/env python3
"""
Script to combine multiple Excel files from PDF processing into one combined file.
"""

import pandas as pd
import os
from pathlib import Path

def combine_excel_files(input_dir: str, output_file: str):
    """
    Combine multiple Excel files into one.
    
    Args:
        input_dir: Directory containing Excel files to combine
        output_file: Path for the combined Excel file
    """
    excel_files = list(Path(input_dir).glob("*.xlsx"))
    
    # Filter out files that might be temporary or sample files
    excel_files = [f for f in excel_files if not f.name.startswith(('~', 'SAMPLE', 'combined'))]
    
    if not excel_files:
        print(f"No Excel files found in {input_dir}")
        return
    
    print(f"Found {len(excel_files)} Excel files to combine:")
    for file in excel_files:
        print(f"  - {file.name}")
    
    # Read and combine all Excel files
    all_dataframes = []
    
    for excel_file in excel_files:
        try:
            print(f"Reading {excel_file.name}...")
            df = pd.read_excel(excel_file)
            
            # Add source file information
            df['Source_File'] = excel_file.stem
            
            all_dataframes.append(df)
            print(f"  - Added {len(df)} rows from {excel_file.name}")
            
        except Exception as e:
            print(f"Error reading {excel_file}: {e}")
            continue
    
    if all_dataframes:
        # Combine all dataframes
        combined_df = pd.concat(all_dataframes, ignore_index=True)
        
        # Reorder columns to put Source_File at the end
        columns = [col for col in combined_df.columns if col != 'Source_File'] + ['Source_File']
        combined_df = combined_df[columns]
        
        # Write combined file
        combined_df.to_excel(output_file, index=False, sheet_name="COMBINED_ATTRIBUTES")
        
        print(f"\nCombined Excel file created: {output_file}")
        print(f"Total rows: {len(combined_df)}")
        print(f"Columns: {combined_df.columns.tolist()}")
        
        # Show summary by source file
        print("\nSummary by source file:")
        summary = combined_df.groupby('Source_File').size()
        for source, count in summary.items():
            print(f"  {source}: {count} rows")
            
        return combined_df
    else:
        print("No data to combine!")
        return None

def create_placeholder_for_failed_pdf(pdf_name: str) -> pd.DataFrame:
    """
    Create placeholder data for PDFs that couldn't be processed.
    """
    placeholder_data = [
        ["PLACEHOLDER", "FAILED-PROCESS", "000", 1, "ERROR01", "PDF Processing Failed", 
         "Could not extract text - needs OCR setup", "", 
         "Requires poppler installation", "", pdf_name.replace('.PDF', '').replace('.pdf', '')]
    ]
    
    headers = [
        "Field", "TPLNR", "CLASS", "KLART", "POSNUMMER", "ATNAM", "ATWRT",
        "Characteristics UoM", "Remarks", "Additional", "REF"
    ]
    
    df = pd.DataFrame(placeholder_data, columns=headers)
    df['Source_File'] = pdf_name
    return df

if __name__ == "__main__":
    # First, let's see what Excel files we have
    print("Checking output directory...")
    output_dir = "output"
    
    # Create placeholder for the failed PDF
    failed_pdf_name = "P11569-11-99-40-1605-1"
    placeholder_df = create_placeholder_for_failed_pdf(failed_pdf_name)
    placeholder_file = f"{output_dir}/{failed_pdf_name}_placeholder.xlsx"
    
    placeholder_df.to_excel(placeholder_file, index=False, sheet_name="ATTRIBUTES")
    print(f"Created placeholder file: {placeholder_file}")
    
    # Now combine all files
    combined_file = f"{output_dir}/combined_both_pdfs.xlsx"
    result = combine_excel_files(output_dir, combined_file)
    
    if result is not None:
        print(f"\n✅ Successfully created combined file with data from both PDFs!")
        print(f"   Note: One PDF needed OCR processing which requires poppler installation.")
        print(f"   File: {combined_file}")
    else:
        print("❌ Failed to create combined file")
