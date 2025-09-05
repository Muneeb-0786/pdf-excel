#!/usr/bin/env python3
"""
Script to examine the improved Excel output.
"""

import pandas as pd

def examine_excel_output(excel_path):
    """Examine the generated Excel file."""
    print(f"Examining Excel file: {excel_path}")
    print("="*50)
    
    df = pd.read_excel(excel_path)
    
    print(f"Shape: {df.shape}")
    print(f"Columns: {list(df.columns)}")
    
    print("\nFirst 10 rows:")
    print("-"*50)
    print(df.head(10).to_string())
    
    # Look for document references
    doc_pattern = r'\d+\.\d+\.\d+\.\d+'
    doc_rows = df[df['Item Name'].str.contains(doc_pattern, na=False)]
    
    print(f"\nRows with document references: {len(doc_rows)}")
    if len(doc_rows) > 0:
        print("-"*50)
        print(doc_rows.head(10).to_string())
    
    # Look for instrument/equipment data
    equipment_keywords = ['DETECTOR', 'INSTRUMENT', 'HOUSING', 'TERMINALS', 'SCREW']
    equipment_rows = df[df['Description'].str.contains('|'.join(equipment_keywords), case=False, na=False)]
    
    print(f"\nRows with equipment/instrument data: {len(equipment_rows)}")
    if len(equipment_rows) > 0:
        print("-"*50)
        print(equipment_rows.head(10).to_string())
    
    # Data quality analysis
    print(f"\nData Quality Analysis:")
    print("-"*30)
    print(f"Non-empty Item Names: {df['Item Name'].notna().sum()}")
    print(f"Non-empty Descriptions: {df['Description'].notna().sum()}")
    print(f"Non-empty Serial Numbers: {df['Serial Number'].notna().sum()}")
    print(f"Non-empty Dates: {df['Date'].notna().sum()}")
    print(f"Non-empty Quantities: {df['Quantity'].notna().sum()}")

if __name__ == "__main__":
    examine_excel_output("output/improved_output.xlsx")
