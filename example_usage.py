#!/usr/bin/env python3
"""
Example script showing how to use the PDF to Excel processor.
"""

from pdf_to_excel import PDFProcessor
import os

def main():
    """Example usage of PDFProcessor."""
    processor = PDFProcessor()
    
    # Example 1: Process a single PDF file
    print("Example 1: Single PDF processing")
    try:
        pdf_file = "sample_data/example.pdf"  # Replace with your PDF
        excel_file = "output/example.xlsx"
        
        if os.path.exists(pdf_file):
            rows = processor.process_pdf_to_excel(pdf_file, excel_file)
            print(f"Successfully processed {pdf_file} -> {excel_file}")
            print(f"Rows processed: {rows}")
        else:
            print(f"PDF file not found: {pdf_file}")
    except Exception as e:
        print(f"Error processing single PDF: {e}")
    
    print("\n" + "="*50 + "\n")
    
    # Example 2: Batch process multiple PDFs
    print("Example 2: Batch processing")
    try:
        input_dir = "sample_data"
        output_dir = "output"
        
        results = processor.batch_process(input_dir, output_dir)
        print("Batch processing results:")
        for result in results:
            if result["status"] == "success":
                print(f"✓ {result['file']}: {result['rows']} rows")
            else:
                print(f"✗ {result['file']}: {result['error']}")
    except Exception as e:
        print(f"Error in batch processing: {e}")
    
    print("\n" + "="*50 + "\n")
    
    # Example 3: Force OCR processing
    print("Example 3: OCR processing (for scanned PDFs)")
    try:
        pdf_file = "sample_data/scanned.pdf"  # Replace with scanned PDF
        excel_file = "output/scanned_ocr.xlsx"
        
        if os.path.exists(pdf_file):
            rows = processor.process_pdf_to_excel(pdf_file, excel_file, use_ocr=True)
            print(f"OCR processed {pdf_file} -> {excel_file}")
            print(f"Rows processed: {rows}")
        else:
            print(f"Scanned PDF file not found: {pdf_file}")
    except Exception as e:
        print(f"Error in OCR processing: {e}")

if __name__ == "__main__":
    main()
