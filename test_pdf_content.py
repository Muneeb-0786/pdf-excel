#!/usr/bin/env python3
"""
Test script to examine PDF content and improve parsing logic.
"""

import PyPDF2
from pdf_to_excel import PDFProcessor

def examine_pdf_content(pdf_path):
    """Examine the content of a PDF to understand its structure."""
    print(f"Examining PDF: {pdf_path}")
    print("="*50)
    
    # Extract raw text
    processor = PDFProcessor()
    text = processor.extract_text_from_pdf(pdf_path)
    
    if text:
        lines = text.split('\n')
        print(f"Total lines extracted: {len(lines)}")
        print("\nFirst 20 lines:")
        print("-"*30)
        for i, line in enumerate(lines[:20]):
            if line.strip():
                print(f"{i+1:3d}: {line[:100]}")  # Show first 100 chars
        
        print(f"\nLast 10 lines:")
        print("-"*30)
        for i, line in enumerate(lines[-10:]):
            if line.strip():
                print(f"{len(lines)-10+i+1:3d}: {line[:100]}")
        
        # Look for patterns
        print(f"\nLine length analysis:")
        line_lengths = [len(line) for line in lines if line.strip()]
        if line_lengths:
            print(f"Average line length: {sum(line_lengths)/len(line_lengths):.1f}")
            print(f"Min line length: {min(line_lengths)}")
            print(f"Max line length: {max(line_lengths)}")
        
        # Sample some lines to understand structure
        print(f"\nSample lines with good length (20-100 chars):")
        print("-"*50)
        sample_lines = [line for line in lines if 20 <= len(line.strip()) <= 100]
        for i, line in enumerate(sample_lines[:10]):
            print(f"{i+1}: {line.strip()}")
            
    else:
        print("No text extracted from PDF")

if __name__ == "__main__":
    examine_pdf_content("sample_data/P11569-11-99-40-2619.pdf")
