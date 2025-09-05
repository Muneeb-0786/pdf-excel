# PDF to Excel Processor - Project Instructions

This workspace contains a Python application that processes PDF files and generates Excel attribute tables.

## Project Overview
- **Language**: Python 3.13
- **Purpose**: Extract data from PDFs (text-based and scanned) and format as Excel attribute tables
- **Key Libraries**: PyPDF2, pytesseract, pandas, openpyxl

## Key Files
- `pdf_to_excel.py` - Main application with PDFProcessor class
- `example_usage.py` - Usage examples and demonstrations
- `requirements.txt` - Python dependencies
- `sample_data/` - Input PDF files for testing
- `output/` - Generated Excel files and sample formats

## Usage
```bash
python pdf_to_excel.py input.pdf -o output.xlsx [--ocr] [--verbose]
```

## Development Guidelines
- Follow PEP 8 style guidelines
- Use type hints where possible
- Include comprehensive error handling
- Test with both text-based and scanned PDFs
- Maintain compatibility with SAMPLE.xlsx format structure

Project is complete and ready for use.
