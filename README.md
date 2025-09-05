# PDF to Excel Processor

A Python application that processes PDF files (both text-based and scanned) and extracts technical data into structured Excel files matching specific attribute table formats.

## Features

- **Dual PDF Processing**: Handles both standard text-based PDFs and scanned PDFs using OCR
- **Smart Text Extraction**: Uses PyPDF2 for standard PDFs and pytesseract for OCR processing
- **Structured Output**: Generates Excel files matching specific attribute table formats
- **Technical Data Focus**: Optimized for instrument data sheets and technical documents
- **Batch Processing**: Process multiple PDF files at once
- **Customizable Parsing**: Easily adaptable parsing logic for different document types

## Installation

1. **Clone or download this repository**
2. **Install Python 3.8 or later**
3. **Install required packages:**
   ```bash
   pip install -r requirements.txt
   ```

### Dependencies

- `PyPDF2==3.0.1` - PDF text extraction
- `pdf2image==1.17.0` - Convert PDF pages to images for OCR
- `pytesseract==0.3.10` - OCR text extraction
- `pandas==2.1.4` - Data manipulation and Excel export
- `openpyxl==3.1.2` - Excel file formatting
- `Pillow==10.1.0` - Image processing support

### Additional Requirements for OCR

For OCR functionality, you'll need to install Tesseract:

- **Windows**: Download from [GitHub Tesseract releases](https://github.com/UB-Mannheim/tesseract/wiki)
- **macOS**: `brew install tesseract`
- **Linux**: `sudo apt-get install tesseract-ocr`

## Usage

### Command Line Interface

#### Basic Usage
```bash
python pdf_to_excel.py input.pdf -o output.xlsx
```

#### Advanced Options
```bash
# Force OCR processing
python pdf_to_excel.py input.pdf -o output.xlsx --ocr

# Batch process a directory
python pdf_to_excel.py input_folder/ -o output_folder/ --batch

# Enable verbose logging
python pdf_to_excel.py input.pdf -o output.xlsx --verbose
```

#### Command Line Arguments
- `input`: Input PDF file or directory (required)
- `-o, --output`: Output Excel file or directory (optional, auto-generated if not provided)
- `--ocr`: Force OCR processing even if text extraction works
- `--batch`: Batch process all PDFs in a directory
- `--verbose`: Enable detailed logging

### Python API Usage

```python
from pdf_to_excel import PDFProcessor

# Initialize processor
processor = PDFProcessor()

# Process single PDF
rows_processed = processor.process_pdf_to_excel(
    pdf_path="sample.pdf",
    excel_path="output.xlsx"
)

# Process with OCR
rows_processed = processor.process_pdf_to_excel(
    pdf_path="scanned.pdf", 
    excel_path="output.xlsx",
    use_ocr=True
)

# Batch processing
results = processor.batch_process(
    input_dir="pdf_files/",
    output_dir="excel_files/",
    use_ocr=False
)
```

## Output Format

The application generates Excel files with the following structure:

| Column | Description | Example |
|--------|-------------|---------|
| Field | Functional Location | 11-18-XTGD-5403 |
| TPLNR | Equipment/Class Number | FG-FGAS |
| CLASS | Class Type | 003 |
| KLART | Position Number | 1, 2, 3... |
| POSNUMMER | Characteristic Name | NACC01, HSE-HAZ_AREA |
| ATNAM | Characteristic Value | ±10, Hazardous Area |
| ATWRT | Description | Accuracy, Area Classification |
| Characteristics UoM | Unit of Measure | %, ppm, °C |
| Remarks | Additional Remarks | |
| Additional | Additional Information | |
| REF | Document Reference | P11569-11-99-40-2619-1 |

## Project Structure

```
pdf_to_excel/
│
├── pdf_to_excel.py          # Main application
├── example_usage.py         # Usage examples
├── examine_sample.py        # Utility to examine Excel files
├── test_pdf_content.py      # Utility to examine PDF content
├── requirements.txt         # Python dependencies
├── README.md               # This documentation
│
├── sample_data/            # Input PDF files
│   └── P11569-11-99-40-2619.pdf
│
├── output/                 # Generated Excel files
│   ├── SAMPLE.xlsx         # Reference format
│   └── *.xlsx              # Generated outputs
│
└── .github/
    └── copilot-instructions.md
```

## Customization

### Parsing Logic

The parsing logic in `format_text_to_structure()` can be customized for different document types:

1. **Modify attribute templates** in the `equipment_attributes` list
2. **Adjust regex patterns** for document references and technical data
3. **Update skip patterns** for headers/footers specific to your documents
4. **Customize functional location detection** based on your naming conventions

### Output Format

The Excel output format can be modified in the `write_to_excel()` method:

1. **Change column headers** in the `headers` list
2. **Adjust formatting** in `_format_excel_sheet_attributes()`
3. **Modify column widths** and styles as needed

## Troubleshooting

### Common Issues

1. **OCR taking too long**: OCR processing can be slow for large PDFs. Use `--verbose` to monitor progress.

2. **No text extracted**: Some PDFs may need OCR even if they appear to have text. Try using the `--ocr` flag.

3. **Import errors**: Ensure all dependencies are installed:
   ```bash
   pip install -r requirements.txt
   ```

4. **Tesseract not found**: Make sure Tesseract is installed and added to your system PATH.

### Logging

Enable verbose logging to troubleshoot issues:
```bash
python pdf_to_excel.py input.pdf --verbose
```

## Examples

See `example_usage.py` for comprehensive usage examples including:
- Single PDF processing
- Batch processing
- OCR processing
- Error handling

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test thoroughly
5. Submit a pull request

## License

This project is provided as-is for educational and development purposes.

## Support

For issues or questions:
1. Check the troubleshooting section
2. Review the example usage
3. Enable verbose logging to diagnose problems
4. Examine the sample files provided
