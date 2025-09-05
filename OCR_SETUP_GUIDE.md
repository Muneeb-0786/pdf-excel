# Setup Guide for OCR Processing

## Installing Poppler for PDF to Image Conversion

To process scanned PDFs with OCR, you need to install poppler-utils which is required by pdf2image.

### Windows Installation

1. **Download Poppler for Windows**:
   - Go to: https://github.com/oschwartz10612/poppler-windows/releases/
   - Download the latest release (e.g., `poppler-xx.xx.x-h66b187d_0.tar.gz`)

2. **Extract and Install**:
   - Extract the downloaded file to a folder like `C:\Program Files\poppler`
   - Add the `bin` folder to your Windows PATH:
     - `C:\Program Files\poppler\bin`

3. **Add to Windows PATH**:
   - Press `Win + X` and select "System"
   - Click "Advanced system settings"
   - Click "Environment Variables"
   - Under "System Variables", find "Path" and click "Edit"
   - Click "New" and add: `C:\Program Files\poppler\bin`
   - Click "OK" on all dialogs

4. **Verify Installation**:
   ```cmd
   pdftoppm -h
   ```
   You should see the help text for pdftoppm.

### Alternative: Using Conda

If you have conda installed:
```bash
conda install -c conda-forge poppler
```

### Alternative: Using Chocolatey

If you have Chocolatey installed:
```cmd
choco install poppler
```

## After Installing Poppler

Once poppler is installed, you can process the scanned PDF:

```bash
# Process the scanned PDF with OCR
python pdf_to_excel.py sample_data/P11569-11-99-40-1605-1.PDF -o output/P11569-11-99-40-1605-1_OCR.xlsx --ocr --verbose

# Then create a new combined file with actual data from both PDFs
python pdf_to_excel.py sample_data/ -o output/ --batch --verbose
```

## Testing OCR Setup

You can test if OCR is working with this command:
```bash
python -c "from pdf2image import convert_from_path; import pytesseract; print('OCR setup is working!')"
```

If you get any errors, the most common issues are:
1. Poppler not in PATH
2. Tesseract not in PATH
3. Missing dependencies

## Current Status

✅ **Tesseract**: Already installed on your system
❌ **Poppler**: Needs to be installed for pdf2image
✅ **Python packages**: All installed (PyPDF2, pdf2image, pytesseract, pandas, openpyxl)

## Files Created

Currently you have:
- `output/combined_both_pdfs_clean.xlsx` - Clean combined file with:
  - PDF 1 data: 176 rows (actual data from P11569-11-99-40-2619.pdf)
  - PDF 2 data: 176 rows (placeholder data for P11569-11-99-40-1605-1.PDF)

Once poppler is installed, you can process the actual scanned PDF and get real data instead of placeholder data.
