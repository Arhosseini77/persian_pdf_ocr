# Persian PDF to Word Converter

This script converts all PDF files in a directory to Word documents using Persian Tesseract OCR.

## Quick Installation

### Installing Tesseract OCR

#### For Ubuntu:

```bash
sudo apt-get update
sudo apt-get install tesseract-ocr
sudo apt-get install tesseract-ocr-fas  # For Persian language support
```

#### For Windows:

You can download and install Tesseract OCR from the [official website](https://github.com/tesseract-ocr/tesseract).

### Installing Python Dependencies

You can install the required Python packages using pip:

```bash
pip install tqdm pdf2image pytesseract Pillow python-docx
```

## Running the Script

After installing the necessary dependencies and Tesseract OCR, you can run the script `extract_all_persian_pdf_to_word.py`:

```bash
python extract_all_persian_pdf_to_word.py
```

