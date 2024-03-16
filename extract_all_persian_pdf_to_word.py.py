import os
import re
import tempfile
from tqdm import tqdm
from pdf2image import convert_from_path
import pytesseract
from PIL import Image
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH

def sanitize_text(text):
    """Sanitize text to make it XML compatible."""
    # Remove any characters that are not XML compatible
    sanitized_text = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', '', text)
    return sanitized_text

def pdf_to_word(pdf_path, lang='fas', **kwargs):
    """Converts PDF to Word document using pytesseract OCR."""
    # Derive the Word document path from the PDF path
    word_path = pdf_path.rsplit('.', 1)[0] + ".docx"
    
    # Check if the Word document already exists
    if os.path.exists(word_path):
        print(f"Word document already exists: {word_path}")
        return
    
    pdf_name = os.path.basename(pdf_path).split('.')[0]
    pages = convert_from_path(pdf_path, **kwargs)
    texts = []

    print(f'Converting PDF to document: {pdf_name} [{len(pages)} pages]')
    for i, page in tqdm(enumerate(pages), total=len(pages), position=0, leave=True):
        with tempfile.TemporaryDirectory() as img_dir:
            img_name = f'{pdf_name}-{i+1}.jpg'
            img_path = os.path.join(img_dir, img_name)
            
            page.save(img_path, 'JPEG')
            text = pytesseract.image_to_string(Image.open(img_path), lang=lang)
            sanitized_text = sanitize_text(text)
            texts.append(sanitized_text)

    document = Document()
    for i, text in enumerate(texts):
        if i:  # Add a page break if not the first page
            document.add_page_break()
        paragraph = document.add_paragraph(text)
        paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    # Save the Word document next to the original PDF
    document.save(word_path)
    print(f'Document saved: "{word_path}"')

def process_directory(root_dir, lang='fas'):
    """Recursively process each PDF file in directory and subdirectories."""
    for root, dirs, files in os.walk(root_dir):
        for file in files:
            if file.lower().endswith(".pdf"):
                pdf_path = os.path.join(root, file)
                # Get the number of pages in the PDF
                with open(pdf_path, 'rb') as f:
                    num_pages = len(convert_from_path(pdf_path))
                # Process the PDF if the number of pages is 60 or fewer
                pdf_to_word(pdf_path, lang)

# Specify the root directory where your PDFs are located
root_pdf_directory = './root/'

# Process all PDFs in the root directory and its subdirectories
process_directory(root_pdf_directory)
