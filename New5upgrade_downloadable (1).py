import os
import glob
import pandas as pd
import re
import pytesseract
from pdf2image import convert_from_path
from PyPDF2 import PdfReader, PdfWriter
from pdfminer.high_level import extract_text
from tqdm import tqdm
import logging
import shutil
import uuid
from datetime import datetime
from concurrent.futures import ProcessPoolExecutor, as_completed
from PIL import Image, ImageFilter

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# Ensure Tesseract is available
pytesseract.pytesseract.tesseract_cmd = r'C:\Users\RajatGupta\AppData\Local\Programs\Tesseract-OCR\tesseract.exe'

# Load vendor names from Excel
file_path = r'C:\Users\RajatGupta\Me\Task\Pdf split\vendor list\Vendor list.xlsx'
df = pd.read_excel(file_path, usecols=['Vendor Name'])
VENDOR_NAMES = df['Vendor Name'].tolist()

def preprocess_image(image):
    """Enhances image for better OCR accuracy."""
    processed = image.convert('L')  # Convert to grayscale
    processed = processed.filter(ImageFilter.MedianFilter())
    processed = processed.point(lambda x: 0 if x < 140 else 255, '1')  # Binarization
    return processed

def extract_text_from_page(pdf_path, page_num, dpi=300):
    """Extracts text from a specific page of a PDF file using pdfminer and OCR."""
    try:
        text = extract_text(pdf_path, page_numbers=[page_num]).strip()
    except Exception as e:
        logging.error(f"Error extracting text from page {page_num + 1} of {pdf_path}: {e}")
        text = ""
    
    if not text:
        try:
            images = convert_from_path(pdf_path, first_page=page_num + 1, last_page=page_num + 1, fmt='png', dpi=dpi)
            for image in images:
                processed_image = preprocess_image(image)
                text = pytesseract.image_to_string(processed_image).strip()
                break  # Only process the first image/page
        except Exception as e:
            logging.error(f"Error processing image from page {page_num + 1} of {pdf_path}: {e}")
            text = ""
    
    return text

def extract_invoice_no(text):
    """Extracts the Invoice No from the given text using regex."""
    patterns = [
        r'Invoice\s*No\.?\s*[:\-]?\s*([\w\-_=]+)',
        r'INVOICE\s*#\s*[:\-]?\s*([\w\-_=]+)',
        r'Invoice\s*Number\s*[:\-]?\s*([\w\-_=]+)',
        r'Invoice\s*ID\s*[:\-]?\s*([\w\-_=]+)',
    ]
    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            invoice_no = match.group(1).strip()
            return invoice_no
    return None

def extract_account_no(text):
    """Extracts the Account No from the given text using regex."""
    patterns = [
        r'Account\s*No\.?\s*[:\-]?\s*([\w\-_=]+)', 
        r'Account\s*Number\s*[:\-]?\s*([\w\-_=]+)', 
        r'Acct\s*No\.?\s*[:\-]?\s*([\w\-_=]+)', 
        ]
    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            account_no = match.group(1).strip()
            return account_no
    return None

def extract_vendor_name(text):
    """Extracts the vendor name from the given text based on predefined vendor names."""
    for vendor in VENDOR_NAMES:
        if vendor.lower() in text.lower():
            return vendor
    return None

def find_split_keyword(vendor_name):
    """Finds the split keyword corresponding to the vendor name."""
    vendor_list_path = r'C:\Users\RajatGupta\Me\Task\Pdf split\vendor list\Vendor list.xlsx'
    df = pd.read_excel(vendor_list_path, sheet_name='Sheet1')
    matching_rows = df[df['Vendor Name'].str.contains(vendor_name, case=False, na=False)]
    if not matching_rows.empty:
        index = matching_rows.index[0]
        split_keyword_value = df.at[index, 'Split_keyword']
        return split_keyword_value
    else:
        logging.warning(f'No matching vendor found for: {vendor_name}')
        return None

def find_first_keyword(text, keywords):
    """Finds the first keyword from the list of keywords in the given text."""
    for keyword in keywords:
        if keyword.lower() in text.lower():
            return keyword
    return None

def find_split_points_for_first_keyword(pdf_path, first_keyword):
    """Identify all page ranges based on the first keyword found in the PDF."""
    reader = PdfReader(pdf_path)
    total_pages = len(reader.pages)
    split_ranges = []
    current_start = 0

    for page_num in range(total_pages):
        text = extract_text_from_page(pdf_path, page_num)
        if first_keyword.lower() in text.lower():
            if current_start < page_num:  # Only add a range if it's valid
                split_ranges.append((current_start, page_num - 1))
            current_start = page_num

    if current_start < total_pages:
        split_ranges.append((current_start, total_pages - 1))

    return split_ranges

def is_valid_invoice_no(invoice_no):
    """Checks if the invoice number is valid (example validation)."""
    return bool(invoice_no)  # Placeholder: implement actual validation logic

def is_valid_account_no(account_no):
    """Checks if the account number is valid (example validation)."""
    return bool(account_no)  # Placeholder: implement actual validation logic

def split_pdf(pdf_path, output_dir, split_ranges):
    """Split the PDF into separate invoices based on split_ranges."""
    reader = PdfReader(pdf_path)
    base_name = os.path.splitext(os.path.basename(pdf_path))[0]
    split_outputs = []

    # Create date-specific folder
    date_folder = datetime.now().strftime('%Y-%m-%d')

    for idx, (start, end) in enumerate(split_ranges):
        writer = PdfWriter()
        
        keyword_page_text = extract_text_from_page(pdf_path, start)
        writer.add_page(reader.pages[start])

        for page_num in range(start + 1, end + 1):
            writer.add_page(reader.pages[page_num])

        invoice_no = extract_invoice_no(keyword_page_text)
        account_no = extract_account_no(keyword_page_text)
        vendor_name = extract_vendor_name(keyword_page_text)

        if invoice_no and is_valid_invoice_no(invoice_no):
            output_subdir = os.path.join('Utility Invoice', date_folder)
            output_pdf_name = f"Invoice_{invoice_no}_{uuid.uuid4().hex[:8]}"
            if vendor_name:
                output_pdf_name += f"_{vendor_name}"
        elif account_no and is_valid_account_no(account_no):
            output_subdir = os.path.join('Utility Invoice', date_folder)
            output_pdf_name = f"Account_{account_no}_{uuid.uuid4().hex[:8]}"
            if vendor_name:
                output_pdf_name += f"_{vendor_name}"
        elif vendor_name or account_no or invoice_no:
            output_subdir = os.path.join('Utility Invoice', date_folder)
            output_pdf_name = f"{vendor_name or 'Unknown'}_Other_{uuid.uuid4().hex[:8]}"
        else:
            output_subdir = os.path.join('Invalid', date_folder)
            output_pdf_name = f"Invalid_{uuid.uuid4().hex[:8]}"

        base_folder_path = os.path.join(output_dir, output_subdir, base_name)
        os.makedirs(base_folder_path, exist_ok=True)

        output_path = os.path.join(base_folder_path, f"{output_pdf_name}.pdf")
        with open(output_path, 'wb') as f:
            writer.write(f)

        split_outputs.append(output_path)
        logging.info(f"Processed {output_pdf_name}. PDF saved at {output_path}")

    return split_outputs

def process_pdf_file(pdf_path, output_dir, keywords):
    """Main function to process a single PDF file based on the first keyword found."""
    first_keyword_found = None
    reader = PdfReader(pdf_path)
    total_pages = len(reader.pages)
    
    vendor_name = None
    for page_num in range(total_pages):
        text = extract_text_from_page(pdf_path, page_num)
        vendor_name = extract_vendor_name(text)
        if vendor_name:
            logging.info(f"Vendor name '{vendor_name}' found in {pdf_path} on page {page_num + 1}")
            break

    if vendor_name:
        split_keyword = find_split_keyword(vendor_name)
        if split_keyword:
            logging.info(f"Using split keyword '{split_keyword}' for PDF splitting.")
            first_keyword_found = split_keyword
        else:
            logging.warning(f"No split keyword found for vendor '{vendor_name}'. Using default keywords.")
    else:
        logging.warning(f"Vendor name not found in {pdf_path}. Using default keywords.")
    
    if not first_keyword_found:
        for page_num in range(total_pages):
            text = extract_text_from_page(pdf_path, page_num)
            first_keyword_found = find_first_keyword(text, keywords)
            if first_keyword_found:
                logging.info(f"First keyword '{first_keyword_found}' found in {pdf_path} on page {page_num + 1}")
                break

    if not first_keyword_found:
        logging.warning(f"No keywords found in {pdf_path}. Moving to 'Invalid' folder.")
        invalid_folder = os.path.join(output_dir, 'Invalid')
        os.makedirs(invalid_folder, exist_ok=True)
        shutil.move(pdf_path, os.path.join(invalid_folder, os.path.basename(pdf_path)))
        return []

    split_ranges = find_split_points_for_first_keyword(pdf_path, first_keyword_found)
    return split_pdf(pdf_path, output_dir, split_ranges)

def main(input_dir, output_dir, keywords):
    """Processes all PDF files in the input directory."""
    pdf_files = glob.glob(os.path.join(input_dir, '*.pdf'))
    if not pdf_files:
        logging.info("No PDF files found for processing.")
        return

    with ProcessPoolExecutor() as executor:
        futures = {executor.submit(process_pdf_file, pdf, output_dir, keywords): pdf for pdf in pdf_files}
        for future in as_completed(futures):
            pdf = futures[future]
            try:
                output_files = future.result()
                logging.info(f"Finished processing {pdf}. Output files: {output_files}")
            except Exception as e:
                logging.error(f"Error processing {pdf}: {e}")

if __name__ == '__main__':
    input_dir = r'C:\Users\RajatGupta\Me\Task\Pdf split\Input'
    output_dir = r'C:\Users\RajatGupta\Me\Task\Pdf split\Output'
    keywords = ['Invoice', 'Bill', 'Payment']  # Keywords to identify split points

    main(input_dir, output_dir, keywords)
