import os
from PyPDF2 import PdfReader, PdfWriter
from pdfminer.high_level import extract_text

def extract_text_from_page(pdf_path, page_num):
    """
    Extracts text from a specific page of a PDF file.
    """
    text = extract_text(pdf_path, page_numbers=[page_num])
    return text

def split_pdf_by_keyword(input_pdf_path, output_dir, keyword):
    # Ensure the output directory exists
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # Open the input PDF
    with open(input_pdf_path, "rb") as input_pdf_file:
        reader = PdfReader(input_pdf_file)

        current_writer = None
        invoice_count = 0

        # Loop through all the pages of the PDF
        for page_num in range(len(reader.pages)):
            # Extract text from the current page
            text = extract_text_from_page(input_pdf_path, page_num)

            # Check if the keyword exists in the text
            if keyword.lower() in text.lower():
                # If we find the keyword, start a new PDF file
                if current_writer:
                    # Save the previous invoice PDF
                    output_pdf_path = os.path.join(output_dir, f"invoice_{invoice_count}.pdf")
                    with open(output_pdf_path, "wb") as output_pdf_file:
                        current_writer.write(output_pdf_file)
                    print(f"Saved: {output_pdf_path}")
                
                # Start a new writer for the next invoice
                current_writer = PdfWriter()
                invoice_count += 1
            
            # Add the current page to the current invoice
            if current_writer:
                current_writer.add_page(reader.pages[page_num])

        # Save the last invoice PDF
        if current_writer:
            output_pdf_path = os.path.join(output_dir, f"invoice_{invoice_count}.pdf")
            with open(output_pdf_path, "wb") as output_pdf_file:
                current_writer.write(output_pdf_file)
            print(f"Saved: {output_pdf_path}")

# Example usage
input_pdf_path = r"C:\Users\Rajat Gupta\Downloads\Naukri_Rajat CV34.pdf"  # Input PDF with multiple invoices
output_dir = "C:\Me\Learning\Task\pdf split"  # Directory to save individual invoice PDFs
keyword = "Invoice"  # Keyword that indicates the start of a new invoice

split_pdf_by_keyword(input_pdf_path, output_dir, keyword)
