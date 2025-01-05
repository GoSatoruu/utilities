import PyPDF2
from docx import Document
import csv
import re
import os

def clean_text(text):
    """Remove control characters and NULL bytes from the text."""
    clean = re.sub(r'[\x00-\x1F\x7F]', ' ', text)
    return clean

def pdf_to_text(pdf_path, txt_path):
    """Extract text from a PDF and save it to a text file."""
    try:
        # Check if PDF exists and is accessible
        if not os.path.exists(pdf_path) or not os.access(pdf_path, os.R_OK):
            print(f"Permission denied or file does not exist: {pdf_path}")
            return False
        
        with open(pdf_path, 'rb') as pdf_file:
            reader = PyPDF2.PdfReader(pdf_file)
            all_text = ""
            
            # Loop through all pages in the PDF
            for page_num in range(len(reader.pages)):
                page = reader.pages[page_num]
                all_text += page.extract_text() or ""  # Handle case where text is None
            
            # Check if there is any text extracted
            if not all_text.strip():
                print(f"No text extracted from {pdf_path}")
                return False
            
            # Save the extracted text to a text file
            try:
                with open(txt_path, 'w', encoding='utf-8') as text_file:
                    text_file.write(all_text)
            except PermissionError:
                print(f"Permission denied while writing to {txt_path}")
                return False

    except Exception as e:
        print(f"Error processing PDF {pdf_path}: {e}")
        return False
    return True

def txt_to_docx(txt_path, docx_path):
    """Convert the cleaned text file into a DOCX file."""
    try:
        # Check if TXT file exists and is accessible
        if not os.path.exists(txt_path) or not os.access(txt_path, os.R_OK):
            print(f"Permission denied or file does not exist: {txt_path}")
            return False
        
        with open(txt_path, 'r', encoding='utf-8') as txt_file:
            lines = txt_file.readlines()
        
        # Create a new Document
        doc = Document()
        
        # Add each line as a new paragraph in the document
        for line in lines:
            cleaned_line = clean_text(line.strip())  # Clean each line
            if cleaned_line:  # Avoid adding empty lines
                doc.add_paragraph(cleaned_line)
        
        # Save the document as a .docx file
        try:
            doc.save(docx_path)
        except PermissionError:
            print(f"Permission denied while saving DOCX: {docx_path}")
            return False
    
    except Exception as e:
        print(f"Error converting text to DOCX {txt_path} to {docx_path}: {e}")
        return False
    return True

def delete_txt_file(txt_path):
    """Delete the temporary text file after conversion."""
    try:
        if os.path.exists(txt_path):
            os.remove(txt_path)
            print(f"Deleted temporary text file: {txt_path}")
        else:
            print(f"Text file not found: {txt_path}")
    except Exception as e:
        print(f"Error deleting text file {txt_path}: {e}")

def process_files(csv_file_path):
    """Process the PDF files listed in the CSV and convert them to DOCX."""
    count = 0
    try:
        with open(csv_file_path, 'r', encoding='utf-8') as file:
            file_name = csv.reader(file)
            file_list = list(file_name)
            
            for name in file_list:
                count += 1
                pdf_path = name[0]  # Path to the PDF file
                txt_path = f"{name[0]}_output.txt"  # Path for the text file
                docx_path = f"{name[0]}_output.docx"  # Path for the DOCX file
                
                # Convert PDF to text
                if not pdf_to_text(pdf_path, txt_path):
                    print(f"Skipping PDF: {pdf_path} due to error.")
                    continue  # Skip to the next file if PDF conversion failed
                
                # Convert text to DOCX
                if not txt_to_docx(txt_path, docx_path):
                    print(f"Skipping DOCX conversion for {txt_path}.")
                    continue  # Skip to the next file if DOCX conversion failed
                
                # Delete the temporary TXT file after conversion
                delete_txt_file(txt_path)
                
                print(f"Successfully processed {count} file(s): {pdf_path}")
                
    except FileNotFoundError:
        print(f"CSV file not found: {csv_file_path}")
    except Exception as e:
        print(f"An error occurred: {e}")
    return count

# Example usage
csv_file = "file_name.csv"  # Make sure this path is correct
processed_count = process_files(csv_file)
print(f"Processed {processed_count} files successfully.")
