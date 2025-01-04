import PyPDF2
from docx import Document
import csv
import re

def clean_text(text):
    # Remove any control characters and NULL bytes
    # Replaces control characters with a space
    clean = re.sub(r'[\x00-\x1F\x7F]', ' ', text)
    return clean

def pdf_to_text(pdf_path, txt_path):
    # Open the PDF file
    with open(pdf_path, 'rb') as pdf_file:
        reader = PyPDF2.PdfReader(pdf_file)
        
        # Initialize an empty string to store all the text
        all_text = ""
        
        # Loop through all pages in the PDF
        for page_num in range(len(reader.pages)):
            page = reader.pages[page_num]
            all_text += page.extract_text()
        
        # Save the extracted text to a text file
        with open(txt_path, 'w', encoding='utf-8') as text_file:
            text_file.write(all_text)

    #print(f"Text has been saved to {txt_path}")

def txt_to_docx(txt_path, docx_path):
    # Open the text file
    with open(txt_path, 'r', encoding='utf-8') as txt_file:
        # Read all lines of the text file
        lines = txt_file.readlines()
    
    # Create a new Document
    doc = Document()
    
    # Add each line as a new paragraph in the document
    for line in lines:
        # Clean the line to remove any control characters or NULL bytes
        cleaned_line = clean_text(line.strip())  # Remove leading/trailing spaces
        if cleaned_line:  # Avoid adding empty lines
            doc.add_paragraph(cleaned_line)
    
    # Save the document as a .docx file
    doc.save(docx_path)
    
    #print(f"Text from {txt_path} has been saved as {docx_path}")

# Example usage
Count = 0
with open("file_name.csv", 'r', encoding='utf-8') as file:
    file_name = csv.reader(file)
    file_list = list(file_name)
    for name in file_list:
        Count += 1
        pdf_path = name[0]  # Replace with the path to your PDF
        txt_path = str(name[0]) + "_output.txt"  # Path where you want to save the text
        pdf_to_text(pdf_path, txt_path)

        txt_path = str(name[0]) + "_output.txt"  # Replace with the path to your text file
        docx_path = str(name[0]) + "_output.docx"  # Path where you want to save the Word file
        txt_to_docx(txt_path, docx_path)
        print("Done " + str(Count) + " File")
