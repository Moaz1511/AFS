import PyPDF2
import os
import tkinter as tk
from tkinter import filedialog

def get_pdf_page_count(pdf_path):
    try:
        with open(pdf_path, 'rb') as file:
            pdf_reader = PyPDF2.PdfReader(file)
            return len(pdf_reader.pages)
    except Exception as e:
        print(f"Error reading {pdf_path}: {e}")
        return 0

def count_pages_in_multiple_pdfs(pdf_folder_path):
    pdf_page_counts = {}
    
    # Loop through all files in the specified directory
    for filename in os.listdir(pdf_folder_path):
        if filename.endswith('.pdf'):
            # Ensure the correct file path format
            file_path = os.path.join(pdf_folder_path, filename)
            page_count = get_pdf_page_count(file_path)
            pdf_page_counts[filename] = page_count
            
    return pdf_page_counts

def select_folder():
    root = tk.Tk()
    root.withdraw()  # Hide the main Tkinter window
    folder_path = filedialog.askdirectory(title="Select Folder Containing PDFs")
    return folder_path

# Example Usage
folder_path = select_folder()  # Open folder selection dialog

if folder_path:
    page_counts = count_pages_in_multiple_pdfs(folder_path)
    
    # Print the page count for each PDF
    for pdf, pages in page_counts.items():
        print(f'{pdf}: {pages} pages')
else:
    print("No folder selected.")
