import os
from pathlib import Path
import re
from PyPDF2 import PdfReader
import tkinter as tk
from tkinter import filedialog, messagebox

def extract_date(pdf_path):
    """Extract month and year from a PDF file."""
    try:
        with open(pdf_path, 'rb') as file:
            reader = PdfReader(file)
            text = reader.pages[0].extract_text()
            
            # First try to find MM-YYYY format
            mm_yyyy_match = re.search(r'(\d{2})-(\d{4})', text)
            if mm_yyyy_match:
                month, year = mm_yyyy_match.groups()
                return f"{year}{month}"
            
            # If not found, try Month YYYY format
            month_pattern = r'(?:January|February|March|April|May|June|July|August|September|October|November|December|Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)'
            year_pattern = r'20\d{2}'
            
            # Find first occurrence of month followed by year
            match = re.search(f"{month_pattern}.*?({year_pattern})", text, re.IGNORECASE)
            if match:
                year = match.group(1)
                # Find the month that precedes the year
                month_text = text[max(0, match.start()-20):match.start()].strip()
                month_match = re.search(month_pattern, month_text, re.IGNORECASE)
                if month_match:
                    month = month_match.group(0)
                    # Convert month to number
                    month_map = {
                        'jan': '01', 'feb': '02', 'mar': '03', 'apr': '04',
                        'may': '05', 'jun': '06', 'jul': '07', 'aug': '08',
                        'sep': '09', 'oct': '10', 'nov': '11', 'dec': '12'
                    }
                    month_num = month_map.get(month.lower()[:3])
                    if month_num:
                        return f"{year}{month_num}"
            
            print(f"\nNo date found in: {pdf_path.name}")
            print(f"First 500 chars of text:")
            print(text[:500])
    except Exception as e:
        print(f"Error processing {pdf_path}: {str(e)}")
    return None

def rename_files(directory):
    """Rename PDFs in the directory based on their date."""
    dir_path = Path(directory)
    renamed = 0
    errors = 0
    
    for pdf_file in dir_path.glob('*.pdf'):
        try:
            print(f"\nProcessing: {pdf_file.name}")
            date_str = extract_date(pdf_file)
            if date_str:
                new_name = f"{date_str}_{pdf_file.name}"
                pdf_file.rename(pdf_file.parent / new_name)
                print(f"Renamed: {pdf_file.name} -> {new_name}")
                renamed += 1
            else:
                errors += 1
        except Exception as e:
            print(f"Error with {pdf_file.name}: {str(e)}")
            errors += 1
    
    return renamed, errors

def main():
    directory = filedialog.askdirectory(title="Select Directory with PDFs")
    if not directory:
        print("No directory selected")
        return
    
    print(f"\nProcessing files in: {directory}")
    renamed, errors = rename_files(directory)
    
    print(f"\nDone! Renamed: {renamed}, Errors: {errors}")
    messagebox.showinfo("Complete", f"Renamed: {renamed}\nErrors: {errors}")

if __name__ == "__main__":
    main() 