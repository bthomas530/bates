import os
from pathlib import Path
import re
from PyPDF2 import PdfReader
import tkinter as tk
from tkinter import filedialog, messagebox
from datetime import datetime

def extract_email_info_from_pdf(pdf_path):
    """Extract date, sender, and subject from email headers in PDF content."""
    try:
        with open(pdf_path, 'rb') as file:
            reader = PdfReader(file)
            
            # Look through first few pages for email headers
            for page_num in range(min(3, len(reader.pages))):
                text = reader.pages[page_num].extract_text()
                
                # Extract sender email - look for mail@tx.lotto.com specifically
                from_match = re.search(r'From:.*?(mail@tx\.lotto\.com)', text)
                sender = from_match.group(1) if from_match else 'unknown'
                
                # Extract subject - look for text between Subject: and Date:
                subject_match = re.search(r'Subject:\s*(.*?)(?:\nDate:|$)', text, re.DOTALL)
                subject = subject_match.group(1).strip() if subject_match else 'No Subject'
                
                # Extract date and time
                date_match = re.search(r'Date:\s*(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2},\s+\d{4}\s+at\s+(\d{1,2}):(\d{2})\s+(?:AM|PM)', text)
                if date_match:
                    date_str = date_match.group(0)
                    # Extract just the date part
                    date_str = re.search(r'(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2},\s+\d{4}', date_str).group(0)
                    try:
                        # Parse the date
                        date_obj = datetime.strptime(date_str, '%B %d, %Y')
                        # Get hours and minutes
                        hour = int(date_match.group(1))
                        minute = int(date_match.group(2))
                        # Adjust for AM/PM
                        if 'PM' in date_match.group(0) and hour != 12:
                            hour += 12
                        elif 'AM' in date_match.group(0) and hour == 12:
                            hour = 0
                        # Format as YYYYMMDDHHMM
                        return date_obj.strftime('%Y%m%d') + f'{hour:02d}{minute:02d}', sender, subject
                    except ValueError:
                        continue
            
            print(f"\nNo email header date found in: {pdf_path.name}")
            print(f"First 500 chars of text:")
            print(text[:500])
            
    except Exception as e:
        print(f"Error processing {pdf_path}: {str(e)}")
    return None, None, None

def rename_files(directory):
    """Rename PDFs in the directory based on their date."""
    dir_path = Path(directory)
    renamed = 0
    errors = 0
    total_files = 0
    
    print(f"\nChecking directory: {dir_path}")
    print(f"Directory exists: {dir_path.exists()}")
    print(f"Directory is directory: {dir_path.is_dir()}")
    
    # List all files in directory
    print("\nFiles in directory:")
    for item in dir_path.iterdir():
        print(f"- {item.name} ({item.suffix})")
    
    # First count total PDF files
    for _ in dir_path.rglob('*.pdf'):
        total_files += 1
    
    if total_files == 0:
        print(f"\nNo PDF files found in directory: {directory}")
        print("Please make sure you selected a directory containing PDF files.")
        return 0, 0
    
    print(f"\nFound {total_files} PDF files to process")
    
    # Process all PDF files recursively
    for pdf_file in dir_path.rglob('*.pdf'):
        try:
            print(f"\nProcessing: {pdf_file.name}")
            date_str, sender, subject = extract_email_info_from_pdf(pdf_file)
            if date_str:
                # Clean subject for filename
                subject = re.sub(r'[<>:"/\\|?*]', '_', subject)
                subject = subject[:50]  # Limit length
                
                new_name = f"{date_str}_sender:{sender}_subject:{subject}.pdf"
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
    directory = filedialog.askdirectory(title="Select Directory with PDF Files")
    if not directory:
        print("No directory selected")
        return
    
    print(f"\nProcessing files in: {directory}")
    renamed, errors = rename_files(directory)
    
    print(f"\nDone! Renamed: {renamed}, Errors: {errors}")
    messagebox.showinfo("Complete", f"Renamed: {renamed}\nErrors: {errors}")

if __name__ == "__main__":
    main() 