import os
from pathlib import Path
import re
from PyPDF2 import PdfReader
from datetime import datetime

def extract_statement_date(pdf_path):
    """Extract the statement date from a PDF file."""
    try:
        with open(pdf_path, 'rb') as file:
            reader = PdfReader(file)
            # Look through first few pages for the statement date
            for page_num in range(min(3, len(reader.pages))):
                page = reader.pages[page_num]
                text = page.extract_text()
                
                # Look for "Statement Date" followed by a date pattern
                match = re.search(r'Statement Date:\s*(\d{1,2}[-/]\d{1,2}[-/]\d{2,4})', text)
                if match:
                    date_str = match.group(1)
                    # Convert date to YYYYMMDD format
                    try:
                        # Try different date formats
                        for fmt in ['%m/%d/%Y', '%m/%d/%y', '%d/%m/%Y', '%d/%m/%y']:
                            try:
                                date = datetime.strptime(date_str, fmt)
                                return date.strftime('%Y%m%d')
                            except ValueError:
                                continue
                    except ValueError:
                        print(f"Could not parse date {date_str} in {pdf_path}")
                        return None
    except Exception as e:
        print(f"Error processing {pdf_path}: {str(e)}")
        return None
    return None

def rename_pdfs(directory):
    """Rename PDFs in the directory based on their statement date."""
    dir_path = Path(directory)
    
    # Process all PDF files in the directory
    for pdf_file in dir_path.glob('*.pdf'):
        try:
            # Extract the statement date
            date_str = extract_statement_date(pdf_file)
            
            if date_str:
                # Create new filename with date prefix
                new_name = f"{date_str}_{pdf_file.name}"
                new_path = pdf_file.parent / new_name
                
                # Rename the file
                pdf_file.rename(new_path)
                print(f"Renamed: {pdf_file.name} -> {new_name}")
            else:
                print(f"No statement date found in: {pdf_file.name}")
                
        except Exception as e:
            print(f"Error processing {pdf_file.name}: {str(e)}")

if __name__ == "__main__":
    import sys
    
    if len(sys.argv) != 2:
        print("Usage: python rename_by_statement_date.py <directory>")
        sys.exit(1)
        
    directory = sys.argv[1]
    if not os.path.isdir(directory):
        print(f"Error: {directory} is not a valid directory")
        sys.exit(1)
        
    rename_pdfs(directory) 