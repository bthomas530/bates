import os
import email
from email import policy
from email.parser import BytesParser
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox
from datetime import datetime
import re

def extract_email_info(eml_path):
    """Extract date, sender, and subject from an email file."""
    try:
        with open(eml_path, 'rb') as f:
            msg = BytesParser(policy=policy.default).parse(f)
        
        # Get date received
        date_str = msg.get('date', '')
        try:
            date_obj = email.utils.parsedate_to_datetime(date_str)
            date_str = date_obj.strftime('%Y%Y%Y%Y%m%m')  # Format: YYYYMM
        except:
            print(f"Could not parse date from {eml_path.name}")
            return None
        
        # Get sender email
        from_addr = msg.get('from', 'Unknown')
        # Extract email address if it's in the format "Name <email@domain.com>"
        email_match = re.search(r'<(.+?)>', from_addr)
        if email_match:
            sender_email = email_match.group(1)
        else:
            # If no angle brackets, try to find an email pattern
            email_match = re.search(r'[\w\.-]+@[\w\.-]+', from_addr)
            if email_match:
                sender_email = email_match.group(0)
            else:
                sender_email = 'Unknown'
        
        # Clean sender email for filename
        sender_email = re.sub(r'[<>:"/\\|?*]', '_', sender_email)
        sender_email = sender_email[:30]  # Limit length
        
        # Get subject
        subject = msg.get('subject', 'No Subject')
        # Clean subject for filename
        subject = re.sub(r'[<>:"/\\|?*]', '_', subject)
        subject = subject[:50]  # Limit length
        
        return date_str, sender_email, subject
        
    except Exception as e:
        print(f"Error processing {eml_path}: {str(e)}")
        return None

def rename_files(directory):
    """Rename email files in the directory based on their metadata."""
    dir_path = Path(directory)
    renamed = 0
    errors = 0
    
    for eml_file in dir_path.glob('*.eml'):
        try:
            print(f"\nProcessing: {eml_file.name}")
            info = extract_email_info(eml_file)
            if info:
                date_str, sender_email, subject = info
                new_name = f"{date_str}_{sender_email}_{subject}.eml"
                eml_file.rename(eml_file.parent / new_name)
                print(f"Renamed: {eml_file.name} -> {new_name}")
                renamed += 1
            else:
                errors += 1
        except Exception as e:
            print(f"Error with {eml_file.name}: {str(e)}")
            errors += 1
    
    return renamed, errors

def main():
    directory = filedialog.askdirectory(title="Select Directory with Email Files")
    if not directory:
        print("No directory selected")
        return
    
    print(f"\nProcessing files in: {directory}")
    renamed, errors = rename_files(directory)
    
    print(f"\nDone! Renamed: {renamed}, Errors: {errors}")
    messagebox.showinfo("Complete", f"Renamed: {renamed}\nErrors: {errors}")

if __name__ == "__main__":
    main() 