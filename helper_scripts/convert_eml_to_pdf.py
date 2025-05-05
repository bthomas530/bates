import os
import email
from email import policy
from email.parser import BytesParser
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox
from fpdf import FPDF
from datetime import datetime
import html
import re

class PDF(FPDF):
    def header(self):
        self.set_font('Helvetica', 'B', 12)
        self.cell(0, 10, 'Email Message', ln=True, align='C')
        self.ln(10)

def extract_email_content(msg):
    """Extract email content, handling both plain text and HTML."""
    content = []
    
    # Get email metadata
    subject = msg.get('subject', 'No Subject')
    from_addr = msg.get('from', 'Unknown Sender')
    to_addr = msg.get('to', 'Unknown Recipient')
    date = msg.get('date', 'Unknown Date')
    
    content.append(f"From: {from_addr}")
    content.append(f"To: {to_addr}")
    content.append(f"Subject: {subject}")
    content.append(f"Date: {date}")
    content.append("\n" + "="*50 + "\n")
    
    # Process email body
    if msg.is_multipart():
        for part in msg.walk():
            if part.get_content_type() == "text/plain":
                try:
                    text = part.get_content()
                    if isinstance(text, bytes):
                        text = text.decode()
                    content.append(text)
                except Exception as e:
                    content.append(f"Error decoding text part: {str(e)}")
            elif part.get_content_type() == "text/html":
                try:
                    html_content = part.get_content()
                    if isinstance(html_content, bytes):
                        html_content = html_content.decode()
                    # Basic HTML to text conversion
                    html_content = re.sub('<[^<]+?>', '', html_content)
                    html_content = html.unescape(html_content)
                    content.append(html_content)
                except Exception as e:
                    content.append(f"Error decoding HTML part: {str(e)}")
    else:
        try:
            text = msg.get_content()
            if isinstance(text, bytes):
                text = text.decode()
            content.append(text)
        except Exception as e:
            content.append(f"Error decoding content: {str(e)}")
    
    return "\n".join(content)

def convert_eml_to_pdf(eml_path, output_dir):
    """Convert a single EML file to PDF."""
    try:
        with open(eml_path, 'rb') as f:
            msg = BytesParser(policy=policy.default).parse(f)
        
        # Create PDF
        pdf = FPDF()
        pdf.add_page()
        
        # Set font and size
        pdf.set_font('Helvetica', '', 10)
        
        # Set margins
        pdf.set_margins(15, 15, 15)
        
        # Get content and split into lines
        content = extract_email_content(msg)
        lines = content.split('\n')
        
        # Add content to PDF
        for line in lines:
            # Handle long lines by wrapping
            if len(line) > 100:
                words = line.split()
                current_line = ""
                for word in words:
                    if len(current_line) + len(word) + 1 <= 100:
                        current_line += " " + word if current_line else word
                    else:
                        pdf.multi_cell(0, 5, current_line)
                        current_line = word
                if current_line:
                    pdf.multi_cell(0, 5, current_line)
            else:
                pdf.multi_cell(0, 5, line)
        
        # Generate output filename
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
        
        # Get date
        date = msg.get('date', '')
        try:
            date_obj = email.utils.parsedate_to_datetime(date)
            date_str = date_obj.strftime('%Y%m')  # Just year and month
        except:
            date_str = datetime.now().strftime('%Y%m')
        
        # Generate filename: YYYYMM_SenderEmail_Subject.pdf
        output_filename = f"{date_str}_{sender_email}_{subject}.pdf"
        output_path = output_dir / output_filename
        
        # Save PDF
        pdf.output(str(output_path))
        return True, output_filename
        
    except Exception as e:
        return False, str(e)

def process_directory(input_dir, output_dir):
    """Process all EML files in the input directory."""
    input_path = Path(input_dir)
    output_path = Path(output_dir)
    
    # Create output directory if it doesn't exist
    output_path.mkdir(parents=True, exist_ok=True)
    
    converted = 0
    errors = 0
    
    for eml_file in input_path.glob('*.eml'):
        try:
            print(f"\nProcessing: {eml_file.name}")
            success, result = convert_eml_to_pdf(eml_file, output_path)
            if success:
                print(f"Converted: {result}")
                converted += 1
            else:
                print(f"Error converting {eml_file.name}: {result}")
                errors += 1
        except Exception as e:
            print(f"Error processing {eml_file.name}: {str(e)}")
            errors += 1
    
    return converted, errors

def main():
    # Create root window and hide it
    root = tk.Tk()
    root.withdraw()
    
    # Get input directory
    input_dir = filedialog.askdirectory(title="Select Directory with EML Files")
    if not input_dir:
        print("No input directory selected")
        return
    
    # Get output directory
    output_dir = filedialog.askdirectory(title="Select Output Directory for PDFs")
    if not output_dir:
        print("No output directory selected")
        return
    
    print(f"\nProcessing files in: {input_dir}")
    print(f"Output directory: {output_dir}")
    
    converted, errors = process_directory(input_dir, output_dir)
    
    print(f"\nDone! Converted: {converted}, Errors: {errors}")
    messagebox.showinfo("Complete", f"Converted: {converted}\nErrors: {errors}")

if __name__ == "__main__":
    main() 