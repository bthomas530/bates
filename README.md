# Enhanced Bates Numbering Utility

A powerful and user-friendly utility for automatically applying Bates numbers to PDF documents, with intelligent date extraction and comprehensive reporting. This project builds upon and enhances several existing Bates numbering utilities.

## Acknowledgments

This project incorporates ideas and improvements from several open-source Bates numbering utilities:

- [auto-bates-main](https://github.com/username/auto-bates-main)
- [bates-main](https://github.com/username/bates-main)
- [bates_numbering-main](https://github.com/username/bates_numbering-main)
- [pliny_the_stamper-master](https://github.com/goingforbrooke/pliny_the_stamper)

## Features

### Core Functionality
- Automatically applies Bates numbers to all files
- Converts non-PDF files to PDF format
- Maintains original directory structure
- Preserves original files (creates copies with Bates numbers)
- Supports custom prefixes for Bates numbers
- Configurable number of digits (default: 5)
- Customizable starting number
- Links original files with their PDF versions through Bates numbers
- Handles secured PDFs and encrypted documents

### PDF Conversion
- Automatically converts supported file types to PDF:
  - Excel spreadsheets (.xlsx, .xlsm, .xls)
  - Images (.png, .jpg, .jpeg, .gif, .bmp)
  - Email files (.eml)
- Direct Excel to PDF conversion without external dependencies
- High-quality image conversion with proper scaling and margins
- Maintains original file alongside PDF version
- Skips unsupported file types (e.g., .zip)

### Date Extraction
- Automatically extracts dates from PDF content and metadata
- Supports multiple date formats:
  - Standard numeric formats (MM/DD/YYYY, YYYY/MM/DD)
  - Full month names (e.g., "December 15, 2023")
  - Month abbreviations (e.g., "Dec 15, 2023")
  - ISO format (YYYY-MM-DD)
  - European format (DD.MM.YYYY)
  - Military format (DD MMM YYYY)
  - Dates with time (e.g., "12/15/2023 14:30")
  - Dates with timezone (e.g., "12/15/2023 14:30 EST")
- Falls back to file creation date if no date is found in content
- Validates dates to ensure they're not in the future

### Document Organization
- Maintains original folder structure
- Adds Bates number ranges to folder names
- Creates organized output with clear naming conventions
- Handles duplicate files and aliases appropriately

### Reporting
- Generates comprehensive Excel report including:
  - Bates numbers
  - Original filenames
  - PDF filenames
  - File types
  - Page counts
  - Extracted dates
  - Creation dates
  - File paths
  - Tree-like folder structure visualization
- Summary statistics including:
  - Total folders
  - Total files
  - Total pages
  - File type distribution
- Detailed logging of the process

### User Interface
- Modern, intuitive GUI interface with drag-and-drop support
- Configurable settings:
  - Prefix (default: "ABC")
  - Number of digits (default: 5)
  - Starting number (default: 1)
  - Stamp appearance:
    - Color (black, red, blue, green, gray)
    - Box width
    - Opacity (0-100%)
    - Position (9-position grid)
    - X/Y offset
- Progress tracking through detailed logging
- Quick access buttons for:
  - Starting the process
  - Opening log file
- Support for default directory configuration

## Installation

1. Clone the repository:
```bash
git clone [repository-url]
cd bates
```

2. Install the required dependencies:
```bash
pip install -r requirements.txt
```

## Usage

### GUI Mode
Run the program without arguments to start in GUI mode:
```bash
python bates_enhanced_v02.py
```

1. Drop a directory or file into the GUI or use the browse button
2. Configure settings (optional):
   - Prefix (default: "ABC")
   - Number of digits (default: 5)
   - Starting number (default: 1)
   - Stamp appearance settings
3. Click "Stamp Directory" or "Stamp File" to begin processing
4. Monitor progress through the log

### Command Line Mode
Run the program with arguments for command-line operation:
```bash
python bates_enhanced_v02.py input_directory output_directory [options]
```

Options:
- `--prefix`: Prefix for Bates numbers (default: '')
- `--zero-pad`: Number of digits for Bates numbers (default: 5)
- `--start`: Starting number for Bates numbering (default: 1)

Example:
```bash
python bates_enhanced_v02.py /path/to/files /path/to/output --prefix "ABC" --zero-pad 5 --start 1
```

## Output Structure

The utility creates the following structure in the output directory:

```
BATES_[input_directory]/
├── bates_index.xlsx         # Comprehensive Excel report
├── bates_process.log        # Detailed processing log
└── [maintained directory structure with Bates ranges]
    ├── ABC00001-ABC00010_FolderName/
    │   ├── ABC00001_document.pdf
    │   └── ABC00002_spreadsheet.pdf
    └── ABC00011-ABC00020_AnotherFolder/
        ├── ABC00011_image.pdf
        └── ABC00012_document.pdf
```

### Excel Report Contents
- Bates Number
- Folder Structure (tree view)
- Original Filename
- PDF Filename
- File Type
- Page Count
- Extracted Date
- Creation Date
- Date Added
- Original File Path
- PDF File Path

## Error Handling

The utility includes comprehensive error handling:
- Handles secured and encrypted PDFs
- Recovers gracefully from conversion errors
- Provides detailed error logging
- Continues processing even when individual files fail
- Falls back to file copy when stamping fails
- Creates "_FILES WITH ISSUES" directory for problematic files

## Requirements

- Python 3.6 or higher
- PyPDF4 (for PDF processing)
- tqdm (for progress bars)
- pathlib (for file system operations)
- tkinter (for GUI)
- python-docx (for Word document handling)
- unoconv (for document conversion on non-Windows systems)
- pywin32 and comtypes (for Windows-specific operations)
- PyPDF2 (for PDF processing)
- reportlab (for PDF generation)
- tkinterdnd2 (for drag-and-drop support)
- openpyxl (for Excel report generation)
- fpdf2 (for PDF generation)

## Notes

- The utility preserves original files by creating copies with Bates numbers
- Date extraction checks both PDF content and metadata
- The utility maintains the original directory structure in the output location
- All operations are logged for troubleshooting and auditing
- Handles secured PDFs by creating unsecured copies for stamping
- Provides fallback mechanisms for problematic files
- Supports both directory and single file processing modes

## Support

For issues, feature requests, or questions, please open an issue in the repository.

## License

[Add your license information here]
 
