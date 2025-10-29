# TIFF and PDF Metadata Extractor

A comprehensive Python GUI application for extracting and analyzing metadata from TIFF images and PDF files. Export results to Excel with detailed file information and validation.

![Python](https://img.shields.io/badge/Python-3.7%2B-blue)
![License](https://img.shields.io/badge/License-MIT-green)
![Platform](https://img.shields.io/badge/Platform-Windows%20%7C%20Mac%20%7C%20Linux-lightgrey)

## 📋 Table of Contents

- [Features](#features)
- [Screenshots](#screenshots)
- [Installation](#installation)
- [Usage](#usage)
- [Supported Formats](#supported-formats)
- [Extracted Metadata](#extracted-metadata)
- [Filename Validation](#filename-validation)
- [Export Options](#export-options)
- [Requirements](#requirements)
- [Troubleshooting](#troubleshooting)
- [Contributing](#contributing)
- [License](#license)
- [Support](#support)

## ✨ Features

### TIFF File Processing
- **Complete Metadata Extraction**: Format, extension, DPI, compression, color depth
- **Batch Processing**: Analyze multiple TIFF files simultaneously
- **EXIF Data Reading**: Extract embedded metadata and resolution information
- **Color Space Detection**: Identify grayscale, RGB, CMYK, and other color modes

### PDF File Processing
- **Per-Page Analysis**: Extract metadata from each page individually
- **Image Content Detection**: Identify pages containing images vs. text
- **Embedded Image Analysis**: Extract DPI, color depth, and compression from PDF images
- **Multi-Page Support**: Process PDFs with any number of pages

### User Interface
- **Intuitive GUI**: Easy-to-use graphical interface built with Tkinter
- **Real-time Results**: View extracted metadata in sortable tables
- **Progress Tracking**: Visual progress bar for batch operations
- **Error Handling**: Graceful error handling with detailed messages

### Data Management
- **Excel Export**: Save all metadata to structured Excel spreadsheets
- **Filename Validation**: Automatic validation against ISBN13 naming conventions
- **Batch Export**: Export results from multiple files in one operation

## 🖼️ Screenshots

### Main Interface
[TIFF and PDF Metadata Extractor]
├── File Type Selection (TIFF/PDF)
├── File Selection Button
├── Extract Metadata Button
├── Export to Excel Button
├── Progress Bar
└── Results Table with Sortable Columns



### TIFF Analysis View
Filename Format Extension DPI Compression Color Depth Valid
9781234567890_00001.tif TIFF .tif 300 x 300 LZW 24-bit Color Yes
9781234567890_00002.tif TIFF .tif 300 x 300 JPEG 8-bit Grayscale Yes



### PDF Analysis View
Filename Page Type Color Depth DPI Compression Valid
9781234567890.pdf 1 PDF Image 24-bit Color 300 x 300 Flate, JPEG Yes
9781234567890.pdf 2 PDF Text - - - Yes



## 🚀 Installation

### Prerequisites
- Python 3.7 or higher
- pip (Python package manager)

### Method 1: Standard Installation

1. **Clone the repository**:
   ```bash
   git clone https://github.com/yourusername/tiff-pdf-metadata-extractor.git
   cd tiff-pdf-metadata-extractor
Install required dependencies:

bash
pip install -r requirements.txt
Method 2: Manual Dependency Installation
If you prefer to install dependencies manually:

bash
pip install Pillow>=10.0.0 pandas>=2.0.0 openpyxl>=3.0.0 PyPDF2>=3.0.0
Method 3: Virtual Environment (Recommended)
For isolated installation:

bash
# Create virtual environment
python -m venv metadata_env

# Activate virtual environment
# On Windows:
metadata_env\Scripts\activate
# On Mac/Linux:
source metadata_env/bin/activate

# Install dependencies
pip install -r requirements.txt
📖 Usage
Starting the Application
bash
python metadata_extractor.py
Step-by-Step Guide
Select File Type

Choose between TIFF Files or PDF Files using the radio buttons

Select Files

Click "Select Files" button

Choose one or multiple files in the file dialog

Supported formats: .tif, .tiff for TIFF; .pdf for PDF

Extract Metadata

Click "Extract Metadata" button

View progress in the progress bar

Results display in the sortable table

Export Results

Click "Export to Excel" button

Choose save location and filename

Data exports to structured Excel format

Command Line Usage (Advanced)
For batch processing, you can modify the script for command-line use:

python
# Example for advanced users
extractor = MetadataExtractor()
# Add custom file processing logic here
📁 Supported Formats
TIFF Files
Extensions: .tif, .tiff

Compression: LZW, JPEG, Deflate, PackBits, CCITT, Uncompressed

Color Modes: 1-bit, Grayscale, RGB, RGBA, CMYK, LAB, HSV

Bit Depths: 1-bit to 32-bit

PDF Files
Extension: .pdf

Content Types: Image-based PDFs, Text PDFs, Mixed content

Image Compression: Flate, JPEG, JPEG2000, CCITT, LZW

Color Spaces: DeviceRGB, DeviceGray, DeviceCMYK, Indexed

🔍 Extracted Metadata
TIFF File Metadata
Field	Description	Example
Image Format	File format type	TIFF, TIFF
Extension	File extension	.tif, .tiff
DPI	Resolution in dots per inch	300 x 300, 600 x 600
Compression	Compression algorithm	LZW, JPEG, Deflate
Color Depth	Bits per pixel and color mode	24-bit Color, 8-bit Grayscale
Filename Valid	Conforms to naming convention	Yes/No/Wrong extension
PDF File Metadata (Per Page)
Field	Description	Example
Page Number	Page index (1-based)	1, 2, 3...
Type	Content type on page	PDF Image, PDF Text
Color Depth	Color information	24-bit Color, 8-bit Grayscale
DPI	Image resolution	300 x 300, 150 x 150
Compression	Compression filters	Flate, JPEG, CCITT
Filename Valid	Conforms to naming convention	Yes/No/Wrong extension
📛 Filename Validation
The tool automatically validates filenames against industry-standard conventions:

TIFF Filename Pattern

ISBN13_#####.tif
ISBN13: 13-digit International Standard Book Number

#####: 5-digit sequence number (00001-99999)

Extension: Must be .tif

Examples:

✅ Valid: 9780123456789_00001.tif

✅ Valid: 9780123456789_12345.tif

❌ Invalid: 9780123456789_0001.tif (4-digit sequence)

❌ Invalid: 9780123456789_00001.tiff (wrong extension)

PDF Filename Pattern

ISBN13.pdf
ISBN13: 13-digit International Standard Book Number

Extension: Must be .pdf

Examples:

✅ Valid: 9780123456789.pdf

❌ Invalid: 9780123456789.PDF (wrong case)

❌ Invalid: 9780123456789_.pdf (extra characters)

💾 Export Options
Excel Export Features
Structured Worksheets: Separate sheets for different file types

Complete Metadata: All extracted fields included

Formatted Columns: Professional spreadsheet formatting

Multiple Files: Batch export from multiple analysis sessions

Error Tracking: Failed files marked with error information

Export Columns
TIFF Export:

File Type, Filename, Format, Extension, DPI, Compression, Color Depth, Filename Valid, Full Path

PDF Export:

File Type, Filename, Page, Type, Color Depth, DPI, Compression, Filename Valid, Full Path

📋 Requirements
Python Packages
Pillow (≥10.0.0): Image processing and TIFF metadata extraction

pandas (≥2.0.0): Data manipulation and Excel export

openpyxl (≥3.0.0): Excel file creation and formatting

PyPDF2 (≥3.0.0): PDF metadata extraction and analysis

System Requirements
Operating System: Windows 10+, macOS 10.14+, or Linux

Python Version: 3.7 or higher

RAM: 4GB minimum, 8GB recommended for large files

Storage: 100MB free space for application and dependencies

🛠️ Troubleshooting
Common Issues
1. "ModuleNotFoundError"

bash
# Reinstall requirements
pip install --force-reinstall -r requirements.txt
2. TIFF Files Not Loading

Ensure files are not corrupted

Check file permissions

Verify TIFF file format compatibility

3. PDF Analysis Slow

Large PDFs may take longer to process

Consider processing in smaller batches

Close other memory-intensive applications

4. Excel Export Fails

Ensure no Excel file with same name is open

Check write permissions in target directory

Verify sufficient disk space

Performance Tips
Batch Processing: Process files in batches of 50-100 for optimal performance

Memory Management: Close application between large batch operations

File Organization: Group files by type for faster processing

🤝 Contributing
We welcome contributions! Please see our contributing guidelines:

How to Contribute
Fork the repository

Create a feature branch: git checkout -b feature/amazing-feature

Commit your changes: git commit -m 'Add amazing feature'

Push to the branch: git push origin feature/amazing-feature

Open a Pull Request

Development Setup
bash
# Clone and setup
git clone https://github.com/yourusername/tiff-pdf-metadata-extractor.git
cd tiff-pdf-metadata-extractor

# Create development environment
python -m venv dev_env
source dev_env/bin/activate  # Windows: dev_env\Scripts\activate

# Install development dependencies
pip install -r requirements.txt
Feature Requests
Please use GitHub Issues to:

Report bugs

Request new features

Suggest improvements

📄 License
This project is licensed under the MIT License - see the LICENSE file for details.

MIT License Features:

✅ Free to use for personal and commercial projects

✅ Permission to modify and distribute

✅ No requirement to open-source modifications

✅ No warranty or liability

🆘 Support
Documentation
This README

In-application tooltips and status messages

Code comments in source files

Getting Help
Check Existing Issues: Search GitHub Issues for similar problems

Create New Issue: Provide detailed description and error messages

Email Support: Contact maintainer for direct support

Community
Star the repository if you find it useful!

Share your use cases in GitHub Discussions

Follow updates by watching the repository

🎯 Use Cases
Publishing Industry
Quality control for book production

Verification of image specifications

Batch metadata extraction for large projects

Digital Archives
Preservation metadata collection

Format validation and standardization

Migration planning and assessment

Graphic Design
Asset management and organization

Client delivery verification

Print preparation checks

Software Development
Automated testing of file processing pipelines

Quality assurance for file generation systems

Batch analysis of generated content

⭐ If this tool helped you, please consider giving it a star on GitHub!
