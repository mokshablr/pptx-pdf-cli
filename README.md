# PPTX to PDF Converter & PDF Merger (CLI Tool)

A simple command-line tool to convert PowerPoint (.pptx) files to PDF and merge multiple PDFs using **LibreOffice** and **PyPDF2**.  


## Features

✅ Convert multiple `.pptx` files to `.pdf` in parallel using LibreOffice  
✅ Merge multiple PDF files into a single PDF  
✅ Supports glob patterns for batch processing

## Prerequisites  
- **Python 3**  
- **LibreOffice** (Installed on macOS)  
    ```
    brew install --cask libreoffice
    ```
- **Dependencies** (Install using pip):  
    ```
    pip install pypdf2
    ```
    
## Installation

1. Clone this repository:
    ```
    git clone https://github.com/mokshablr/pptx-pdf-cli.git
    cd pptx-pdf-cli
    ```

2. Make the script executable:
    ```
    chmod +x pdf.py
    ```

3. Move it to /usr/local/bin to run it from anywhere:
    ```
    sudo mv pdf.py /usr/local/bin/pdf
    ```

## Usage/Examples

### Convert PPTX to PDF

    pdf --convert ppt1.pptx ppt2.pptx ./pdfs
    pdf --convert *.pptx ./pdfs

### Merge PDFs

    pdf --merge output.pdf file1.pdf file2.pdf
    pdf --merge output.pdf *.pdf


