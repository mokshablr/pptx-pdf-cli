#!/usr/bin/env python3

import os
import sys
import time
import subprocess
import glob
from PyPDF2 import PdfMerger
import concurrent.futures

LIBREOFFICE_PATH = "/Applications/LibreOffice.app/Contents/MacOS/soffice"

def convert_single(input_file, output_dir):
    if not os.path.exists(input_file):
        print(f"File '{input_file}' does not exist. Skipping.")
        return None
    try:
        subprocess.run(
            [LIBREOFFICE_PATH, "--headless", "--convert-to", "pdf", "--outdir", output_dir, input_file],
            check=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE
        )
        print(f"Converted: {input_file} -> {output_dir}/{input_file}")
        return input_file
    except subprocess.CalledProcessError as e:
        print(f"Error converting {input_file}: {e}")
        return None

def convert_pptx_to_pdf(input_files, output_dir):
    if isinstance(input_files, str):
        pptx_files = glob.glob(input_files)
    elif isinstance(input_files, list):
        pptx_files = []
        for item in input_files:
            if "*" in item or "?" in item:
                pptx_files.extend(glob.glob(item))
            else:
                pptx_files.append(item)
    else:
        pptx_files = []

    valid_files = []
    for file in pptx_files:
        if file.lower().endswith('.pptx'):
            valid_files.append(file)
        else:
            print(f"Invalid file type. Skipping.")

    if not valid_files:
        print("No valid .pptx files found.")
        sys.exit(1)

    print("PPTX files found:", valid_files)
    
    os.makedirs(output_dir, exist_ok=True)

    # Use ProcessPoolExecutor to convert files in parallel.
    with concurrent.futures.ProcessPoolExecutor(max_workers=4) as executor:
        futures = [executor.submit(convert_single, file, output_dir) for file in valid_files]
        # Wait for all futures to complete.
        for future in concurrent.futures.as_completed(futures):
            future.result()  # This will also re-raise any exceptions caught in the worker.

def merge_pdfs(pdf_patterns, output_file):
    pdf_files = []
    for item in pdf_patterns:
        if "*" in item or "?" in item:
            pdf_files.extend(glob.glob(item))
        else:
            pdf_files.append(item)
    
    valid_files = []
    for file in pdf_files:
        if file.lower().endswith('.pdf'):
            valid_files.append(file)
        else:
            print(f"Skipping file '{file}' - invalid file type (expected .pdf)")
    
    if len(valid_files) < 2:
        print("You need at least two valid PDF files to merge.")
        sys.exit(1)

    print("PDF files to merge:", valid_files)
    
    if not output_file.lower().endswith('.pdf'):
        print("Output file does not have a .pdf extension.")

    merger = PdfMerger()
    for pdf in valid_files:
        if not os.path.exists(pdf):
            print(f"File '{pdf}' does not exist. Skipping.")
            continue
        merger.append(pdf)
    
    try:
        merger.write(output_file)
        merger.close()
        print(f"PDFs merged successfully! Output: {output_file}")
    except Exception as e:
        print(f"Error during merging: {e}")
        sys.exit(1)

if len(sys.argv) < 4:
    print("Usage:")
    print("Convert PPTX to PDF: pp.py --convert <pptx_file(s) or pattern> <output_directory>")
    print("Merge PDFs: pp.py --merge <output_pdf> <input_pdf(s) or patterns>")
    sys.exit(1)

action = sys.argv[1]
if action == "--convert":
    start_time = time.time()

    input_args = sys.argv[2:-1]
    output_directory = sys.argv[-1]

    if len(input_args) == 1 and ("*" in input_args[0] or "?" in input_args[0]):
        convert_pptx_to_pdf(input_args[0], output_directory)
    else:
        convert_pptx_to_pdf(input_args, output_directory)
    end_time = time.time()
    elapsed_time = end_time - start_time
    print(f"Executed in {elapsed_time:.2f} seconds")

elif action == "--merge":
    start_time = time.time()
    output_pdf = sys.argv[2]
    input_pdfs = sys.argv[3:]
    merge_pdfs(input_pdfs, output_pdf)
    end_time = time.time()
    elapsed_time = end_time - start_time
    print(f"Executed in {elapsed_time:.2f} seconds")

else:
    print("Use '--convert' or '--merge'.")
    sys.exit(1)

