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
    """Converts a single PPTX file to PDF using LibreOffice."""
    if not os.path.exists(input_file):
        print(f"❌ Error: File '{input_file}' does not exist. Skipping.")
        return None
    try:
        result = subprocess.run(
            [LIBREOFFICE_PATH, "--headless", "--convert-to", "pdf", "--outdir", output_dir, input_file],
            check=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE
        )
        print(f"✅ Converted: {input_file} → {output_dir}")
        return input_file
    except subprocess.CalledProcessError as e:
        print(f"❌ Error converting {input_file}: {e}")
        return None

def convert_pptx_to_pdf(input_files, output_dir, parallel=False):
    """
    Converts multiple .pptx files to .pdf using LibreOffice.
    Accepts either a glob pattern (string) or an explicit list of file paths.
    Only files with a .pptx extension are processed.
    Runs sequentially by default; use parallel=True for parallel conversion.
    """
    # Expand glob patterns if needed.
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

    # Filter for valid .pptx files.
    valid_files = []
    for file in pptx_files:
        if file.lower().endswith('.pptx'):
            valid_files.append(file)
        else:
            print(f"❌ Skipping file '{file}' - invalid file type (expected .pptx)")

    if not valid_files:
        print("❌ No valid .pptx files found.")
        sys.exit(1)

    print("Valid PPTX files found:", valid_files)
    
    # Ensure output directory exists.
    os.makedirs(output_dir, exist_ok=True)

    if parallel:
        # Use ProcessPoolExecutor to convert files in parallel.
        with concurrent.futures.ProcessPoolExecutor(max_workers=4) as executor:
            futures = [executor.submit(convert_single, file, output_dir) for file in valid_files]
            # Wait for all futures to complete.
            for future in concurrent.futures.as_completed(futures):
                future.result()  # Re-raises any exceptions caught in the worker.
    else:
        # Sequential processing.
        for file in valid_files:
            convert_single(file, output_dir)

def parse_range_input(input_str, max_len):
    """
    Parses a comma-separated input string that may contain individual numbers or ranges.
    Ranges can be specified with a hyphen (-) or colon (:), e.g. "1-4" or "1:4".
    Returns a list of 0-indexed positions.
    """
    indices = []
    tokens = input_str.split(',')
    for token in tokens:
        token = token.strip()
        if '-' in token or ':' in token:
            delim = '-' if '-' in token else ':'
            parts = token.split(delim)
            if len(parts) != 2:
                raise ValueError("Invalid range token: " + token)
            try:
                start = int(parts[0])
                end = int(parts[1])
            except ValueError:
                raise ValueError("Non-integer range bounds in: " + token)
            if start > end:
                raise ValueError("Range start must not be greater than end in: " + token)
            # Add all numbers in the range (inclusive)
            for i in range(start, end + 1):
                indices.append(i - 1)  # Convert to 0-indexed
        else:
            try:
                indices.append(int(token) - 1)
            except ValueError:
                raise ValueError("Invalid token, not an integer: " + token)
    # Validate indices.
    if any(idx < 0 or idx >= max_len for idx in indices):
        raise ValueError("One or more indices are out of valid range.")
    return indices

def interactive_merge_order(valid_files):
    """
    Presents the list of files to the user and asks for the desired merge order.
    The user can enter a comma-separated list of numbers and/or ranges (e.g. "1,3,5-7").
    If the input is empty or invalid, the default order is used.
    """
    print("Files available for merging:")
    for i, file in enumerate(valid_files, 1):
        print(f"  {i}: {file}")
    order_input = input("Enter the numbers in desired order (use '-' or ':' for ranges, e.g. '1,3,5-7') or press Enter for default order: ").strip()
    if not order_input:
        print("Using default order.")
        return valid_files
    try:
        indices = parse_range_input(order_input, len(valid_files))
        ordered_files = [valid_files[idx] for idx in indices]
        return ordered_files
    except ValueError as ve:
        print(f"❌ {ve}. Using default order.")
        return valid_files

def merge_pdfs(pdf_patterns, output_file):
    """
    Merges multiple PDFs into one.
    Accepts explicit filenames or glob patterns.
    Only files with a .pdf extension are processed.
    Provides an interactive interface to select the merge order.
    """
    pdf_files = []
    for item in pdf_patterns:
        if "*" in item or "?" in item:
            pdf_files.extend(glob.glob(item))
        else:
            pdf_files.append(item)
    
    # Filter for valid .pdf files.
    valid_files = []
    for file in pdf_files:
        if file.lower().endswith('.pdf'):
            valid_files.append(file)
        else:
            print(f"❌ Skipping file '{file}' - invalid file type (expected .pdf)")
    
    if len(valid_files) < 2:
        print("❌ Error: You need at least two valid PDF files to merge.")
        sys.exit(1)

    print("PDF files found:")
    print(valid_files)
    
    # Ask the user for the desired order interactively.
    ordered_files = interactive_merge_order(valid_files)
    print("Merging files in the following order:")
    for file in ordered_files:
        print(f"  {file}")
    
    if not output_file.lower().endswith('.pdf'):
        print("❌ Warning: Output file does not have a .pdf extension.")

    # Start the timer after the interactive order is set.
    merge_start_time = time.time()
    
    merger = PdfMerger()
    for pdf in ordered_files:
        if not os.path.exists(pdf):
            print(f"❌ Error: File '{pdf}' does not exist. Skipping.")
            continue
        merger.append(pdf)
    
    try:
        merger.write(output_file)
        merger.close()
        merge_end_time = time.time()
        merge_elapsed = merge_end_time - merge_start_time
        print(f"✅ PDFs merged successfully! Output: {output_file}")
        print(f"Total merge time: {merge_elapsed:.2f} seconds")
    except Exception as e:
        print(f"❌ Error during merging: {e}")
        sys.exit(1)

if __name__ == "__main__":
    # Minimum arguments: script, action, at least one input file/pattern, and an output directory (or output PDF)
    if len(sys.argv) < 4:
        print("Usage:")
        print("  Convert PPTX to PDF (sequential): pp.py --convert <pptx_file(s) or pattern> <output_directory>")
        print("  Convert PPTX to PDF (parallel):   pp.py --convert --parallel <pptx_file(s) or pattern> <output_directory>")
        print("  Merge PDFs: pp.py --merge <output_pdf> <input_pdf(s) or patterns>")
        sys.exit(1)

    action = sys.argv[1]
    if action == "--convert":
        start_time = time.time()
        parallel_mode = False
        args_offset = 2
        if sys.argv[2] == "--parallel":
            parallel_mode = True
            args_offset = 3
        # All arguments after the flag: all except the last are input file(s)/pattern, last is output directory.
        input_args = sys.argv[args_offset:-1]
        output_directory = sys.argv[-1]
        # If there's exactly one argument and it contains a wildcard, treat it as a pattern.
        if len(input_args) == 1 and ("*" in input_args[0] or "?" in input_args[0]):
            convert_pptx_to_pdf(input_args[0], output_directory, parallel=parallel_mode)
        else:
            convert_pptx_to_pdf(input_args, output_directory, parallel=parallel_mode)
        end_time = time.time()
        elapsed_time = end_time - start_time
        print(f"Total conversion time: {elapsed_time:.2f} seconds")

    elif action == "--merge":
        output_pdf = sys.argv[2]
        input_pdfs = sys.argv[3:]
        merge_pdfs(input_pdfs, output_pdf)

    else:
        print("❌ Error: Unknown action. Use '--convert' or '--merge'.")
        sys.exit(1)

