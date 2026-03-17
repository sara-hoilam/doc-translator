#!/usr/bin/env python3
"""
High-fidelity PDF → DOCX converter using pdf2docx.
Usage: python3 pdf2docx_convert.py <input_pdf_path> <output_docx_path>

Reads PDF from input_pdf_path, writes DOCX to output_docx_path.
Exits with code 0 on success, 1 on failure.
Writes progress/error messages to stderr.
"""
import sys
import os

def main():
    if len(sys.argv) != 3:
        print("Usage: pdf2docx_convert.py <input.pdf> <output.docx>", file=sys.stderr)
        sys.exit(1)

    input_path = sys.argv[1]
    output_path = sys.argv[2]

    if not os.path.exists(input_path):
        print(f"Input file not found: {input_path}", file=sys.stderr)
        sys.exit(1)

    try:
        from pdf2docx import Converter
        cv = Converter(input_path)
        cv.convert(output_path, multi_processing=False)
        cv.close()

        if not os.path.exists(output_path) or os.path.getsize(output_path) < 100:
            print("Conversion produced empty or missing output file", file=sys.stderr)
            sys.exit(1)

        print(f"OK:{os.path.getsize(output_path)}", flush=True)
        sys.exit(0)
    except Exception as e:
        print(f"Conversion failed: {e}", file=sys.stderr)
        sys.exit(1)

if __name__ == "__main__":
    main()
