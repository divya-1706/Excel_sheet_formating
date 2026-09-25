"""
Command-line runner script for OTIS Excel Report Formatter.
Allows converting OTIS Test Sequence Excel files directly from terminal.
"""

import os
import sys
import argparse

# Ensure base directory in sys.path
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

import config
from formatter import convert_otis_report
from utils import ensure_directories, format_output_filename
from logger import logger


def main():
    ensure_directories()

    parser = argparse.ArgumentParser(description="Convert OTIS Test Sequence Excel file into formatted OTIS Inspection Report.")
    parser.add_argument("input_file", nargs="?", default=os.path.join(config.SAMPLE_DIR, "sample_test_sequence.xlsx"), help="Path to input Excel file")
    parser.add_argument("-o", "--output", help="Path to output Excel file")

    args = parser.parse_args()

    input_file = args.input_file
    if not os.path.exists(input_file):
        filename = os.path.basename(input_file)
        candidates = [
            os.path.join(config.INPUT_DIR, filename),
            os.path.join(config.SAMPLE_DIR, filename),
            os.path.join(BASE_DIR, filename),
        ]
        found = False
        for c in candidates:
            if os.path.exists(c):
                input_file = c
                found = True
                break
        if not found:
            print(f"Error: Input file '{args.input_file}' does not exist.")
            print(f"\nPlease provide a valid file path or copy your file into the 'input' folder:\n  '{config.INPUT_DIR}'")
            sys.exit(1)

    output_file = args.output
    if not output_file:
        output_file = os.path.join(config.OUTPUT_DIR, format_output_filename(os.path.basename(input_file)))

    print(f"Converting OTIS Test Sequence file: '{input_file}'...")
    try:
        saved_path, blocks, summary = convert_otis_report(input_file, output_file)
        print("\n" + "=" * 50)
        print("CONVERSION SUCCESSFUL!")
        print("=" * 50)
        print(f"Total Reports Processed : {summary['total_reports']}")
        print(f"Total Rows Written      : {summary['total_rows_written']}")
        print(f"Processing Time (s)     : {summary['processing_time_seconds']}s")
        print(f"Done! Output saved as   : {saved_path}")
        print("=" * 50)
    except Exception as e:
        print(f"Error during conversion: {e}")
        logger.error("CLI conversion failed", exc_info=True)
        sys.exit(1)


if __name__ == "__main__":
    main()
