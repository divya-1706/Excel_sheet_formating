"""
Formatter module for OTIS Excel Report Formatter.
Serves as the high-level orchestration interface connecting input parsing,
data structuring, openpyxl excel generation, and output file saving.
"""

import os
import time
import io
from typing import Union, Optional, Tuple, List, Dict, Any

import config
from logger import logger, log_completion, log_error
from utils import ensure_directories, format_output_filename
from parser import parse_test_sequence_workbook, ReportBlock
from excel_writer import build_otis_inspection_workbook, save_workbook


def convert_otis_report(
    input_file: Union[str, io.BytesIO, bytes],
    output_file_path: Optional[str] = None,
    original_filename: str = "Test_Sequence.xlsx"
) -> Tuple[str, List[ReportBlock], Dict[str, Any]]:
    """
    Orchestrates full OTIS report conversion workflow:
    1. Validates inputs & ensures directories
    2. Parses input workbook into ReportBlock objects
    3. Builds formatted openpyxl inspection report workbook
    4. Saves output workbook
    5. Returns output path, report blocks, and summary statistics dict.
    """
    start_time = time.time()
    ensure_directories()

    logger.info("Starting OTIS report conversion workflow...")

    if isinstance(input_file, str):
        if not os.path.exists(input_file):
            raise FileNotFoundError(f"Input file not found: '{input_file}'")
        original_filename = os.path.basename(input_file)

    # Determine output filepath
    if not output_file_path:
        out_filename = format_output_filename(original_filename)
        output_file_path = os.path.join(config.OUTPUT_DIR, out_filename)

    try:
        # Step 1: Parse input file
        report_blocks, warnings, parse_summary = parse_test_sequence_workbook(input_file)

        if not report_blocks:
            raise ValueError(
                "No valid report blocks (Sl.No 1.1, 1.2, etc.) were found in the uploaded workbook."
            )

        # Step 2: Build formatted openpyxl workbook
        wb = build_otis_inspection_workbook(report_blocks)

        # Step 3: Save workbook
        saved_path = save_workbook(wb, output_file_path)

        elapsed_time = time.time() - start_time

        # Calculate statistics
        total_reports = len(report_blocks)
        total_rows_written = (total_reports * 5) + 1  # 5 rows per block + 1 header
        voltage_rows_count = sum(1 for b in report_blocks if b.voltage is not None)
        current_rows_count = sum(1 for b in report_blocks if b.current is not None)
        dmm_rows_count = total_reports

        summary_stats = {
            "total_reports": total_reports,
            "total_rows_written": total_rows_written,
            "voltage_rows": voltage_rows_count,
            "current_rows": current_rows_count,
            "dmm_rows": dmm_rows_count,
            "processing_time_seconds": round(elapsed_time, 3),
            "output_file_path": saved_path,
            "warnings": warnings,
            "original_filename": original_filename
        }

        log_completion(
            filename=original_filename,
            total_reports=total_reports,
            output_path=saved_path,
            elapsed_time=elapsed_time
        )

        return saved_path, report_blocks, summary_stats

    except Exception as e:
        log_error(f"Error during OTIS report conversion for '{original_filename}'", e)
        raise e
