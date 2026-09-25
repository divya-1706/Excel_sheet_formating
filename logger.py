"""
Logger module for OTIS Excel Report Formatter.
Sets up logging to both console and logs/application.log file.
"""

import logging
import os
import sys
from typing import Optional
import config


def setup_logger(name: str = "otis_formatter") -> logging.Logger:
    """
    Initializes and configures the logger.
    Logs are written to both logs/application.log and sys.stdout.
    """
    # Ensure logs directory exists
    os.makedirs(config.LOG_DIR, exist_ok=True)
    log_file_path = os.path.join(config.LOG_DIR, "application.log")

    logger = logging.getLogger(name)
    logger.setLevel(logging.DEBUG)

    # Avoid duplicate handlers if logger is already configured
    if logger.hasHandlers():
        return logger

    # Formatter for log lines
    formatter = logging.Formatter(
        fmt="[%(asctime)s] [%(levelname)s] [%(filename)s:%(lineno)d] - %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S"
    )

    # File Handler
    file_handler = logging.FileHandler(log_file_path, encoding="utf-8")
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)

    # Console Handler
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(logging.INFO)
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)

    return logger


# Default global logger instance
logger = setup_logger()


def get_logger() -> logging.Logger:
    """Returns the configured application logger instance."""
    return logger


def log_upload(filename: str, file_size: Optional[int] = None) -> None:
    """Logs details when a user uploads a new Excel file."""
    size_str = f" ({file_size} bytes)" if file_size is not None else ""
    logger.info(f"File uploaded: '{filename}'{size_str}")


def log_completion(filename: str, total_reports: int, output_path: str, elapsed_time: float) -> None:
    """Logs completion status after processing a report."""
    logger.info(
        f"Processing complete for '{filename}'. "
        f"Processed {total_reports} report blocks into '{output_path}' "
        f"in {elapsed_time:.2f} seconds."
    )


def log_error(context_message: str, exc: Optional[Exception] = None) -> None:
    """Logs an error with optional exception traceback."""
    if exc:
        logger.error(f"{context_message}: {str(exc)}", exc_info=True)
    else:
        logger.error(context_message)
