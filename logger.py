"""
Logger module for Universal AI Excel Format Converter.
Logs events, mappings, prompts, errors, and performance execution times to logs/application.log.
"""

import logging
import os
import sys
from typing import Optional, Dict, Any
import config


def setup_logger(name: str = "universal_excel_converter") -> logging.Logger:
    """Initializes application logger writing to console and logs/application.log."""
    os.makedirs(config.LOG_DIR, exist_ok=True)
    log_file_path = os.path.join(config.LOG_DIR, "application.log")

    logger = logging.getLogger(name)
    logger.setLevel(logging.DEBUG)

    if logger.hasHandlers():
        return logger

    formatter = logging.Formatter(
        fmt="[%(asctime)s] [%(levelname)s] [%(filename)s:%(lineno)d] - %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S"
    )

    # File handler
    file_handler = logging.FileHandler(log_file_path, encoding="utf-8")
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)

    # Console handler
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(logging.INFO)
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)

    return logger


logger = setup_logger()


def get_logger() -> logging.Logger:
    return logger


def log_transformation(
    filename: str,
    preset: str,
    prompt: str,
    input_rows: int,
    output_rows: int,
    elapsed_time: float
) -> None:
    """Logs transformation summary metrics."""
    logger.info(
        f"Transformation Completed: File='{filename}', Preset='{preset}', "
        f"Prompt='{prompt}', InputRows={input_rows}, OutputRows={output_rows}, "
        f"Elapsed={elapsed_time:.3f}s"
    )


def log_error(context_message: str, exc: Optional[Exception] = None) -> None:
    """Logs errors with traceback."""
    if exc:
        logger.error(f"{context_message}: {str(exc)}", exc_info=True)
    else:
        logger.error(context_message)
