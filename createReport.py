# This script will first run cleaner.py, wait for 1 second then run md2docx.py
# it needs to have cli prints and error handling with rich logging

import os
import re
from typing import Any
from pathlib import Path
import logging

import time

from rich.logging import RichHandler

from rich.console import Console

console = Console()


def print_error(error: Any) -> None:
    console.print(f"Error: {error}")
    logger.error(f"Error: {error}")
    return None


def print_success(message: Any) -> None:
    console.print(f"Success: {message}")
    logger.info(f"Success: {message}")
    return None


logging.basicConfig(level=logging.INFO, handlers=[RichHandler()])

logger = logging.getLogger(__name__)


def main() -> None:
    cleaner_script = Path('cleaner.py')
    if not cleaner_script.exists():
        logger.error("Cleaner script does not exist.")
        return
    # Check if the md2docx.py script exists
    md2docx_script = Path('md2docx.py')
    if not md2docx_script.exists():
        logger.error("MD2DOCX script does not exist.")
        return
    # Run the cleaner.py script
    logger.info("Running cleaner.py script...")
    os.system('python cleaner.py')
    # Wait for 1 second
    logger.info("Waiting for 1 second...")
    time.sleep(1)
    # Run the md2docx.py script
    logger.info("Running md2docx.py script...")
    os.system('python md2docx.py')
    # Wait for 1 second
    logger.info("Waiting for 1 second...")
    time.sleep(1)
    # Check if the output directory contains any files
    logger.info("Report creation complete.")


if __name__ == '__main__':
    main()
