import re
from typing import Any
from pathlib import Path
import logging
from rich.logging import RichHandler
import shutil

logging.basicConfig(level=logging.INFO, handlers=[RichHandler()])
logger = logging.getLogger(__name__)

INPUT_DIR = 'input'
IMAGE_DIR = INPUT_DIR + '/Template'
OUTPUT_DIR = INPUT_DIR + '/CleanedTemplate'
OUTPUT_IMAGES_DIR = OUTPUT_DIR + '/Images'
REF_DIR = INPUT_DIR + '/Reference'
CLEANED_REF_DIR = OUTPUT_DIR + '/Reference'
Path(INPUT_DIR).mkdir(exist_ok=True)
Path(IMAGE_DIR).mkdir(exist_ok=True)
Path(OUTPUT_DIR).mkdir(exist_ok=True)
Path(OUTPUT_IMAGES_DIR).mkdir(exist_ok=True)
Path(REF_DIR).mkdir(exist_ok=True)
Path(CLEANED_REF_DIR).mkdir(exist_ok=True)


def read_template_file(file_path: Path) -> str:
    try:
        with file_path.open('r', encoding='utf-8') as file:
            return file.read()
    except FileNotFoundError:
        logger.error(f"File not found: {file_path}")
        return ""


def save_md_file(file_path: Path, data: str) -> None:
    try:
        output_file_path = Path(OUTPUT_DIR) / file_path.name
        with output_file_path.open('w', encoding='utf-8') as file:
            file.write(data)
        logger.info(f"Markdown file saved: {output_file_path}")
    except IOError as e:
        logger.error(f"Error writing file: {e}")


def process_markdown(data: str) -> str:
    start_marker = r"\s*# Table of Contents"
    end_marker = r"\n\n---"
    start_match = re.search(start_marker, data)
    end_match = re.search(end_marker, data[start_match.end():], re.MULTILINE) if start_match else None
    if start_match and end_match:
        logger.info("Found TOC")
        section_start = start_match.start()
        section_end = start_match.end() + end_match.end()
        section = data[section_start:section_end]
        pattern = r"\[.*?\]\(.*?\)"
        updated_section = re.sub(pattern, lambda match: match.group(0).split(']')[0] + ']', section)
        data = data[:section_start] + updated_section + data[section_end:]
    data = re.sub(r"(Template)%.*?/", r"Template/", data)
    data = re.sub(r"(Template)/", r"input/CleanedTemplate/Images/", data)
    data = re.sub(r"Untitled%20(\d+)", r"Untitled\1", data)  # New pattern to remove '%20' from untitled images
    data = re.sub(r"\[Untitled\]", r"[]", data)  # New pattern to remove '[Untitled]' to '[]'
    return data


def rename_all_images(directory) -> tuple[list[Any], list[Any]]:
    original_images = []
    new_images = []
    logger.info(f'rename_all_images Directory: {directory}')
    for file in Path(directory).iterdir():
        if file.suffix == '.png':
            original_images.append(file.name)
            new_images.append(file.stem.replace(' ', '') + file.suffix)
            file.rename(file.parent / new_images[-1])
    return original_images, new_images


def copy_all_pngs(input_folder, output_folder) -> None:
    logger.info(f'copy_all_pngs Directory: {input_folder}')
    for file in Path(input_folder).iterdir():
        if file.suffix == '.png':
            new_file_path = Path(output_folder) / file.name
            shutil.copy(file, new_file_path)
            # logger.info(f"Copied file {file} to {new_file_path}")
    return None


def copy_reference_folder(input_folder, output_folder) -> list[tuple[Any, Any]]:
    logger.info(f'copy_reference_folder Directory: {input_folder}')
    copied_files = []
    for file in Path(input_folder).iterdir():
        if file.is_dir():
            new_file_path = Path(output_folder) / file.name
            shutil.copytree(file, new_file_path, copy_function=shutil.copy2)
            logger.info(f"Copied folder {file} to {new_file_path}")
            copied_files.append((file, new_file_path))
        else:
            new_file_path = Path(output_folder) / file.name
            shutil.copy2(file, new_file_path)
            logger.info(f"Copied file {file} to {new_file_path}")
            copied_files.append((file, new_file_path))
    return copied_files


def main() -> None:
    input_dir = Path(INPUT_DIR)
    try:
        for file_path in input_dir.glob("*.md"):
            logger.info(f'Processing Markdown file: {file_path}')
            data = read_template_file(file_path)
            processed_data = process_markdown(data)
            cleaned_file_path = file_path.with_name(file_path.stem + "_cleaned.md")
            save_md_file(cleaned_file_path, processed_data)
        logger.info("Processing complete.")
    except Exception as e:
        logger.error(f"Error processing Markdown files: {e}")
        return None
    try:
        rename_all_images(IMAGE_DIR)
        copy_all_pngs(IMAGE_DIR, OUTPUT_IMAGES_DIR)
        logger.info(f"Moved filed files from {IMAGE_DIR} to {OUTPUT_IMAGES_DIR}")
    except Exception as e:
        logger.error(f"Error renaming images: {e}")
        return None
    try:
        copy_reference_folder(REF_DIR, CLEANED_REF_DIR)
        logger.info(f"Copied reference folder from {REF_DIR} to {CLEANED_REF_DIR}")
    except Exception as e:
        logger.error(f"Error copying reference folder: {e}")
        return None


if __name__ == '__main__':
    main()
