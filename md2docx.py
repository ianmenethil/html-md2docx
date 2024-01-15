import subprocess
from pathlib import Path
import logging
from typing import Any
from rich.logging import RichHandler
from rich.traceback import install
from docx import Document
from docx.oxml import OxmlElement, parse_xml
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.shape import WD_INLINE_SHAPE
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Pt, RGBColor, Cm, Mm
from docx.oxml.ns import nsdecls, qn

install()

logging.basicConfig(level=logging.INFO,
                    format="%(message)s",
                    handlers=[
                        RichHandler(level=logging.INFO,
                                    show_time=True,
                                    show_path=True,
                                    show_level=True,
                                    rich_tracebacks=True,
                                    tracebacks_extra_lines=0,
                                    tracebacks_show_locals=False)
                    ])
logger = logging.getLogger(__name__)

INPUT_DIR = "input/CleanedTemplate"
OUTPUT_DIR = "final_output"
REFERENCE_DIR = INPUT_DIR + '/Reference'
REFERENCE_DOC = REFERENCE_DIR + '/refdoc.docx'
Path(OUTPUT_DIR).mkdir(exist_ok=True)
Path(INPUT_DIR).mkdir(exist_ok=True)
Path(REFERENCE_DIR).mkdir(exist_ok=True)
TOP_MARGIN = Cm(1)
BOTTOM_MARGIN = Cm(1)
LEFT_MARGIN = Cm(1)
RIGHT_MARGIN = Cm(1)


def convert_md_to_docx(file_path) -> Any:
    output_file = Path(OUTPUT_DIR) / (file_path.stem + ".docx")
    # pandoc_command = ["pandoc", str(file_path), "--reference-doc", REFERENCE_DOC, "-o", str(output_file)]
    pandoc_command = ["pandoc", str(file_path), "-o", str(output_file)]

    try:
        subprocess.run(pandoc_command, check=True)
        logger.info(f"Successfully converted {file_path} to {output_file}")
        return output_file
    except subprocess.CalledProcessError as e:
        logger.error(f"Error converting {file_path} to docx: {e}")
        return file_path


def style_table(table, header_fill, header_font_color, content_fill, content_font_color) -> None:
    for row in table.rows:
        for cell in row.cells:
            paragraphs = cell.paragraphs
            for paragraph in paragraphs:
                for run in paragraph.runs:
                    run.font.size = Pt(10)
                    run.font.name = 'Open Sans'
                paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                paragraph.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            set_cell_background_color(cell, header_fill if row == table.rows[0] else content_fill)
            set_font_color(cell, header_font_color if row == table.rows[0] else content_font_color)


def set_font_color(cell, font_color) -> None:
    for paragraph in cell.paragraphs:
        for run in paragraph.runs:
            run.font.color.rgb = font_color


def set_cell_background_color(cell, color_str) -> None:
    shading_elm = OxmlElement('w:shd')
    shading_elm.set(qn('w:fill'), color_str)
    cell._tc.get_or_add_tcPr().append(shading_elm)


def set_cell_borders(cell) -> None:
    tc = cell._element
    tcPr = tc.get_or_add_tcPr()
    for border in ["top", "left", "bottom", "right"]:
        border_elm = OxmlElement(f'w:{border}')
        border_elm.set(qn('w:val'), 'single')
        border_elm.set(qn('w:sz'), '4')
        border_elm.set(qn('w:space'), '0')
        border_elm.set(qn('w:color'), 'auto')
        tcPr.append(border_elm)


# Helper function to create a qualified name (QName)
def qname(tag) -> str:
    return f'{{{nsdecls("w").strip()}}}{tag}'


def set_document_font(doc, font_name='Open Sans', font_size=Pt(10)) -> None:
    try:
        for paragraph in doc.paragraphs:
            for run in paragraph.runs:
                run.font.name = font_name
                run.font.size = font_size
        logger.info(f"Document font set to: {font_name} and size: {font_size}")
    except Exception as e:
        logger.error(f"Error setting document font: {e}", exc_info=True, stacklevel=2, stack_info=True)


def autofit_tables_to_window(doc) -> None:
    for table in doc.tables:
        table.autofit = False  # Disable autofit
        # Set table width to 100% of the page width
        tbl_width = parse_xml(r'<w:tblW {} w:w="5000" w:type="pct"/>'.format(nsdecls('w')))
        table._element.tblPr.append(tbl_width)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER


def apply_custom_styles(doc) -> None:
    try:
        for table in doc.tables:
            header_cells = table.rows[0].cells
            header_texts = [cell.text.strip() for cell in header_cells if cell.text.strip() != '']

            # Azure section tables
            if is_azure_table(header_texts):
                style_azure_table(table)

            # WPEngine section table
            elif is_wpengine_table(header_texts):
                style_wpengine_table(table)

            # Cisco section tables
            elif is_cisco_table(header_texts):
                style_cisco_table(table)
    except Exception as e:
        logger.error(f"Error applying custom styles: {e}")


def post_process_docx(doc, output_file_path) -> None:
    try:
        set_document_font(doc)
        apply_custom_styles(doc)
        autofit_tables_to_window(doc)
        doc.save(output_file_path)
        logger.info("Post-processing completed and styles applied based on table headers.")
    except Exception as e:
        logger.error(f"Error applying styles to tables: {e}", exc_info=True, stacklevel=2, stack_info=True)


def is_azure_table(header_texts) -> bool:
    azure_headers = [["Failing Controls - UGC", "Failing Controls - ZenPay"], ["Control States:", "UGC", "ZenPay"],
                     ["Resource States:", "UGC", "ZenPay"]]
    azure_other_header = len(header_texts) == 6 and header_texts[3] == ''
    return header_texts in azure_headers or azure_other_header


def is_wpengine_table(header_texts) -> bool:
    wpengine_texts = ["Plugins updated", "Domains secured", "Platform enhancements", "Attacks blocked"]
    return header_texts == wpengine_texts


def is_cisco_table(header_texts) -> bool:
    cisco_headers = [[
        "Total Data Transferred", "Total Data - DOWNLOADED", "Total Data - UPLOADED", "Total Unique Clients", "Average of clients per day",
        "Average usage per client"
    ], ["Top clients by usage", "Usage", "Usage", "Top Blocked Sites by URL", "Category", "Sites"]]
    return header_texts in cisco_headers


def style_azure_table(table) -> None:
    azure_header_fill = '5B9BD5'
    azure_content_fill = 'DEEBF7'
    azure_header_font_color = RGBColor(255, 255, 255)
    azure_content_font_color = RGBColor(0, 0, 0)
    style_table(table, azure_header_fill, azure_header_font_color, azure_content_fill, azure_content_font_color)


def style_wpengine_table(table) -> None:
    wpengine_header_fill = 'A9D18E'
    wpengine_content_fill = 'E2EFD9'
    wpengine_header_font_color = RGBColor(255, 255, 255)
    wpengine_content_font_color = RGBColor(0, 0, 0)
    style_table(table, wpengine_header_fill, wpengine_header_font_color, wpengine_content_fill, wpengine_content_font_color)


def style_cisco_table(table) -> None:
    cisco_header_fill = 'FFC000'
    cisco_content_fill = 'FFF2CC'
    cisco_header_font_color = RGBColor(255, 255, 255)
    cisco_content_font_color = RGBColor(0, 0, 0)
    style_table(table, cisco_header_fill, cisco_header_font_color, cisco_content_fill, cisco_content_font_color)


def modify_styles(doc) -> None:
    logger.info(f"Document loaded from: {doc}")
    try:
        if 'Heading 1' not in doc.styles:
            h1 = doc.styles.add_style('Heading 1', WD_STYLE_TYPE.PARAGRAPH)
        else:
            h1 = doc.styles['Heading 1']

        h1.font.name = 'Arial'
        h1.font.size = Pt(16)
        h1.font.bold = True
        h1.font.color.rgb = RGBColor(0, 0, 255)  # Blue color
        h1.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
        if 'Heading 2' not in doc.styles:
            h2 = doc.styles.add_style('Heading 2', WD_STYLE_TYPE.PARAGRAPH)
        else:
            h2 = doc.styles['Heading 2']

        h2.font.name = 'Times New Roman'
        h2.font.size = Pt(14)
        h2.font.bold = True
        h2.font.color.rgb = RGBColor(0, 5, 255)  # Blue color
        h2.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT

        if 'Block Text' not in doc.styles:
            bt = doc.styles.add_style('Block Text', WD_STYLE_TYPE.PARAGRAPH)
        else:
            bt = doc.styles['Block Text']

        bt.font.name = 'Times New Roman'
        bt.font.size = Pt(14)
        bt.font.bold = True
        bt.font.color.rgb = RGBColor(0, 5, 255)  # Blue color
        bt.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
    except Exception as e:
        logger.error(f"Error modifying 'Heading 1' style: {e}")


def set_margins(doc, top_cm, bottom_cm, left_cm, right_cm) -> None:
    for section in doc.sections:
        section.top_margin = top_cm
        section.bottom_margin = bottom_cm
        section.left_margin = left_cm
        section.right_margin = right_cm
        logger.info(f"Margins set to top: {top_cm}, bottom: {bottom_cm}, left: {left_cm}, right: {right_cm}")


# def autofit_images_to_window(doc):
#     default_width = Mm(210)  # A4 width in mm
#     default_margin_left = Cm(0.5)  # Default left margin
#     default_margin_right = Cm(0.5)  # Default right margin

#     for section in doc.sections:
#         page_width = section.page_width or default_width
#         left_margin = section.left_margin or default_margin_left
#         right_margin = section.right_margin or default_margin_right

#         usable_width = page_width - left_margin - right_margin

#         for shape in doc.inline_shapes:
#             if shape.type in (WD_INLINE_SHAPE.PICTURE, WD_INLINE_SHAPE.LINKED_PICTURE):
#                 aspect_ratio = shape.height / shape.width
#                 new_width = usable_width
#                 new_height = int(new_width * aspect_ratio)
#                 shape.width = new_width
#                 shape.height = new_height
#                 logger.info(f"Image resized to width: {new_width} pt, height: {new_height}")


def autofit_images_to_window(doc):
    # Define A4 size and margins
    default_width = Mm(210)  # A4 width in mm
    default_margin_left = Cm(1)
    default_margin_right = Cm(1)

    for section in doc.sections:
        page_width = section.page_width or default_width
        left_margin = section.left_margin or default_margin_left
        right_margin = section.right_margin or default_margin_right

        usable_width_emus = page_width - left_margin - right_margin

        for shape in doc.inline_shapes:
            if shape.type in (WD_INLINE_SHAPE.PICTURE, WD_INLINE_SHAPE.LINKED_PICTURE):
                aspect_ratio = float(shape.height) / float(shape.width)
                new_width = usable_width_emus
                new_height = round(new_width * aspect_ratio)
                shape.width = new_width
                shape.height = new_height
                logger.info(f"Image resized to width: {new_width} EMUs, height: {new_height} EMUs")


def print_section_margins(doc) -> None:
    for section in doc.sections:
        logger.info(f"Top margin: {section.top_margin}")
        logger.info(f"Bottom margin: {section.bottom_margin}")
        logger.info(f"Left margin: {section.left_margin}")
        logger.info(f"Right margin: {section.right_margin}")


def main() -> None:
    files_converted = False
    input_dir = Path(INPUT_DIR)
    for file_path in input_dir.iterdir():
        if file_path.suffix == ".md":
            logger.info(f'Processing Markdown file: {file_path}')
            docx_file_path = convert_md_to_docx(file_path)
            if docx_file_path:
                doc = Document(docx_file_path)
                post_process_docx(doc, docx_file_path)
                modify_styles(doc)
                set_margins(doc, TOP_MARGIN, BOTTOM_MARGIN, LEFT_MARGIN, RIGHT_MARGIN)
                autofit_images_to_window(doc)
                print_section_margins(doc)
                doc.save(docx_file_path)
                files_converted = True
    if not files_converted:
        logger.info("No files found.")


if __name__ == "__main__":
    main()
