import subprocess
from pathlib import Path
import logging
from typing import Any
import re
from rich.logging import RichHandler
from rich.traceback import install
from docx import Document
from docx.oxml import OxmlElement, parse_xml
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.shape import WD_INLINE_SHAPE
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Pt, RGBColor, Cm, Mm
from docx.oxml.ns import nsdecls, qn

install()


def configure_logging() -> None:
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


def initialize_directories():
    INPUT_DIR = "input/CleanedTemplate"  # pylint: disable=redefined-outer-name
    OUTPUT_DIR = "final_output"  # pylint: disable=redefined-outer-name
    REFERENCE_DIR = f'{INPUT_DIR}/Reference'  # pylint: disable=redefined-outer-name
    REFERENCE_DOC = f'{REFERENCE_DIR}/refdoc.docx'  # pylint: disable=redefined-outer-name
    CUSTOM_TEXT = ["1.Azure", "2.AWS", "3.WPEngine", "4.FraudWatch", "5.Cisco", "6.Barracuda", "7.Websites", "8.Summary"]  # pylint: disable=redefined-outer-name
    Path(INPUT_DIR).mkdir(exist_ok=True)
    Path(REFERENCE_DIR).mkdir(exist_ok=True)
    Path(OUTPUT_DIR).mkdir(exist_ok=True)
    return INPUT_DIR, OUTPUT_DIR, REFERENCE_DIR, REFERENCE_DOC, CUSTOM_TEXT


configure_logging()
INPUT_DIR, OUTPUT_DIR, REFERENCE_DIR, REFERENCE_DOC, CUSTOM_TEXT = initialize_directories()
TOP_MARGIN = Cm(1)
BOTTOM_MARGIN = Cm(1)
LEFT_MARGIN = Cm(1)
RIGHT_MARGIN = Cm(1)

logger = logging.getLogger(__name__)


class DocxProcessor:

    def __init__(self, input_dir: str, output_dir: str, reference_dir: str, reference_doc: str) -> None:
        self.input_dir = Path(input_dir)
        self.output_dir = Path(output_dir)
        self.reference_dir = Path(reference_dir)
        self.reference_doc = reference_doc

    def convert_md_to_docx(self, file_path: Path) -> Path:
        """Converts a Markdown file to a DOCX file using Pandoc.
        Args:file_path (Path): The path to the Markdown file.
        Returns:Path: The path to the converted DOCX file."""
        output_file = self.output_dir / f"{file_path.stem}.docx"
        pandoc_command = ["pandoc", str(file_path), "-o", str(output_file)]
        try:
            subprocess.run(pandoc_command, check=True)
            logger.info(f"Successfully converted {file_path} to {output_file}")
            return output_file
        except subprocess.CalledProcessError as e:
            logger.error(f"Error converting {file_path} to docx: {e}")
            return file_path

    def keep_table_together(self, table: Any) -> None:
        """Modifies a table in a DOCX document to keep it together on a single page.
        Args:table (Any): The table to be modified."""
        if tblPr := table._element.xpath('w:tblPr'):  # pylint: disable=protected-access
            tblKeep = OxmlElement('w:tblpPr')
            tblKeep.set(qn('w:keepLines'), "1")
            tblPr[0].append(tblKeep)

    def keep_sections_together(self, doc) -> None:
        # Convert document to a single string for regex processing
        doc_text = '\n'.join([p.text for p in doc.paragraphs])
        # Find the TOC section using regular expressions
        start_marker = r"\s*# Table of Contents"
        end_marker = r"\n\n---"
        start_match = re.search(start_marker, doc_text)
        end_match = re.search(end_marker, doc_text[start_match.end():], re.MULTILINE) if start_match else None
        # Identify the range of the TOC section
        if start_match and end_match:
            toc_section_start = start_match.start()
            toc_section_end = start_match.end() + end_match.end()
            for paragraph in doc.paragraphs:
                if toc_section_start <= paragraph._element.getparent().getparent().index(paragraph._element) <= toc_section_end:  # pylint: disable=protected-access
                    continue  # Skip paragraphs in the TOC section
                if paragraph.text.startswith("2. AWS") or paragraph.text.startswith("3. WPEngine"):
                    run = paragraph.add_run()
                    run.add_break(WD_BREAK.PAGE)

    def add_page_break_before_section(self, doc, section_titles, ignore_toc=False):
        if not ignore_toc:
            return
        # Convert document to a single string for regex processing
        doc_text = '\n'.join([p.text for p in doc.paragraphs])
        # Find the TOC section using regular expressions
        start_marker = r"\s*# Table of Contents"
        end_marker = r"\n\n---"
        start_match = re.search(start_marker, doc_text)
        end_match = re.search(end_marker, doc_text[start_match.end():], re.MULTILINE) if start_match else None
        # Identify the range of the TOC section
        if start_match and end_match:
            toc_section_start = start_match.start()
            toc_section_end = start_match.end() + end_match.end()
            # Process each paragraph
            for paragraph in doc.paragraphs:
                if toc_section_start <= paragraph._element.getparent().getparent().index(paragraph._element) <= toc_section_end:  # pylint: disable=protected-access
                    continue  # Skip paragraphs in the TOC section
                for title in section_titles:
                    if title in paragraph.text:
                        # Add a page break before this paragraph
                        run = paragraph.insert_paragraph_before().add_run()
                        run.add_break(WD_BREAK.PAGE)
                        break

    def set_cell_background_color(self, cell, color_str) -> None:
        shading_elm = OxmlElement('w:shd')
        shading_elm.set(qn('w:fill'), color_str)
        cell._tc.get_or_add_tcPr().append(shading_elm)  # pylint: disable=protected-access

    def style_table(self, table, header_fill, header_font_color, content_fill, content_font_color) -> None:
        for row in table.rows:
            for cell in row.cells:
                self.set_cell_borders(cell)
                paragraphs = cell.paragraphs
                for paragraph in paragraphs:
                    for run in paragraph.runs:
                        run.font.size = Pt(10)
                        run.font.name = 'Open Sans'
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    paragraph.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                self.set_cell_background_color(cell, header_fill if row == table.rows[0] else content_fill)
                self.set_font_color(cell, header_font_color if row == table.rows[0] else content_font_color)

    def set_font_color(self, cell, font_color) -> None:
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.font.color.rgb = font_color

    def set_cell_borders(self, cell) -> None:
        tc = cell._element  # pylint: disable=protected-access
        tcPr = tc.get_or_add_tcPr()
        for border in ["top", "left", "bottom", "right"]:
            border_elm = OxmlElement(f'w:{border}')
            border_elm.set(qn('w:val'), 'single')
            border_elm.set(qn('w:sz'), '4')
            border_elm.set(qn('w:space'), '0')
            border_elm.set(qn('w:color'), 'auto')
            tcPr.append(border_elm)

    # Helper function to create a qualified name (QName)
    def qname(self, tag) -> str:
        return f'{{{nsdecls("w").strip()}}}{tag}'

    def set_document_font(self, doc, font_name='Open Sans', font_size=Pt(10)) -> None:
        try:
            for paragraph in doc.paragraphs:
                for run in paragraph.runs:
                    run.font.name = font_name
                    run.font.size = font_size
            logger.info(f"Document font set to: {font_name} and size: {font_size}")
        except Exception as e:
            logger.error(f"Error setting document font: {e}", exc_info=True, stacklevel=2, stack_info=True)

    def autofit_tables_to_window(self, doc) -> None:
        for table in doc.tables:
            table.autofit = False  # Disable autofit
            # Set table width to 100% of the page width
            # tbl_width = parse_xml(r'<w:tblW {} w:w="5000" w:type="pct"/>'.format(nsdecls('w')))
            tbl_width = parse_xml(f"""<w:tblW {nsdecls('w')} w:w="5000" w:type="pct"/>""")
            table._element.tblPr.append(tbl_width)  # pylint: disable=protected-access
            table.alignment = WD_TABLE_ALIGNMENT.CENTER

    def apply_custom_styles(self, doc) -> None:
        try:
            for table in doc.tables:
                header_cells = table.rows[0].cells
                header_texts = [cell.text.strip() for cell in header_cells if cell.text.strip() != '']

                if CS.is_azure_table(header_texts):
                    CS.style_azure_table(self, table)
                elif CS.is_wpengine_table(header_texts):
                    CS.style_wpengine_table(self, table)
                elif CS.is_cisco_table(header_texts):
                    CS.style_cisco_table(self, table)
        except Exception as e:
            logger.error(f"Error applying custom styles: {e}")

    def post_process_docx(self, doc, output_file_path) -> None:
        try:
            self.set_document_font(doc)
            self.apply_custom_styles(doc)
            self.autofit_tables_to_window(doc)
            doc.save(output_file_path)
            logger.info("Post-processing completed and styles applied based on table headers.")
        except Exception as e:
            logger.error(f"Error applying styles to tables: {e}", exc_info=True, stacklevel=2, stack_info=True)

    def _apply_style(self, font_name, style, font_size, font_color_offset) -> None:
        try:
            style.font.name = font_name
            style.font.size = Pt(font_size)
            style.font.bold = True
            style.font.color.rgb = RGBColor(0, font_color_offset, 255)
            style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
        except Exception as e:
            logger.error(f"Error applying style: {e}")

    def set_margins(self, doc, top_cm, bottom_cm, left_cm, right_cm) -> None:
        for section in doc.sections:
            section.top_margin = top_cm
            section.bottom_margin = bottom_cm
            section.left_margin = left_cm
            section.right_margin = right_cm
            # logger.info(f"Margins set to top: {top_cm}, bottom: {bottom_cm}, left: {left_cm}, right: {right_cm}")

    def autofit_images_to_window(self, doc) -> None:
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
                    # logger.info(f"Image resized to width: {new_width} EMUs, height: {new_height} EMUs")

    def print_section_margins(self, doc) -> None:
        for section in doc.sections:
            logger.info(f"Top margin: {section.top_margin}")
            logger.info(f"Bottom margin: {section.bottom_margin}")
            logger.info(f"Left margin: {section.left_margin}")
            logger.info(f"Right margin: {section.right_margin}")

    def modify_document_styles(self, doc) -> None:
        try:
            heading1_style = (doc.styles.add_style('Heading 1', WD_STYLE_TYPE.PARAGRAPH)
                              if 'Heading 1' not in doc.styles else doc.styles['Heading 1'])
            self._apply_style('Arial', heading1_style, 16, 0)

            if 'Heading 2' not in doc.styles:
                heading2_style = doc.styles.add_style('Heading 2', WD_STYLE_TYPE.PARAGRAPH)
            else:
                heading2_style = doc.styles['Heading 2']

            self._apply_style('Times New Roman', heading2_style, 14, 5)

            if 'Block Text' not in doc.styles:
                block_text_style = doc.styles.add_style('Block Text', WD_STYLE_TYPE.PARAGRAPH)
            else:
                block_text_style = doc.styles['Block Text']

            self._apply_style('Times New Roman', block_text_style, 14, 5)
        except Exception as e:
            logger.error(f"Error modifying 'Heading 1' style: {e}")


class CS():

    @staticmethod
    def style_azure_table(doc_processor, table) -> None:
        azure_header_fill = '5B9BD5'
        azure_content_fill = 'DEEBF7'
        azure_header_font_color = RGBColor(255, 255, 255)
        azure_content_font_color = RGBColor(0, 0, 0)
        doc_processor.style_table(table, azure_header_fill, azure_header_font_color, azure_content_fill, azure_content_font_color)

    @staticmethod
    def is_wpengine_table(header_texts) -> bool:
        wpengine_texts = ["Plugins updated", "Domains secured", "Platform enhancements", "Attacks blocked"]
        return header_texts == wpengine_texts

    @staticmethod
    def is_cisco_table(header_texts) -> bool:
        cisco_headers = [[
            "Total Data Transferred", "Total Data - DOWNLOADED", "Total Data - UPLOADED", "Total Unique Clients", "Average of clients per day",
            "Average usage per client"
        ], ["Top clients by usage", "Usage", "Usage", "Top Blocked Sites by URL", "Category", "Sites"]]  # pylint: disable=line-too-long
        return header_texts in cisco_headers

    @staticmethod
    def is_azure_table(header_texts) -> bool:
        azure_headers = [["Failing Controls - UGC", "Failing Controls - ZenPay"], ["Control States:", "UGC", "ZenPay"],
                         ["Resource States:", "UGC", "ZenPay"]]
        azure_other_header = len(header_texts) == 6 and header_texts[3] == ''
        return header_texts in azure_headers or azure_other_header

    @staticmethod
    def style_wpengine_table(doc_processor, table) -> None:
        wpengine_header_fill = 'A9D18E'
        wpengine_content_fill = 'E2EFD9'
        wpengine_header_font_color = RGBColor(255, 255, 255)
        wpengine_content_font_color = RGBColor(0, 0, 0)
        doc_processor.style_table(table, wpengine_header_fill, wpengine_header_font_color, wpengine_content_fill, wpengine_content_font_color)  # pylint: disable=line-too-long

    @staticmethod
    def style_cisco_table(doc_processor, table) -> None:
        cisco_header_fill = 'FFC000'
        cisco_content_fill = 'FFF2CC'
        cisco_header_font_color = RGBColor(255, 255, 255)
        cisco_content_font_color = RGBColor(0, 0, 0)
        doc_processor.style_table(table, cisco_header_fill, cisco_header_font_color, cisco_content_fill, cisco_content_font_color)


def main() -> None:
    files_converted = False
    input_dir = Path(INPUT_DIR)
    ref_dir = Path(REFERENCE_DIR)
    ref_doc = Path(REFERENCE_DOC)
    output_dir = Path(OUTPUT_DIR)
    section_titles_with_breaks = list[str](CUSTOM_TEXT)
    docx = DocxProcessor(str(input_dir), str(output_dir), str(ref_dir), str(ref_doc))
    for file_path in input_dir.iterdir():
        if file_path.suffix == ".md":
            logger.info(f'Processing Markdown file: {file_path}')
            if (docx_file_path := docx.convert_md_to_docx(file_path)):
                # doc = Document(docx_file_path)
                doc = Document(str(docx_file_path))
                docx.keep_sections_together(doc)  # Call to keep_sections_together
                docx.post_process_docx(doc, docx_file_path)  # set_document_font, self.apply_custom_styles, self.autofit_tables_to_window
                docx.add_page_break_before_section(doc, section_titles_with_breaks, ignore_toc=True)
                docx.modify_document_styles(doc)  # _apply_style
                logger.info('Modifying document styles')
                docx.set_margins(doc, TOP_MARGIN, BOTTOM_MARGIN, LEFT_MARGIN, RIGHT_MARGIN)
                logger.info('Setting margins')
                docx.autofit_images_to_window(doc)
                logger.info('Autofitting images to window')
                docx.print_section_margins(doc)
                for table in doc.tables:
                    docx.keep_table_together(table)
                logger.info('Printing section margins')
                doc.save(docx_file_path)
                logger.info('Saving document')
                files_converted = True
    if not files_converted:
        logger.info("No files found.")


if __name__ == "__main__":
    main()
