"""
Template Engine module for Universal AI Excel Format Converter.
Reads user-uploaded Excel templates, extracts style properties (fonts, borders, fills,
alignments, row heights, column widths, merged cells), and clones them onto output workbooks.
"""

from typing import Union, Optional, Tuple, Dict, Any
import io
import openpyxl
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from logger import logger


class TemplateEngine:
    """Manages reading, caching, and cloning Excel template formats."""

    def __init__(self, template_input: Optional[Union[str, io.BytesIO, bytes]] = None):
        self.template_wb: Optional[openpyxl.Workbook] = None
        self.template_ws: Optional[Worksheet] = None
        if template_input:
            self.load_template(template_input)

    def load_template(self, template_input: Union[str, io.BytesIO, bytes]) -> None:
        """Loads template workbook from file path, stream, or bytes."""
        try:
            if isinstance(template_input, bytes):
                template_input = io.BytesIO(template_input)
            self.template_wb = openpyxl.load_workbook(template_input)
            self.template_ws = self.template_wb.active
            logger.info(f"Template loaded successfully. Active sheet: '{self.template_ws.title}'")
        except Exception as e:
            logger.error("Failed to load template workbook", exc_info=True)
            self.template_wb = None
            self.template_ws = None

    def is_loaded(self) -> bool:
        """Returns True if a valid template is loaded."""
        return self.template_ws is not None

    def apply_template_to_worksheet(self, target_ws: Worksheet) -> None:
        """Clones fonts, fills, borders, alignments, column widths, and page setup from template to target worksheet."""
        if not self.is_loaded() or self.template_ws is None:
            return

        src_ws = self.template_ws
        logger.info("Applying template styling to target worksheet...")

        # Copy Column Widths
        for col_letter, col_dim in src_ws.column_dimensions.items():
            if col_dim.width:
                target_ws.column_dimensions[col_letter].width = col_dim.width

        # Copy Header Row Style (Row 1)
        for col_idx in range(1, src_ws.max_column + 1):
            src_cell = src_ws.cell(row=1, column=col_idx)
            tgt_cell = target_ws.cell(row=1, column=col_idx)

            if src_cell.font:
                tgt_cell.font = Font(
                    name=src_cell.font.name,
                    size=src_cell.font.size,
                    bold=src_cell.font.bold,
                    italic=src_cell.font.italic,
                    color=src_cell.font.color
                )
            if src_cell.fill and src_cell.fill.fill_type:
                tgt_cell.fill = PatternFill(
                    fill_type=src_cell.fill.fill_type,
                    start_color=src_cell.fill.start_color,
                    end_color=src_cell.fill.end_color
                )
            if src_cell.border:
                tgt_cell.border = Border(
                    left=src_cell.border.left,
                    right=src_cell.border.right,
                    top=src_cell.border.top,
                    bottom=src_cell.border.bottom
                )
            if src_cell.alignment:
                tgt_cell.alignment = Alignment(
                    horizontal=src_cell.alignment.horizontal,
                    vertical=src_cell.alignment.vertical,
                    wrap_text=src_cell.alignment.wrap_text
                )

        # Copy Page Setup
        target_ws.page_setup.orientation = src_ws.page_setup.orientation
        target_ws.page_setup.paperSize = src_ws.page_setup.paperSize
        target_ws.page_setup.fitToWidth = src_ws.page_setup.fitToWidth
        target_ws.page_setup.fitToHeight = src_ws.page_setup.fitToHeight
        target_ws.sheet_properties.pageSetUpPr.fitToPage = src_ws.sheet_properties.pageSetUpPr.fitToPage

        logger.info("Template styling applied successfully.")
