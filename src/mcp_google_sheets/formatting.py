#!/usr/bin/env python
"""Cell formatting for the Google Sheets MCP server.

``batch_update`` already reaches every formatting request the API offers, but
only by hand-writing raw API JSON: the caller has to look up the numeric sheet
id, convert A1 notation to half-open 0-based indices, express colours as 0..1
floats, and get the field mask right. These tools do that plumbing so a colour
change is one call with a hex string.
"""

import logging
import re
from typing import Any, Dict, List, Optional

from mcp.server.fastmcp import Context
from mcp.types import ToolAnnotations

from .server import _column_index_to_letter, _get_sheet_id, _parse_a1_notation, tool

logger = logging.getLogger(__name__)

# Sentinels that reset a colour to the spreadsheet default rather than paint it.
CLEAR_VALUES = frozenset({"clear", "none", "default", "reset"})

NAMED_COLORS = {
    "black": "#000000",
    "white": "#ffffff",
    "red": "#ff0000",
    "green": "#00ff00",
    "blue": "#0000ff",
    "yellow": "#ffff00",
    "orange": "#ff9900",
    "purple": "#9900ff",
    "magenta": "#ff00ff",
    "cyan": "#00ffff",
    "gray": "#808080",
    "grey": "#808080",
    "lightgray": "#d9d9d9",
    "lightgrey": "#d9d9d9",
    "darkgray": "#404040",
    "darkgrey": "#404040",
    # Muted tones that read well as cell fills behind black text.
    "lightred": "#f4cccc",
    "lightgreen": "#d9ead3",
    "lightblue": "#cfe2f3",
    "lightyellow": "#fff2cc",
    "lightorange": "#fce5cd",
    "lightpurple": "#d9d2e9",
}

HORIZONTAL_ALIGNMENTS = frozenset({"LEFT", "CENTER", "RIGHT"})
VERTICAL_ALIGNMENTS = frozenset({"TOP", "MIDDLE", "BOTTOM"})
WRAP_STRATEGIES = frozenset({"OVERFLOW_CELL", "LEGACY_WRAP", "CLIP", "WRAP"})
NUMBER_FORMAT_TYPES = frozenset(
    {"TEXT", "NUMBER", "PERCENT", "CURRENCY", "DATE", "TIME", "DATE_TIME", "SCIENTIFIC"}
)

HEX_RE = re.compile(r"^#?([0-9a-fA-F]{3}|[0-9a-fA-F]{6})$")


def _parse_color(value: Any) -> Optional[Dict[str, float]]:
    """Turn a colour in any accepted form into the API's 0..1 rgb object.

    Accepts ``'#RRGGBB'``, ``'#RGB'``, a name from NAMED_COLORS, or a dict of
    0..1 floats. Returns None for the clear sentinels, meaning "drop this back
    to the default", which the caller expresses by naming the field in the
    update mask while omitting its value.
    """
    if isinstance(value, dict):
        parsed = {}
        for channel in ("red", "green", "blue"):
            channel_value = float(value.get(channel, 0.0))
            if not 0.0 <= channel_value <= 1.0:
                raise ValueError(f"Colour channel '{channel}' must be between 0 and 1")
            parsed[channel] = channel_value
        return parsed

    if not isinstance(value, str):
        raise ValueError(f"Colour must be a hex string, a colour name, or a dict, got: {value!r}")

    text = value.strip().lower()
    if text in CLEAR_VALUES:
        return None

    hex_text = NAMED_COLORS.get(text, text)
    match = HEX_RE.match(hex_text)
    if not match:
        raise ValueError(
            f"Unrecognised colour {value!r}. Use '#RRGGBB', '#RGB', a dict of 0..1 floats, "
            f"one of {sorted(NAMED_COLORS)}, or 'clear' to reset."
        )

    digits = match.group(1)
    if len(digits) == 3:
        digits = "".join(char * 2 for char in digits)
    return {
        "red": int(digits[0:2], 16) / 255.0,
        "green": int(digits[2:4], 16) / 255.0,
        "blue": int(digits[4:6], 16) / 255.0,
    }


def _check_enum(name: str, value: str, allowed: frozenset) -> str:
    upper = value.strip().upper()
    if upper not in allowed:
        raise ValueError(f"{name} must be one of {sorted(allowed)}, got: {value!r}")
    return upper


def _build_cell_format(
    background_color: Optional[Any] = None,
    text_color: Optional[Any] = None,
    bold: Optional[bool] = None,
    italic: Optional[bool] = None,
    underline: Optional[bool] = None,
    strikethrough: Optional[bool] = None,
    font_size: Optional[int] = None,
    font_family: Optional[str] = None,
    horizontal_alignment: Optional[str] = None,
    vertical_alignment: Optional[str] = None,
    wrap_strategy: Optional[str] = None,
    number_format: Optional[str] = None,
) -> tuple[Dict[str, Any], List[str]]:
    """Build a CellFormat plus the field mask naming exactly what it sets.

    A field named in the mask but absent from the format is how the API is told
    to reset that property, so clearing a colour and setting one share a path.
    """
    cell_format: Dict[str, Any] = {}
    fields: List[str] = []

    if background_color is not None:
        parsed = _parse_color(background_color)
        if parsed is not None:
            cell_format["backgroundColorStyle"] = {"rgbColor": parsed}
        fields.append("userEnteredFormat.backgroundColorStyle")

    text_format: Dict[str, Any] = {}
    if text_color is not None:
        parsed = _parse_color(text_color)
        if parsed is not None:
            text_format["foregroundColorStyle"] = {"rgbColor": parsed}
        fields.append("userEnteredFormat.textFormat.foregroundColorStyle")

    for name, value in (
        ("bold", bold),
        ("italic", italic),
        ("underline", underline),
        ("strikethrough", strikethrough),
    ):
        if value is not None:
            text_format[name] = bool(value)
            fields.append(f"userEnteredFormat.textFormat.{name}")

    if font_size is not None:
        if int(font_size) <= 0:
            raise ValueError("font_size must be a positive number of points")
        text_format["fontSize"] = int(font_size)
        fields.append("userEnteredFormat.textFormat.fontSize")

    if font_family is not None:
        text_format["fontFamily"] = font_family
        fields.append("userEnteredFormat.textFormat.fontFamily")

    if text_format:
        cell_format["textFormat"] = text_format

    if horizontal_alignment is not None:
        cell_format["horizontalAlignment"] = _check_enum(
            "horizontal_alignment", horizontal_alignment, HORIZONTAL_ALIGNMENTS
        )
        fields.append("userEnteredFormat.horizontalAlignment")

    if vertical_alignment is not None:
        cell_format["verticalAlignment"] = _check_enum(
            "vertical_alignment", vertical_alignment, VERTICAL_ALIGNMENTS
        )
        fields.append("userEnteredFormat.verticalAlignment")

    if wrap_strategy is not None:
        cell_format["wrapStrategy"] = _check_enum("wrap_strategy", wrap_strategy, WRAP_STRATEGIES)
        fields.append("userEnteredFormat.wrapStrategy")

    if number_format is not None:
        pattern: Optional[str] = None
        if ":" in number_format:
            type_part, _, pattern = number_format.partition(":")
        else:
            type_part = number_format
        format_type = _check_enum("number_format type", type_part, NUMBER_FORMAT_TYPES)
        cell_format["numberFormat"] = {"type": format_type}
        if pattern:
            cell_format["numberFormat"]["pattern"] = pattern
        fields.append("userEnteredFormat.numberFormat")

    return cell_format, fields


def _repeat_cell_request(sheet_id: int, cell_range: str, **format_kwargs) -> Dict[str, Any]:
    """Build one repeatCell request for an A1 range on a sheet."""
    bare_range = cell_range.split("!")[-1].replace("$", "").strip()
    if not bare_range:
        raise ValueError("range cannot be empty")

    grid_range: Dict[str, Any] = {"sheetId": sheet_id}
    grid_range.update(_parse_a1_notation(bare_range))

    cell_format, fields = _build_cell_format(**format_kwargs)
    if not fields:
        raise ValueError(
            f"No formatting given for range {cell_range!r}. Set at least one of "
            "background_color, text_color, bold, italic, underline, strikethrough, "
            "font_size, font_family, horizontal_alignment, vertical_alignment, "
            "wrap_strategy, or number_format."
        )

    return {
        "repeatCell": {
            "range": grid_range,
            "cell": {"userEnteredFormat": cell_format},
            "fields": ",".join(fields),
        }
    }


@tool(
    annotations=ToolAnnotations(
        title="Format Cells",
        destructiveHint=True,
    ),
)
def format_cells(spreadsheet_id: str,
                 sheet: str,
                 range: str,
                 background_color: Optional[str] = None,
                 text_color: Optional[str] = None,
                 bold: Optional[bool] = None,
                 italic: Optional[bool] = None,
                 underline: Optional[bool] = None,
                 strikethrough: Optional[bool] = None,
                 font_size: Optional[int] = None,
                 font_family: Optional[str] = None,
                 horizontal_alignment: Optional[str] = None,
                 vertical_alignment: Optional[str] = None,
                 wrap_strategy: Optional[str] = None,
                 number_format: Optional[str] = None,
                 ctx: Context = None) -> Dict[str, Any]:
    """
    Set the background colour and other formatting on a range of cells.

    Only the properties you pass are changed; everything else in the range keeps
    its current formatting. Pass 'clear' as a colour to reset it to the default.

    Args:
        spreadsheet_id: The ID of the spreadsheet (found in the URL)
        sheet: The name of the sheet
        range: Cell range in A1 notation (e.g. 'B2', 'A1:C10', 'A:C', '1:5')
        background_color: Cell fill. '#RRGGBB', '#RGB', a colour name such as
            'lightgreen', or 'clear' to reset it.
        text_color: Font colour, same accepted forms as background_color
        bold: Bold the text
        italic: Italicise the text
        underline: Underline the text
        strikethrough: Strike the text through
        font_size: Font size in points
        font_family: Font family name (e.g. 'Roboto')
        horizontal_alignment: LEFT, CENTER, or RIGHT
        vertical_alignment: TOP, MIDDLE, or BOTTOM
        wrap_strategy: OVERFLOW_CELL, LEGACY_WRAP, CLIP, or WRAP
        number_format: A type from TEXT, NUMBER, PERCENT, CURRENCY, DATE, TIME,
            DATE_TIME, SCIENTIFIC, optionally with a pattern after a colon
            (e.g. 'CURRENCY:"$"#,##0.00')

    Returns:
        Result of the formatting operation
    """
    sheets_service = ctx.request_context.lifespan_context.sheets_service

    sheet_id = _get_sheet_id(sheets_service, spreadsheet_id, sheet)
    if sheet_id is None:
        return {"error": f"Sheet {sheet!r} not found in spreadsheet {spreadsheet_id}"}

    try:
        request = _repeat_cell_request(
            sheet_id,
            range,
            background_color=background_color,
            text_color=text_color,
            bold=bold,
            italic=italic,
            underline=underline,
            strikethrough=strikethrough,
            font_size=font_size,
            font_family=font_family,
            horizontal_alignment=horizontal_alignment,
            vertical_alignment=vertical_alignment,
            wrap_strategy=wrap_strategy,
            number_format=number_format,
        )
    except ValueError as exc:
        return {"error": str(exc)}

    result = (
        sheets_service.spreadsheets()
        .batchUpdate(spreadsheetId=spreadsheet_id, body={"requests": [request]})
        .execute()
    )

    return {
        "success": True,
        "range": f"{sheet}!{range}",
        "fields": request["repeatCell"]["fields"],
        "result": result,
    }


@tool(
    annotations=ToolAnnotations(
        title="Batch Format Cells",
        destructiveHint=True,
    ),
)
def batch_format_cells(spreadsheet_id: str,
                       sheet: str,
                       formats: Dict[str, Dict[str, Any]],
                       ctx: Context = None) -> Dict[str, Any]:
    """
    Apply different formatting to several ranges in a single API call.

    Use this instead of repeated format_cells calls when colouring a status
    column, banding a table, or otherwise painting many ranges at once.

    Args:
        spreadsheet_id: The ID of the spreadsheet (found in the URL)
        sheet: The name of the sheet
        formats: Map of A1 range to the formatting for that range, using the
            same option names as format_cells. For example:
            {
                "A1:D1": {"background_color": "#4a86e8", "text_color": "white", "bold": true},
                "C2:C50": {"background_color": "lightgreen"},
                "D2:D50": {"number_format": "PERCENT:0.0%"}
            }

    Returns:
        Result of the batch formatting operation
    """
    if not formats:
        return {"error": "formats cannot be empty"}

    sheets_service = ctx.request_context.lifespan_context.sheets_service

    sheet_id = _get_sheet_id(sheets_service, spreadsheet_id, sheet)
    if sheet_id is None:
        return {"error": f"Sheet {sheet!r} not found in spreadsheet {spreadsheet_id}"}

    requests: List[Dict[str, Any]] = []
    for cell_range, options in formats.items():
        if not isinstance(options, dict):
            return {"error": f"Formatting for range {cell_range!r} must be an object"}
        try:
            requests.append(_repeat_cell_request(sheet_id, cell_range, **options))
        except TypeError as exc:
            return {"error": f"Unknown formatting option for range {cell_range!r}: {exc}"}
        except ValueError as exc:
            return {"error": str(exc)}

    result = (
        sheets_service.spreadsheets()
        .batchUpdate(spreadsheetId=spreadsheet_id, body={"requests": requests})
        .execute()
    )

    return {
        "success": True,
        "sheet": sheet,
        "rangesFormatted": list(formats),
        "result": result,
    }


def _rgb_to_hex(color: Dict[str, Any]) -> str:
    """Render an API rgb colour object as '#RRGGBB'."""
    return "#" + "".join(
        f"{round(float(color.get(channel, 0.0)) * 255):02x}"
        for channel in ("red", "green", "blue")
    )


def _summarise_format(cell_format: Dict[str, Any]) -> Dict[str, Any]:
    """Reduce a CellFormat to the properties worth reporting.

    The default white fill and black text are skipped so a sheet whose cells
    were only ever painted back to default reports nothing.
    """
    summary: Dict[str, Any] = {}

    background = cell_format.get("backgroundColorStyle", {}).get("rgbColor")
    if background:
        hex_color = _rgb_to_hex(background)
        if hex_color != "#ffffff":
            summary["backgroundColor"] = hex_color

    text_format = cell_format.get("textFormat", {})
    foreground = text_format.get("foregroundColorStyle", {}).get("rgbColor")
    if foreground:
        hex_color = _rgb_to_hex(foreground)
        if hex_color != "#000000":
            summary["textColor"] = hex_color

    for name in ("bold", "italic", "underline", "strikethrough"):
        if text_format.get(name):
            summary[name] = True

    for name in ("horizontalAlignment", "verticalAlignment"):
        if cell_format.get(name):
            summary[name] = cell_format[name]

    return summary


@tool(
    annotations=ToolAnnotations(
        title="Get Cell Formats",
        readOnlyHint=True,
    ),
)
def get_cell_formats(spreadsheet_id: str,
                     sheet: str,
                     range: Optional[str] = None,
                     ctx: Context = None) -> Dict[str, Any]:
    """
    Read back the formatting actually in effect on a range of cells.

    Reports only cells that carry formatting somebody actually applied, as a
    compact per-cell summary rather than the very large raw grid data the API
    returns. Reads userEnteredFormat rather than effectiveFormat, because the
    latter resolves defaults for every populated cell and would report LEFT and
    BOTTOM alignment on an entirely unstyled sheet.

    Args:
        spreadsheet_id: The ID of the spreadsheet (found in the URL)
        sheet: The name of the sheet
        range: Optional cell range in A1 notation. Defaults to the whole sheet.

    Returns:
        Per-cell background colour, text colour, and text styling
    """
    sheets_service = ctx.request_context.lifespan_context.sheets_service
    full_range = f"{sheet}!{range}" if range else sheet

    response = (
        sheets_service.spreadsheets()
        .get(
            spreadsheetId=spreadsheet_id,
            ranges=[full_range],
            fields=(
                "sheets(data(startRow,startColumn,rowData(values(userEnteredFormat("
                "backgroundColorStyle,textFormat,horizontalAlignment,verticalAlignment)))))"
            ),
        )
        .execute()
    )

    cells: List[Dict[str, Any]] = []
    for sheet_data in response.get("sheets", []):
        for block in sheet_data.get("data", []):
            start_row = block.get("startRow", 0)
            start_column = block.get("startColumn", 0)
            for row_offset, row in enumerate(block.get("rowData", [])):
                for col_offset, cell in enumerate(row.get("values", [])):
                    summary = _summarise_format(cell.get("userEnteredFormat", {}))
                    if not summary:
                        continue
                    summary["cell"] = (
                        f"{_column_index_to_letter(start_column + col_offset)}"
                        f"{start_row + row_offset + 1}"
                    )
                    cells.append(summary)

    return {"sheet": sheet, "count": len(cells), "cells": cells}
