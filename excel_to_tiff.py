"""Convert an Excel form to a TIFF image.

Uses openpyxl for sheet geometry and style, pandas for the text grid, and
Windows .NET System.Drawing (through pythonnet) to encode the TIFF.
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.cell.cell import Cell
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

DPI = 96
WIDTH_PX = 8
TEXT_PAD = 2
DEFAULT_COL_WIDTH = 8.43
DEFAULT_ROW_HEIGHT = 15.0


def load_drawing():
    """Load System.Drawing. Exit if Windows .NET cannot provide it."""
    try:
        from pythonnet import load
    except ImportError as exc:
        raise SystemExit(
            "Windows .NET System.Drawing is required to write a TIFF. "
            "Install pythonnet so Python can load it."
        ) from exc

    errors = []
    for runtime in ("netfx", "coreclr"):
        try:
            load(runtime)
            import clr

            clr.AddReference("System.Drawing")
            return clr
        except Exception as exc:
            errors.append(f"{runtime}: {exc}")
    detail = "; ".join(errors)
    raise SystemExit(
        "Windows .NET System.Drawing is required to write a TIFF. "
        f"Could not load System.Drawing ({detail})."
    )


def _rgb(color, fallback):
    if color is None or getattr(color, "type", None) != "rgb":
        return fallback
    raw = getattr(color, "rgb", None)
    if not raw:
        return fallback
    hex_rgb = str(raw)
    if len(hex_rgb) == 8:
        hex_rgb = hex_rgb[2:]
    if len(hex_rgb) != 6:
        return fallback
    try:
        red = int(hex_rgb[0:2], 16)
        green = int(hex_rgb[2:4], 16)
        blue = int(hex_rgb[4:6], 16)
    except ValueError:
        return fallback
    from System.Drawing import Color

    return Color.FromArgb(red, green, blue)


def _column_width(sheet: Worksheet, column: int) -> int:
    letter = get_column_letter(column)
    width = sheet.column_dimensions[letter].width
    if width is None:
        width = DEFAULT_COL_WIDTH
    return max(1, int(round(float(width) * WIDTH_PX)))


def _row_height(sheet: Worksheet, row: int) -> int:
    height = sheet.row_dimensions[row].height
    if height is None:
        height = DEFAULT_ROW_HEIGHT
    return max(1, int(round(float(height) * DPI / 72.0)))


def _sheet_is_used(sheet: Worksheet) -> bool:
    for row in sheet.iter_rows():
        for cell in row:
            if cell.value is not None:
                return True
    return False


def _string_grid(sheet: Worksheet) -> pd.DataFrame:
    rows = []
    for row in sheet.iter_rows(
        min_row=1,
        max_row=sheet.max_row,
        min_col=1,
        max_col=sheet.max_column,
    ):
        rows.append(["" if cell.value is None else cell.value for cell in row])
    frame = pd.DataFrame(rows)
    return frame.fillna("").astype(str)


def _merge_origins(sheet: Worksheet) -> dict[tuple[int, int], tuple[int, int, int, int]]:
    origins: dict[tuple[int, int], tuple[int, int, int, int]] = {}
    for merged in sheet.merged_cells.ranges:
        bounds = (merged.min_row, merged.min_col, merged.max_row, merged.max_col)
        for row in range(merged.min_row, merged.max_row + 1):
            for col in range(merged.min_col, merged.max_col + 1):
                origins[(row, col)] = bounds
    return origins


def _font(cell: Cell, font_style):
    from System.Drawing import Font

    name = cell.font.name or "Calibri"
    size = float(cell.font.size or 11)
    style = font_style.Regular
    if cell.font.bold:
        style |= font_style.Bold
    if cell.font.italic:
        style |= font_style.Italic
    if cell.font.underline:
        style |= font_style.Underline
    return Font(name, size, style)


def _string_format(cell: Cell, string_alignment, string_format_flags, string_trimming):
    from System.Drawing import StringFormat

    fmt = StringFormat()
    horizontal = (cell.alignment.horizontal or "left").lower()
    vertical = (cell.alignment.vertical or "bottom").lower()
    fmt.Alignment = {
        "left": string_alignment.Near,
        "center": string_alignment.Center,
        "right": string_alignment.Far,
        "justify": string_alignment.Near,
        "general": string_alignment.Near,
    }.get(horizontal, string_alignment.Near)
    fmt.LineAlignment = {
        "top": string_alignment.Near,
        "center": string_alignment.Center,
        "bottom": string_alignment.Far,
        "justify": string_alignment.Near,
    }.get(vertical, string_alignment.Far)
    flags = string_format_flags.LineLimit
    if not cell.alignment.wrap_text:
        flags |= string_format_flags.NoWrap
    fmt.FormatFlags = flags
    fmt.Trimming = string_trimming.EllipsisCharacter
    return fmt


def _draw_borders(graphics, cell: Cell, rect, color_cls, pen_cls) -> None:
    edges = (
        (cell.border.left, rect.Left, rect.Top, rect.Left, rect.Bottom),
        (cell.border.right, rect.Right - 1, rect.Top, rect.Right - 1, rect.Bottom),
        (cell.border.top, rect.Left, rect.Top, rect.Right, rect.Top),
        (cell.border.bottom, rect.Left, rect.Bottom - 1, rect.Right, rect.Bottom - 1),
    )
    for side, x1, y1, x2, y2 in edges:
        if side is None or not side.style:
            continue
        color = _rgb(side.color, color_cls.Black)
        pen = pen_cls(color, 1)
        try:
            graphics.DrawLine(pen, int(x1), int(y1), int(x2), int(y2))
        finally:
            pen.Dispose()


def _render_sheet(sheet: Worksheet, drawing: dict):
    grid = _string_grid(sheet)
    origins = _merge_origins(sheet)
    max_row, max_col = grid.shape
    col_widths = [_column_width(sheet, col + 1) for col in range(max_col)]
    row_heights = [_row_height(sheet, row + 1) for row in range(max_row)]
    width = max(1, sum(col_widths))
    height = max(1, sum(row_heights))

    bitmap_cls = drawing["Bitmap"]
    graphics_cls = drawing["Graphics"]
    color_cls = drawing["Color"]
    rect_cls = drawing["Rectangle"]
    rectf_cls = drawing["RectangleF"]
    brush_cls = drawing["SolidBrush"]
    pen_cls = drawing["Pen"]
    bitmap = bitmap_cls(width, height)
    graphics = graphics_cls.FromImage(bitmap)
    graphics.Clear(color_cls.White)
    graphics.TextRenderingHint = drawing["TextRenderingHint"].AntiAliasGridFit

    y_offsets = [0]
    for size in row_heights:
        y_offsets.append(y_offsets[-1] + size)
    x_offsets = [0]
    for size in col_widths:
        x_offsets.append(x_offsets[-1] + size)

    for row in range(max_row):
        for col in range(max_col):
            excel_row = row + 1
            excel_col = col + 1
            bounds = origins.get((excel_row, excel_col))
            if bounds is not None and (excel_row, excel_col) != (bounds[0], bounds[1]):
                continue
            if bounds is None:
                x = x_offsets[col]
                y = y_offsets[row]
                rect_w = col_widths[col]
                rect_h = row_heights[row]
                cell = sheet.cell(excel_row, excel_col)
            else:
                r1, c1, r2, c2 = bounds
                x = x_offsets[c1 - 1]
                y = y_offsets[r1 - 1]
                rect_w = x_offsets[c2] - x
                rect_h = y_offsets[r2] - y
                cell = sheet.cell(r1, c1)
            rect = rect_cls(x, y, rect_w, rect_h)
            fill = _rgb(
                getattr(cell.fill, "fgColor", None) if cell.fill.patternType else None,
                color_cls.White,
            )
            brush = brush_cls(fill)
            try:
                graphics.FillRectangle(brush, rect)
            finally:
                brush.Dispose()
            _draw_borders(graphics, cell, rect, color_cls, pen_cls)
            text = grid.iat[row, col]
            if bounds is not None:
                text = grid.iat[bounds[0] - 1, bounds[1] - 1]
            if not text:
                continue
            font = _font(cell, drawing["FontStyle"])
            text_brush = brush_cls(_rgb(cell.font.color, color_cls.Black))
            fmt = _string_format(
                cell,
                drawing["StringAlignment"],
                drawing["StringFormatFlags"],
                drawing["StringTrimming"],
            )
            text_rect = rectf_cls(
                float(x + TEXT_PAD),
                float(y + TEXT_PAD),
                float(max(1, rect_w - TEXT_PAD * 2)),
                float(max(1, rect_h - TEXT_PAD * 2)),
            )
            state = graphics.Save()
            try:
                graphics.SetClip(rect)
                graphics.DrawString(text, font, text_brush, text_rect, fmt)
            finally:
                graphics.Restore(state)
                font.Dispose()
                text_brush.Dispose()
                fmt.Dispose()
    graphics.Dispose()
    return bitmap


def _tiff_codec(image_codec_info):
    for info in image_codec_info.GetImageEncoders():
        if info.MimeType == "image/tiff":
            return info
    raise SystemExit(
        "Windows .NET System.Drawing is required to write a TIFF. "
        "The TIFF encoder is not available."
    )


def save_tiff(bitmaps, output: Path, drawing: dict) -> None:
    output.parent.mkdir(parents=True, exist_ok=True)
    path = str(output)
    if len(bitmaps) == 1:
        bitmaps[0].Save(path, drawing["ImageFormat"].Tiff)
        bitmaps[0].Dispose()
        return

    codec = _tiff_codec(drawing["ImageCodecInfo"])
    encoder = drawing["Encoder"]
    encoder_parameter = drawing["EncoderParameter"]
    encoder_parameters = drawing["EncoderParameters"]
    encoder_value = drawing["EncoderValue"]

    def flag(value):
        params = encoder_parameters(1)
        parameter = encoder_parameter(encoder.SaveFlag, int(value))
        current = params.Param
        current[0] = parameter
        params.Param = current
        return params

    first = bitmaps[0]
    multi = flag(encoder_value.MultiFrame)
    first.Save(path, codec, multi)
    page = flag(encoder_value.FrameDimensionPage)
    for frame in bitmaps[1:]:
        first.SaveAdd(frame, page)
        frame.Dispose()
    first.SaveAdd(flag(encoder_value.Flush))
    first.Dispose()


def convert(source: Path, output: Path) -> None:
    load_drawing()
    from System.Drawing import (
        Bitmap,
        Color,
        FontStyle,
        Graphics,
        Pen,
        Rectangle,
        RectangleF,
        SolidBrush,
        StringAlignment,
    )
    from System.Drawing import StringFormatFlags, StringTrimming
    from System.Drawing.Imaging import (
        Encoder,
        EncoderParameter,
        EncoderParameters,
        EncoderValue,
        ImageCodecInfo,
        ImageFormat,
    )
    from System.Drawing.Text import TextRenderingHint

    drawing = {
        "Bitmap": Bitmap,
        "Color": Color,
        "FontStyle": FontStyle,
        "Graphics": Graphics,
        "Pen": Pen,
        "Rectangle": Rectangle,
        "RectangleF": RectangleF,
        "SolidBrush": SolidBrush,
        "StringAlignment": StringAlignment,
        "StringFormatFlags": StringFormatFlags,
        "StringTrimming": StringTrimming,
        "Encoder": Encoder,
        "EncoderParameter": EncoderParameter,
        "EncoderParameters": EncoderParameters,
        "EncoderValue": EncoderValue,
        "ImageCodecInfo": ImageCodecInfo,
        "ImageFormat": ImageFormat,
        "TextRenderingHint": TextRenderingHint,
    }
    workbook = load_workbook(source, data_only=True)
    bitmaps = []
    try:
        for sheet in workbook.worksheets:
            if not _sheet_is_used(sheet):
                continue
            bitmaps.append(_render_sheet(sheet, drawing))
    finally:
        workbook.close()
    if not bitmaps:
        raise SystemExit(f"No used sheets in {source}.")
    save_tiff(bitmaps, output, drawing)


def main(argv: list[str] | None = None) -> None:
    parser = argparse.ArgumentParser(description="Convert an Excel form to a TIFF image.")
    parser.add_argument("input", type=Path, help="Source .xlsx workbook")
    parser.add_argument("output", type=Path, help="Destination .tif file")
    args = parser.parse_args(argv)
    if not args.input.is_file():
        raise SystemExit(f"Workbook not found: {args.input}")
    convert(args.input, args.output)


if __name__ == "__main__":
    main(sys.argv[1:])
