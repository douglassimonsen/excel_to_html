import openpyxl
import static_values
import color_utilities
import math


def handle_color(color, themes):
    if color is None:
        return None
    if color.type == 'indexed':
        if color.indexed in (64, 65):  # see COLOR_INDEX comment about 64/65
            return None
        return '#' + static_values.COLOR_INDEX[color.indexed][2:]
    elif color.type == 'rgb':
        return '#' + color_utilities.rgb_and_tint_to_hex(color.rgb, color.tint)
    elif color.type == 'theme':
        rgb = themes[color.theme]
        return '#' + color_utilities.rgb_and_tint_to_hex(rgb, color.tint)
    elif color.type == 'auto':
        return '#000000'
    else:
        return None


class ParsedCell:
    def __init__(self, cell, ws_meta, row_idx, col_idx):

        self.text = cell.value
        self.font_style = self.handle_font_style(cell, ws_meta['themes'])
        self.border_style = self.handle_border_style(cell, ws_meta['themes'])
        self.rowspan, self.colspan = self.handle_merged_cells(cell, ws_meta['merged_cell_ranges'])
        self.width = ws_meta['column_widths'].get(col_idx, ws_meta['default_col_width'])
        self.height = ws_meta['row_heights'].get(col_idx, ws_meta['default_row_height'])

    @staticmethod
    def handle_merged_cells(cell, merged_cell_ranges):
        for merge_range in merged_cell_ranges:
            if cell.coordinate in merge_range:
                rowspan = merge_range.max_row - merge_range.min_row + 1
                colspan = merge_range.max_col - merge_range.min_col + 1
                return rowspan, colspan
        return 1, 1

    @staticmethod
    def handle_border_style(cell, themes):
        ret = {}
        if (cell.border.top == cell.border.left) and (cell.border.left == cell.border.bottom) and (cell.border.bottom == cell.border.right):  # the borders are all the same
            border = cell.border.top
            border_color = handle_color(border.color, themes)
            border_width = static_values.border_style_to_width.get(border.style, '0px')
            border_style = static_values.border_style_to_style.get(border.style, '0px')
            ret['border'] = f'{border_width} {border_style} {border_color}'
        else:
            for side in ['top', 'right', 'bottom', 'left']:
                border = getattr(cell.border, side)
                border_color = handle_color(border.color, themes)
                border_width = static_values.border_style_to_width.get(border.style, '0px')
                border_style = static_values.border_style_to_style.get(border.style, '0px')
                ret[f'border-{side}'] = f'{border_width} {border_style} {border_color}'
        return ret

    @staticmethod
    def handle_font_style(cell, themes):
        ret = {}
        if cell.font.i:  # italics
            ret['font-style'] = 'italic'
        if cell.font.b:  # bold
            ret['font-weight'] = 'bold'
        if cell.font.u:  # underline
            ret['text-decoration'] = 'underline'
        ret['font-family'] = f"'{cell.font.name}'"
        ret['font-size'] = f"{cell.font.sz}px"

        background_color = handle_color(cell.fill.fgColor, themes) or handle_color(cell.fill.bgColor, themes)  # foreground color will show above background, right?
        if background_color is not None:
            ret['background-color'] = background_color

        font_color = handle_color(cell.font.color, themes)
        if font_color is not None:
            ret['color'] = font_color
        return ret


def main(pathname):
    wb = openpyxl.open(pathname)
    ws = wb['Sheet1']
    ws_meta = {
        'themes': color_utilities.get_theme_colors(wb),
        'merged_cell_ranges': ws.merged_cells.ranges,
        'column_widths': {openpyxl.utils.cell.column_index_from_string(i): math.ceil(x.width * 7) for i, x in ws.column_dimensions.items()},  # converting excel units to pixels
        'default_col_width': ws.sheet_format.defaultColWidth or 64,
        'row_heights': {(i - 1): x.height * (4 / 3) for i, x in ws.row_dimensions.items()},  # converting excel units to pixels
        'default_row_height': ws.sheet_format.defaultRowHeight or 20,
    }
    parsed_sheet = []
    for i, row in enumerate(ws.iter_rows()):
        parsed_row = []
        for j, cell in enumerate(row):
            if isinstance(cell, openpyxl.cell.cell.Cell):
                parsed_row.append(ParsedCell(cell, ws_meta, i, j))
        parsed_sheet.append(parsed_row)



main("test.xlsx")
