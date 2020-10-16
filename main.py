import openpyxl
import static_values
import color_utilities
import math
import jinja2
# todo: handle case when area cuts a merged cell in half
# todo: include/remove alpha channel option in handle_color
# todo: exclude implicit borders that lines up with explicit border, otherwise they can be covered up
# todo: fails with zero rows


def handle_color(color, themes, alpha=False):
    if color is None:
        return None
    if color.type == 'indexed':
        if color.indexed in (64, 65):  # see COLOR_INDEX comment about 64/65
            return None
        color = static_values.COLOR_INDEX[color.indexed]
    elif color.type == 'rgb':
        color = '00' + color_utilities.rgb_and_tint_to_hex(color.rgb, color.tint)
    elif color.type == 'theme':
        rgb = themes[color.theme]
        color = '00' + color_utilities.rgb_and_tint_to_hex(rgb, color.tint)
    elif color.type == 'auto':
        color = '00000000'
    else:
        return None

    if not alpha:
        color = color[2:]
    return '#' + color


class ParsedCell:
    def __init__(self, cell, ws_meta, row_idx, col_idx):

        self.text = cell.value or ''
        self.font_style = self.handle_font_style(cell, ws_meta['themes'])
        self.border_style = self.handle_border_style(cell, ws_meta['themes'])
        self.rowspan, self.colspan = self.handle_merged_cells(cell, ws_meta['merged_cell_ranges'])

        self.sizing_style = self.handle_sizing(cell, ws_meta, row_idx, col_idx)

    @staticmethod
    def handle_sizing(cell, ws_meta, row_idx, col_idx):
        ret = {}
        width = ws_meta['column_widths'].get(col_idx, ws_meta['default_col_width'])
        height = ws_meta['row_heights'].get(row_idx, ws_meta['default_row_height'])
        ret['width'] = str(width) + 'px'
        ret['height'] = str(height) + 'px'
        horizontal = cell.alignment.horizontal or 'left'
        vertical = cell.alignment.vertical or 'bottom'
        if vertical == 'center':
            vertical = 'middle'
        ret['text-align'] = horizontal
        ret['vertical-align'] = vertical
        return ret

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

            border_color = handle_color(border.color, themes) or '#000000'

            border_width = static_values.border_style_to_width.get(border.style)
            border_style = static_values.border_style_to_style.get(border.style, '0px')
            if border_width is not None:
                ret['border'] = f'{border_width} {border_style} {border_color}'
            else:
                ret['border'] = '1px solid #D9D9D9'
        else:
            for side in ['top', 'right', 'bottom', 'left']:
                border = getattr(cell.border, side)

                border_color = handle_color(border.color, themes) or '#000000'

                border_width = static_values.border_style_to_width.get(border.style)
                border_style = static_values.border_style_to_style.get(border.style, '0px')
                if border_width is not None:
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

        if cell.fill.patternType is not None:
            background_color = handle_color(cell.fill.fgColor, themes) or handle_color(cell.fill.bgColor, themes)  # foreground color will show above background, right?
            if background_color is not None:
                ret['background-color'] = background_color

        font_color = handle_color(cell.font.color, themes)
        if font_color is not None:
            ret['color'] = font_color
        return ret

    def get_style(self):
        style = []
        for k, v in self.font_style.items():
            style.append(f'{k}: {v}')
        for k, v in self.border_style.items():
            style.append(f'{k}: {v}')
        for k, v in self.sizing_style.items():
            style.append(f'{k}: {v}')
        return '; '.join(style)


def to_html(parsed_sheet):
    return jinja2.Template('''
        <table style="border-collapse:collapse">
            {% for row in parsed_sheet %}
                <tr style="height: {{row[0].height}}">
                    {% for cell in row %}
                        <td style="{{cell.get_style()}}" rowspan={{cell.rowspan}} colspan={{cell.colspan}}>
                            {{cell.text}}
                        </td>
                    {% endfor %}
                </tr>
            {% endfor %}
        </table>
    ''').render(parsed_sheet=parsed_sheet)


def main(pathname, min_row=None, max_row=None, min_col=None, max_col=None, openpyxl_kwargs={}):
    wb = openpyxl.open(pathname, **openpyxl_kwargs)
    ws = wb['Sheet1']
    ws_meta = {
        'themes': color_utilities.get_theme_colors(wb),
        'merged_cell_ranges': ws.merged_cells.ranges,
        'column_widths': {(openpyxl.utils.cell.column_index_from_string(i) - 1): math.ceil(x.width * 7) for i, x in ws.column_dimensions.items()},  # converting excel units to pixels
        'default_col_width': ws.sheet_format.defaultColWidth or 64,
        'row_heights': {(i - 1): x.height * (4 / 3) for i, x in ws.row_dimensions.items()},  # converting excel units to pixels
        'default_row_height': ws.sheet_format.defaultRowHeight or 20,
    }
    parsed_sheet = []
    for i, row in enumerate(ws.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col)):
        parsed_row = []
        for j, cell in enumerate(row):
            if isinstance(cell, openpyxl.cell.cell.Cell):
                parsed_row.append(ParsedCell(cell, ws_meta, i, j))
        parsed_sheet.append(parsed_row)
    body = to_html(parsed_sheet)
    with open('test.html', 'w') as f:
        f.write(body)
    return body


main("test.xlsx", min_row=4)
