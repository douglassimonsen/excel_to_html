import openpyxl
import static_values
import color_utilities


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
    def __init__(self, cell, wb_meta):

        self.text = cell.value
        self.font_style = self.handle_font_style(cell, wb_meta['themes'])
        self.border_style = self.handle_border_style(cell, wb_meta['themes'])
        # self.rowspan, row.colspan = self.handle_merged_cells(cell)


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
    wb_meta = {
        'themes': color_utilities.get_theme_colors(wb)
    }
    ws = wb['Sheet1']
    parsed_sheet = []
    for row in list(ws.iter_rows()):
        parsed_row = []
        for cell in row:
            parsed_row.append(ParsedCell(cell, wb_meta))
        parsed_sheet.append(parsed_row)



main("test.xlsx")
