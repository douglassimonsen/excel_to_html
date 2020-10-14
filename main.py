import openpyxl


class ParsedCell:
    def __init__(self, cell):
        self.text = cell.value
        self.italics = cell.font.i
        self.bold = cell.font.b
        self.underline = cell.font.u is not None  # single, double all get translated to underline
        self.text_color = cell.font.color.rgb
        self.background_color = cell.fill.bgColor.rgb
        self.font = cell.font.name
        print(dir(cell.fill))
        print(cell.font.name)
        exit()


def main(ws):
    parsed_sheet = []
    for row in list(ws.iter_rows())[:]:
        parsed_row = []
        for cell in row[2:]:
            parsed_row.append(ParsedCell(cell))
        parsed_sheet.append(parsed_row)


wb = openpyxl.open("test.xlsx")
ws = wb['Sheet1']
main(ws)
