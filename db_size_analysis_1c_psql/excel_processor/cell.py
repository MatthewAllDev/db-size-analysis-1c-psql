from datetime import datetime

import openpyxl.styles as styles
from openpyxl.cell import Cell as BaseCell
from openpyxl.worksheet.worksheet import Worksheet


class Cell(BaseCell):
    def __init__(self, sheet: Worksheet, row: int, column: int, data: str or dict):
        """
        :param data: value string or excel processor format dict
        """
        self.parent_cell: BaseCell = sheet.cell(row, column)
        super(Cell, self).__init__(self.parent_cell.parent,
                                   self.parent_cell.row,
                                   self.parent_cell.column,
                                   self.parent_cell.value,
                                   self.parent_cell._style)
        self.set(data)

    def set(self, data):
        """
        :param data: value string or excel processor format dict
        """
        if type(data) == str:
            self.parent_cell.value = data
        elif type(data) == int:
            self.parent_cell.value = str(data)
        elif type(data) == datetime:
            self.parent_cell.value = str(data)
        elif type(data) == dict:
            self.parent_cell.value = data['value']
            if data.get('color'):
                self.paint(data.get('color'))
            if data.get('font'):
                self.font_style(**data.get('font'))
            if data.get('border'):
                self.set_borders(**data.get('border'))
            if data.get('wrap_text'):
                self.set_wrap_text()

    def font_style(self, **kwargs):
        """
        :Keyword Arguments:
        * *name*
        * *size*
        * *bold*
        * *italic*
        * *charset*
        * *underline*
        * *strike*
        * *color*
        * *scheme*
        * *family*
        * *vertAlign*
        * *outline*
        * *shadow*
        * *condense*
        * *extend*
        """
        self.parent_cell.font = styles.Font(**kwargs)
        return self

    def set_borders(self,
                    style: str = 'thin',
                    color: str = '00000000',
                    left_style: str = None,
                    right_style: str = None,
                    top_style: str = None,
                    bottom_style: str = None):
        color: styles.colors.Color = styles.colors.Color(rgb=color)
        left_style: str = style if left_style is None else left_style
        right_style: str = style if right_style is None else right_style
        top_style: str = style if top_style is None else top_style
        bottom_style: str = style if bottom_style is None else bottom_style
        self.parent_cell.border = styles.borders.Border(left=styles.borders.Side(style=left_style, color=color),
                                                        right=styles.borders.Side(style=right_style, color=color),
                                                        top=styles.borders.Side(style=top_style, color=color),
                                                        bottom=styles.borders.Side(style=bottom_style, color=color))
        return self

    def paint(self, color: str, fill_type: str = 'solid'):
        self.parent_cell.fill = styles.PatternFill(start_color=color, end_color=color, fill_type=fill_type)

    def set_wrap_text(self):
        self.parent_cell.alignment = self.parent_cell.alignment.copy(wrapText=True)
