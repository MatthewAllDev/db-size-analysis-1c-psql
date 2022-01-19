from openpyxl.worksheet.worksheet import Worksheet

from .cell import Cell


class Row:
    def __init__(self, sheet: Worksheet, index: int):
        self.__sheet: Worksheet = sheet
        self.index: int = index
        self.cells: tuple = ()
        self.__current_column_index: int = 1
        # if cells is None:
        #     self.cells: tuple = ()
        # else:
        #     self.cells: tuple = cells

    def set(self, *args):
        """
        :param args: arg is value string or excel processor format dict
        """
        for element in args:
            self.add(element)

    def add(self, data: str or dict) -> Cell:
        """
        :param data: value string or excel processor format dict
        """
        cell: Cell = Cell(self.__sheet, self.index, self.__current_column_index, data)
        self.__current_column_index += 1
        self.cells += (cell, )
        return cell

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
        for cell in self.cells:
            cell: Cell
            cell.font_style(**kwargs)
        return self

    def set_borders(self,
                    style: str = 'thin',
                    color: str = '00000000',
                    left_style: str = None,
                    right_style: str = None,
                    top_style: str = None,
                    bottom_style: str = None):
        for index, cell in enumerate(self.cells):
            cell: Cell
            if index == 0:
                ls: str = left_style
                rs: str = style
            elif index == len(self.cells) - 1:
                ls: str = style
                rs: str = right_style
            else:
                ls: str = style
                rs: str = style
            cell.set_borders(style, color, ls, rs, top_style, bottom_style)
        return self

    def paint(self, color: str, fill_type: str = 'solid'):
        for cell in self.cells:
            cell: Cell
            cell.paint(color, fill_type)

    def set_wrap_text(self):
        for cell in self.cells:
            cell.set_wrap_text()
