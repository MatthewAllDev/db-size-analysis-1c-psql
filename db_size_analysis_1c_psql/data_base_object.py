import json

from .excel_processor import ExcelProcessor
from .safe_list import SafeList


class DataBaseObject:
    name: str
    children: dict
    size: int
    namespace: str
    parent_object: str
    object: str
    table_name_dbms: str
    table_name_1c: str
    metadata: str
    purpose: str
    fields: dict
    index: dict

    def __init__(self, name: str = None, size: int = 0, **kwargs):
        super().__init__()
        self.name: str = name
        self.size: int = size
        self.children: dict = dict()
        self.set_attributes(**kwargs)

    def get(self, key, default=None):
        return self.children.get(key, default)

    def set_attributes(self, **kwargs):
        for key, value in kwargs.items():
            if value is not None:
                if not (type(value) == dict and len(value) == 0):
                    self.__setattr__(key, value)

    def get_attribute(self, name: str, default=None):
        try:
            return self.__getattribute__(name)
        except AttributeError:
            return default

    def add_data_level(self, keys: SafeList, db_object):
        try:
            key: str = keys.pop(0)
        except IndexError:
            self.set_attributes(**db_object.__dict__)
            return
        self.size += db_object.size
        if self.get(key) is None:
            self[key]: DataBaseObject = DataBaseObject(f'{self.name}.{key}')
        self[key].add_data_level(keys, db_object)

    def sort(self, key: str, reverse: bool = False):
        if len(self.children) == 0:
            return
        sorted_children: list = sorted(self.children.items(),
                                       key=lambda element: element[1].get_attribute(key),
                                       reverse=reverse)
        self.children = dict(sorted_children)
        for child in self.children.values():
            child.sort(key, reverse)

    def print(self, max_level: int = 99, level: int = 0):
        print('\t' * level + str(self))
        if level == max_level:
            return
        for child in self.children.values():
            child.print(max_level, level + 1)

    def export_to_json(self, file_path: str):
        with open(file_path, "w", encoding='utf-8') as write_file:
            json.dump(self, write_file, default=lambda x: x.__dict__, indent=4, ensure_ascii=False)

    def export_to_excel(self, file_path: str = None, excel: ExcelProcessor = None, tree_level: int = 1) \
            -> str or tuple:
        if excel is None:
            if file_path is None:
                raise TypeError('file_path or excel must be defined')
            excel: ExcelProcessor = ExcelProcessor()
            excel.add_row('Name',
                          'Size',
                          'Namespase',
                          'DBMS table name',
                          '1C table name',
                          'Metadata',
                          'Purpose').font_style(size=16, bold=True, italic=True).set_wrap_text()
        row = excel.add_row(*self.__get_excel_row())
        children_index: SafeList = SafeList()
        first_index = last_index = row.index
        for child in self.children.values():
            child_row, child_index = child.export_to_excel(excel=excel, tree_level=tree_level + 1)
            last_index = child_row.index
            children_index.append(child_index)
        if first_index != last_index:
            first_index += 1
        if len(self.children) > 0:
            row.font_style(size=18 - tree_level * 2, italic=True)
            last_index = excel.add_row().index
        else:
            row.font_style(size=18 - tree_level * 2)
        if tree_level == 1:
            excel.set_optimal_column_widths()
            excel.create_rows_tree(first_index, last_index, children_index)
            return excel.save(file_path)
        else:
            return row, {'start': first_index, 'end': last_index, 'children_index': children_index}

    def get_format_size(self) -> str:
        prefix_list: list = ['', 'k', 'M', 'G', 'T', 'P', 'E', 'Z', 'Y']
        index: int = 0
        while self.size > 1024:
            self.size /= 1024
            index += 1
        return f'{round(self.size, 2)} {prefix_list[index]}B'

    def __get_excel_row(self) -> tuple:
        return self.get_attribute('name', ''), \
               self.get_format_size(), \
               self.get_attribute('namespase', ''), \
               self.get_attribute('table_name_dbms', ''), \
               self.get_attribute('table_name_1c', ''), \
               self.get_attribute('metadata', ''), \
               self.get_attribute('purpose', '')

    def __str__(self):
        return f'{self.name} | {self.get_format_size()}'

    def __getitem__(self, item):
        return self.children[item]

    def __setitem__(self, key, value):
        self.children[key] = value
