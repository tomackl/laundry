"""Main class for laundry. This is intended to replace the original laundry script."""

from laundry.constants import data_frame, invalid, photo_formats
from typing import Dict, List, Iterable, Tuple
from docx import Document
from docx.shared import Inches
from pathlib import Path, PurePath


structure_keys = ('sectiontype', 'sectioncontains', 'sectionstyle', 'titlestyle', 'sectionbreak', 'pagebreak', 'path')
batch_keys = ('data_worksheet', 'structure_worksheet', 'header_row', 'remove_columns', 'drop_empty_columns',
              'template_file', 'filter_rows', 'output_file')


def split_str(data_str: str) -> List[str]:
    """
    Split a string into a list of strings. The string will be split at ',' or '\n' with '\n' taking precedence.
    :param data_str: A string to be split.
    :return: List[str]
    """
    if '\n' in data_str:
        return list(str(data_str).splitlines())
    elif ',' in data_str:
        return list(str(data_str).split(','))
    else:
        i = list()
        i.append(data_str)
        return i


def extract_data(record: Dict, header: List[str], format_title: bool = True) -> List[Tuple]:
    """
    Take a dictionary and split in to a list of tuples containing the 'keys' data defined in 'header' as the first tuple
    and the associated values as the second. The function will return a list containing two equal length tuples. The
    function assumes that the data contained in record and header is in the correct order.
    :param record: The dictionary containing the data.
    :param header: The keys that defined the key-values to be extracted.
    :param format_title: If True make the header string title case.
    :return: A list of tuples containing the header and data information.
    """
    hdr_list: List[str] = []
    data_list: List[str] = []
    for each in header:
        hdr_data = remove_underscore(each)
        if format_title is True:
            hdr_list.append(hdr_data.title())
        else:
            hdr_list.append(hdr_data)
        t = record.pop(each)
        data_list.append(t)
    return [tuple(hdr_list), tuple(data_list)]


def remove_underscore(data_str: str) -> str:
    """
    A function to remove underscores from a string returning the string with spaces instead.
    :param data_str:
    :return:
    """
    return data_str.replace('_', ' ')


def confirm_directory_path(filepath: List[str]) -> (Path, str):
    """
    Convert the contents of the passed list into a Path. This function assumes
    that the sum of the passed list will be a single path to a directory.
    :param filepath: a list of path names as string
    :return: Path
    """
    filepath = filepath
    p = PurePath()
    for each in filepath:
        q = PurePath(each.replace('\\', '/').strip('/'))
        p = p / Path(q.as_posix())
    r = Path(p)
    if r.is_dir():
        return r
    return 'Incorrect path.'


class Laundry:
    """
    This class is intended to replace the original Laundry's procedural approach from the single load function.
    The click.progressbar() will not be transferred from the original function.
    Assumptions
    1. The file paths passed are correct and have already been checked.
    2. Data passed to the class is in the correct format.
    3. The output file will be created by the object.
    """
    def __init__(self, structure_dict: Dict, data_dict: Dict, file_template: str, file_path: (Path, str),
                 file_output_path: (Path, str)):
        """
        # The method signature is based on the laundry.single_load() function. This calls self.format_docx()
        :param structure_dict: A dictionary that defines the structure of the documentation. 
        :param data_dict: A dictionary that contains the data to be formatted.
        :param file_template: The Word .docx file that contains the formatting styles to be used.
        :param file_path: The path to the directory containing the spreadsheet.
        :param file_output_path: The path to the output file location.
        """
        self._structure: dict = structure_dict
        self._data: dict = data_dict
        self._file_template: str = file_template   # todo: This should probably be a Path()
        self._file_path: (Path, str) = file_path
        self._file_output: (Path, str) = file_output_path
        self._output_docx = Document()
        self._row_data: List[Dict] = list()
        self.split_into_rows()

    def split_into_rows(self):
        """
        Split self._data and pass each row to self.format_docx.
        """
        self._row_data = list()
        for each in self._data:
            self._row_data.append({each: self._data[each]})

    def start_wash(self):
        """
        Start formatting the output document.
        """
        # for each in self._row_data:
        #     self.format_docx(each)
        pass

    def format_docx(self, rowdict: dict, structdict: dict, output_document: Document(), file_path_input: str,
                    file_output_path: str):
        """
        This is factory method that calls the appropriate the information contained within document structure.
        # The method signature is based on the laundry.format_docx() function. This calls:
        # - self.insert_paragraph()
        # - self.insert_table()
        # - self.insert_photo()
        :param rowdict: dictionary containing the data_str to be formatted. This is a single row from the spreadsheet/
        :param structdict: defines the output file's format structure.
        :param output_document: The file into which the data will be inserted into.
        :param file_path_input: The directory containing the spreadsheet. Resources are referenced from this directory.
        :param file_output_path: The path of the output file. The output file will be saved here.
        :return:
        """
        # rowdict => rowdict
        # structdict => self._structure
        # output_file => self._file_template
        # input_file_path => self._file_path
        pass

    def insert_paragraph(self, text: str, title: str = None, section_style: str = None,
                         title_style: str = None):
        """
        This method defines the characteristics of a paragraph to be added to the document.
        :param text: paragraph data_str
        :param title: title data_str
        :param section_style: paragraph style
        :param title_style: title style. Use this _or_ title_level.
        :return:
        """
        text = text.splitlines()
        if text == '':
            self._output_docx.add_paragraph(text)
        else:
            if (title is not None) and (str(title_style) != 'nan'):
                self._output_docx.add_paragraph(remove_underscore(title), style=title_style)
            for each in text:
                self._output_docx.add_paragraph(each, style=section_style)

    def insert_table(self, cols: int, rows: int, data: List[Iterable[str]], section_style: str = None,
                     autofit: bool = True):
        """
        The function takes data and uses it to create a table for the template_doc.
        The first row of data is assumed to be the table header.
        20190814 'add_table_hdr added to allow the table header to be dropped from the table if not required.
        :param rows: the number of required table rows.
        :param cols: the number of required table columns.
        :param data: The list data to be inserted into the table. The idx[0] is assumed to be the header.
        :param section_style: The style to be used for the table.
        :param autofit: autofit the table to the page width.
        :return:
        """
        table = self._output_docx.add_table(rows=rows, cols=cols, style=section_style)
        table.autofit = autofit
        data = enumerate(data, 0)
        for i, cell_contents in data:
            self._insert_row(table.rows[i].cells, cell_contents)

    @staticmethod
    def _insert_row(row_cells, data: List[str]):
        """
        Populate a table row. The cells are passed as a row and the contents added.
        :param row_cells:
        :param data:
        :return:
        """
        for i, text in enumerate(data):
            row_cells[i].text = str(text)
        return row_cells

    def insert_photo(self, photo: str, width: int = 4):
        """
        Insert a photo located at path into document and set the photo to width.
        :param photo: file path to the
        :param width: width of the image in Inches
        :return:
        """
        for ext in photo_formats:
            photo_path = Path(photo + ext)
            if photo_path.exists():
                self._output_docx.add_picture(str(photo_path), width=Inches(width))
                return True
        print('\nPhoto {} does not exist. Check file extension.'.format(photo))
        self._output_docx.add_paragraph('PHOTO "{}" NOT FOUND\n'.format(str(photo).upper()))
