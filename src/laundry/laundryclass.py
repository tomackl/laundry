"""Main class for laundry. This is intended to replace the original laundry script."""

from laundry.constants import data_frame, invalid, photo_formats
from typing import Dict, List, Iterable, Tuple
from docx import Document
from docx.shared import Inches
from pathlib import Path, PurePath
import pandas as pd


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


def sort_table_data(record: Dict, header: List[str], format_title: bool = True) -> List[Tuple]:
    """
    Take a dictionary and split in to a list of tuples containing the 'keys' cell_data defined in 'header' as the first tuple
    and the associated values as the second. The function will return a list containing two equal length tuples. The
    function assumes that the cell_data contained in record and header is in the correct order.
    :param record: The dictionary containing the cell_data.
    :param header: The keys that defined the key-values to be extracted.
    :param format_title: If True make the header string title case.
    :return: A list of tuples containing the header and cell_data information.
    """
    hdr_list: List[str] = []
    data_list: List[str] = []
    for each in header:
        hdr_data = remove_underscore(each)
        if format_title is True:
            hdr_data = hdr_data.title()
        hdr_list.append(hdr_data)
        data_list.append(record.pop(each))
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
    path = PurePath()
    for each in filepath:
        q = PurePath(each.replace('\\', '/').strip('/'))
        path = path / Path(q.as_posix())
    r = Path(path)
    if r.is_dir():
        return r
    return 'Incorrect path.'


class SingleLoad:
    """
    This class is intended to replace the original Laundry's procedural approach from the single load function.
    The click.progressbar() will not be transferred from the original function.
    Assumptions
    1. The file paths passed are correct and have already been checked.
    2. Data passed to the class is in the correct format.
    3. The output file will be created by the object.
    """
    def __init__(self, structure_dict: Dict, data_dict: Dict, file_template: str, spreadsheet_fp: (Path, str),
                 file_output_path: (Path, str)):
        """
        # The method signature is based on the laundry.single_load() function. This calls self.format_docx()
        :param structure_dict: A dictionary that defines the structure of the documentation. 
        :param data_dict: A dictionary that contains the cell_data to be formatted.
        :param file_template: The Word .docx file that contains the formatting styles to be used.
        :param spreadsheet_fp: The path to the directory containing the spreadsheet.
        :param file_output_path: The path to the output file location.
        """
        self._structure: dict = structure_dict
        self._data: dict = data_dict
        self._file_template: str = file_template   # todo: This should probably be a Path()
        self._spreadsheet_fp: (Path, str) = spreadsheet_fp
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
        for each in self._row_data:
            self.format_docx(each)

    def format_docx(self, row: dict):
        """
        This is factory method that calls the appropriate the information contained within document structure.
        :param row: dictionary containing the data_str to be formatted. This is a single row from the spreadsheet.
        :return:
        """
        # row => each passed from self.start_wash()
        # structdict => self._structure
        # output_document => self._output_docx
        # outputfile => self._file_template
        # input_file_path => self._spreadsheet_fp

        for element in self._structure:
            ele_sect_contains: str = str(element['sectioncontains']).lower()
            ele_sect_style: str = element['sectionstyle']
            ele_sect_type: str = str(element['sectiontype']).lower()
            ele_title_style: str = element['titlestyle']
            ele_sect_break: bool = element['sectionbreak']
            ele_page_break: bool = element['pagebreak']
            ele_row_contains = row[ele_sect_contains]

            if ele_sect_type in ('heading', 'para', 'paragraph'):
                self.insert_paragraph(str(ele_row_contains).lower(), title=ele_sect_contains.title(),
                                      section_style=ele_sect_style, title_style=ele_title_style)

            elif ele_sect_type == 'table':
                table_col_hdr = split_str(ele_sect_contains)
                sorted_row = sort_table_data(row, table_col_hdr)
                self.insert_table(len(table_col_hdr),
                                  len(sorted_row),
                                  sorted_row,
                                  section_style=ele_sect_style,
                                  )

            elif ele_sect_type == 'photo':
                q = confirm_directory_path([self._spreadsheet_fp, element['path']])
                if str(ele_row_contains).lower() not in ['no photo', 'none', 'nan', '-']:
                    photo = split_str(ele_row_contains)
                    for each in photo:
                        loc = q.joinpath(each)
                        self.insert_photo(str(loc), 4)
            else:
                print('Valid section header was not found.')

            if ele_sect_break is True:
                self.insert_paragraph(self._output_docx, '')

            if ele_page_break is True:
                self._output_docx.add_page_break()

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
                     autofit_table: bool = True):
        """
        The function takes data and uses it to create a table for the template_doc.
        The first row of data is assumed to be the table header.
        todo: add 'add_table_hdr added to allow the table header to be dropped from the table if not required.
        :param rows: the number of required table rows.
        :param cols: the number of required table columns.
        :param data: The list data to be inserted into the table. The idx[0] is assumed to be the header.
        :param section_style: The style to be used for the table.
        :param autofit_table: autofit the table to the page width.
        :return:
        """
        table = self._output_docx.add_table(rows=rows, cols=cols, style=section_style)
        table.autofit = autofit_table
        cell_data = enumerate(data, 0)
        for i, cell_contents in cell_data:
            for j, text in enumerate(cell_contents):
                table.rows[i].cells[j].text = str(text)

    def insert_photo(self, photo: str, width: int = 4):
        """
        Insert a photo located at path into document and set the photo width.
        :param photo: file path to the
        :param width: width of the image in Inches
        :return:
        """
        for ext in photo_formats:
            photo_path = Path(photo + ext)
            if photo_path.exists():
                self._output_docx.add_picture(str(photo_path), width=Inches(width))
                return True
        print(f'\nPhoto {photo} does not exist. Check file extension.')
        self._output_docx.add_paragraph(f'PHOTO "{str(photo).upper()}" NOT FOUND\n')

    def issue_document(self):
        """
        Output the file.
        :return:
        """
        self._file_template.save(self._output_docx)


class Laundry:
    """
    """
    def __init__(self, input_fp: Path, data: str, structure: str = None, batch: str = None):
        """

        :param input_fp: The file path to the spreadsheet containing the data
        :param data: The name of the worksheet containing the data to be formatted.
        :param structure: The name of the worksheet containing the output document's structure.
        :param batch: The name of the worksheet containing the batch data.
        """
        if data is None and structure is None and batch is None:
            raise TypeError(f'Name of either the "structure" or the "batch" worksheets must be provided.')
        self._data_wksht: str = data
        self._structure_wksht: str = structure
        self._batch_wksht: str = batch
        # Set up the lists to contain the dictionaries containing: 1. data, 2. output structure, and 3. batch docs
        self._data: List[dict] = []
        self._structure: List[dict] = []
        self._batch: List[dict] = []
        self._input_fp: Path = input_fp
        # todo: add check for input_fp

        # todo: 1. add the keywords listed below as __init__ parameters. These will need to converted into a dictionary
        # todo:     and added to the self._batch for single file operations.
        # todo: 2. these should be used to check the batch spreadsheet to ensure that the correct col hdrs exist.
        # data_worksheet
        # structure_worksheet
        # header_row
        # remove_columns
        # drop_empty_columns
        # template_file
        # filter_rows
        # output_file

    def worksheets_exist(self) -> bool:
        """
        Check the expected worksheets exist.
        :return:
        """
        pass

    def clean_data_excel(self) -> (dict, None):
        """
        Clean the data and get it into the correct format.
        :return:
        """
        pass

    def file_path_exists(self) -> bool:
        """
        Confirm that the passed file exists on the file system.
        :return:
        """
        pass

    def single_load(self) -> SingleLoad:
        """
        Create a Laundry object.
        :return:
        """
        pass

