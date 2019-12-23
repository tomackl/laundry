"""Main class for laundry. This is intended to replace the original laundry script."""

from laundry.constants import data_frame, invalid, photo_formats
from typing import Dict, List, Iterable, Tuple
from docx import Document
from docx.shared import Inches
from pathlib import Path, PurePath
import pandas as pd

old_structure_keys = ('sectiontype', 'sectioncontains', 'sectionstyle', 'titlestyle', 'sectionbreak', 'pagebreak',
                      'path')
structure_keys = ('section_type', 'section_contains', 'section_style', 'title_style', 'section_break', 'page_break',
                  'path')
batch_keys = ('data_worksheet', 'structure_worksheet', 'header_row', 'remove_columns', 'drop_empty_rows',
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


def resolve_file_path(path: (Path, str)) -> (Path, Exception):
    """
    Resolve the passed filepath returning as a Path() or an exception.
    :return:
    """
    p = Path(path)
    return p.resolve(strict=True)


def values_exist(expected: set, actual: set) -> bool:
    """
    Check the expected worksheets exist.
    :return:
    """
    return expected <= actual


def remove_from_iterable(values: Iterable, *args) -> List:
    """Return items in list that are not equal to a value in drop."""
    data = []
    for each in values:
        if each not in args:
            data.append(each)
    return data


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
        self._file_template: str = file_template  # todo: This should probably be a Path()
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

    def __init__(self, input_fp: Path, data_worksheet: str = None, structure_worksheet: str = None,
                 batch_worksheet: str = None, header_row: int = 0, remove_columns: str = None,
                 drop_empty_rows: bool = None, template_file: str = None, filter_rows: str = None,
                 output_fp: (Path, str) = None):
        """
        Instantiating the class will run error checking on the passed information, checking for the following steps:
        1. A basic check that worksheet names have been passed.
        2. The input file's path is correct.
        3. Check that the data, structure and batch worksheet names passed at instantiation exist within the file. This
            is the first worksheet name check, additional check will be made as required.
        4. If the batching information is passed to the object at instantiation, then merge this into a dictionary.
            Error checking will be completed later.
        5.
        :param input_fp: The file path to the spreadsheet containing the data
        :param data_worksheet: The name of the worksheet containing the data to be formatted.
        :param structure_worksheet: The name of the worksheet containing the output document's structure.
        :param batch_worksheet: The name of the worksheet containing the batch data.
        :param header_row: 
        :param remove_columns: 
        :param drop_empty_rows: An explicit tag to drop empty rows from the worksheet if they contain two or more empty
        cells. If this is left as None it will be automatically set to True for the data worksheet.
        :param template_file: 
        :param filter_rows: 
        :param output_fp: 
        """

        # Step 1: Basic data checking.
        t_sheets_expected = remove_from_iterable([data_worksheet, structure_worksheet, batch_worksheet], None)
        try:
            print(f'Check 1: Worksheets are present.')
            self.check_worksheets_basic(t_sheets_expected)
        except Exception as e:
            print(e)
        else:
            print(f'\tWorksheets present: {t_sheets_expected}.')

        # Step 2: Confirm the input file exists.
        self._input_fp: (Path, str) = ''
        try:
            print(f'Check 2:Resolving spreadsheet filepath.')
            self._input_fp = resolve_file_path(input_fp)
        except Exception as e:
            print(f'\t{e}: File {input_fp} does not exist.')
        else:
            print(f'\t{self._input_fp}')

        # Load the Excel file into memory.
        self._washing_basket = pd.ExcelFile(self._input_fp)

        # Gather the worksheet names
        self._sheets_actual: list = self._washing_basket.sheet_names

        # Set up the lists to contain the dictionaries containing: 1. data, 2. output structure, and 3. batch docs
        self._data: List[dict] = []
        self._structure: List[dict] = []
        self._batch: List[dict] = []

        # Define headers for the batch and structure worksheets. These are fixed.
        self.expected_batch_headers = ['data_worksheet', 'structure_worksheet', 'header_row', 'remove_columns',
                                       'drop_empty_rows', 'template_file', 'filter_rows', 'output_file']
        self.expected_structure_headers = ['section_type', 'section_contains', 'section_style', 'title_style', 'section_break',
                                           'page_break', 'path']

        #  Step 3: Check that the data, structure and batch worksheet names passed exist within the file.
        try:
            print(f'Check 3: Worksheets exist in spreadsheet')
            self.compare_lists(t_sheets_expected, self._sheets_actual)
        except Exception as e:
            print(f'{e}')
        else:
            print(f'\tWorksheets present: {self._sheets_actual}.')

        # Step 4. If the batching information is passed to the object at instantiation, then merge this into a
        #   dictionary. Error checking will be completed later.
        # Step 4.1. If all the expected command line parameters are set to the defaults assume that the a batch
        #   approach has been used. We do _not_ test for batch since this will be tested for later.
        input_arg = [data_worksheet, structure_worksheet, header_row, remove_columns, drop_empty_rows, template_file,
                     filter_rows, output_fp]

        self.batch_df: pd.DataFrame = pd.DataFrame()
        # Step 4.2. Since the default values for the input args are all 'None' or 0, if we remove these values from the
        #   list, if the list's length is greater than 0 then there is a chance that a single wash is required. We don't
        #   test for the input file path since this has already occurred.
        if len(remove_from_iterable(input_arg, None, 0)) > 0:

            # If the drop_empty_rows is None set it to True. This will save the user problems.
            if drop_empty_rows is None:
                drop_empty_rows = True

            # Step 4.3. Turn the command line arguments into a dict and store temporarily.
            t_batch_dict = {'data_worksheet': data_worksheet, 'structure_worksheet': structure_worksheet,
                            'header_row': header_row, 'remove_columns': remove_columns,
                            'drop_empty_rows': drop_empty_rows, 'template_file': template_file,
                            'filter_rows': filter_rows, 'output_fp': output_fp}
            self.batch_df.from_dict(t_batch_dict)

        # Step 5. If batch information passed as a worksheet clean and sort the batch data.
        self.batch_df = self.excel_to_dataframe(self._washing_basket, batch_worksheet, header_row=0, clean_header=True)

        # Step 6. Check the batch data.
        try:
            print(f'Check: Checking batch data.')
            print(f'---------------------------')
            self.check_batch_worksheet_data()
        except Exception as e:
            print(f'{e}')
        print(f'---------------------------')

        # Step 6. Convert the batch DataFrame to a dict and store.
        self._batch_dict = self.batch_df.to_dict('records')

        # Step 7. for every batch data in self._batch heck the corresponding structure and data details.
        for t_batch_row in self._batch:

            # Step 7.1. Open the structure worksheet as a DataFrame
            t_batch_row_structure = t_batch_row['structure_worksheet']
            print(f'Checking structure worksheet {t_batch_row_structure}')
            t_structure_df: pd.DataFrame = self.excel_to_dataframe(self._washing_basket,
                                                                   worksheet=t_batch_row_structure,
                                                                   clean_header=True)
            t_structure_headers = list(t_structure_df)
            try:
                for expected, actual in [(self.expected_structure_headers, t_structure_headers)]:
                    self.compare_lists(expected, actual)
            except Exception as e:
                print(e)

            # Step 7.2. Open the data worksheet as a DataFrame
            # We do this before checking the structure dict to allow the data headers to be used
            t_data_df: pd.DataFrame = self.excel_to_dataframe(self._washing_basket, worksheet=t_batch_row['data'],
                                                              header_row=header_row, clean_header=True,
                                                              drop_empty_rows=t_batch_dict['drop_empty_rows'])

            # Step 7.3. Collect the structure and data headers. Check that the expected structure headers are in the
            # structure worksheet.
            t_data_headers = list(t_data_df)

            if set(self.expected_structure_headers) <= set(t_structure_headers) is False:
                raise ValueError(f'The structure headers {self.expected_structure_headers} were expected. The following '
                                 f'headers were found: {t_structure_headers}')

            # Step 7.4. Convert the structure spreadsheet to a dict and check its data.
            t_structure_enum = enumerate(t_structure_df.to_dict('records'), 0)
            for t_structure_row_int, t_structure_dict in t_structure_enum:
                try:
                    print(f'\tRow {t_structure_row_int}: {t_structure_dict}')
                    self.check_structure_data(t_structure_dict, t_data_headers)
                except Exception as e:
                    print(f'{e}')
                else:
                    print(f'\t\tOk.')

        t_data_dict: dict = t_data_df.to_dict('records')

        # Step 4.6. Check file paths that

    @staticmethod
    def check_worksheets_basic(t_sheets_expected: list) -> bool:
        """

        :param t_sheets_expected:
        :return:
        """
        if len(t_sheets_expected) == 0:
            raise ValueError(f'Either the "data" and "structure" worksheets, or the "batch" worksheet must be '
                             f'provided.')

    def check_batch_worksheet_data(self):
        """
        Check the batch worksheet data is in the correct format. Data is checked as a DataFrame. The following checks are
        made.
        Check 1: Confirm expected batch headers exist in the batch worksheet.
        Check 2: Confirm Structure and data worksheets referenced in batch worksheet exist.
        Check 3: Confirm the template files exist and resolve the files.
        Check 4: Confirm an output filename has been provided.
        Check 5: Check if remove_column is None, set it to False.
        Check 6: Check if drop_empty_rows is None, set it to False.
        Check 7: Check if header_row is None, set it to 0.
        :return:
        """
        # Extract the data and structure worksheet names from the batch worksheet for error checking.
        t_batch_headers = list(self.batch_df)
        t_batch_data_worksheets_expected = self.batch_df['data_worksheet'].to_list()
        t_batch_structure_worksheets_expected = self.batch_df['structure_worksheet'].to_list()

        # Check 1.
        print(f'Check: Checking batch work sheet headers.')
        try:
            for expected, actual in [(self.expected_batch_headers, t_batch_headers)]:
                self.compare_lists(expected, actual)
        except Exception as e:
            print(e)
        print(f'\tBatch worksheets ok.')

        # Check 2.
        print(f'Check: Checking data and structure worksheets are correctly referenced.')
        try:
            for expected, actual in [(t_batch_structure_worksheets_expected, self._sheets_actual),
                                     (t_batch_data_worksheets_expected, self._sheets_actual)]:
                self.compare_lists(expected, actual)
        except Exception as e:
            print(e)
        print(f'\tData and structure worksheets present.')

        # Check 3.
        for row in self.batch_df.intertuples():
            print(f'\tChecking row {row.Index}')
            print(f'\t\tChecking file {row.template_file}.')
            try:
                print(f'Checking batch template files.')
                if row.template_file not in invalid:
                    self.batch_df.at[row.Index, 'template_file'] = resolve_file_path(row.template_file)
            except ValueError as v:
                print(f'{v}. The template file does not exist.')
            except Exception as e:
                print(f'{e}')

            # Check 4.
            print(f'\t\tCheck output filename.')
            if str(row.output_file) in invalid:
                raise ValueError(f'The name of the output file has not been provided.')

            # Check 5.
            if row.remove_columns is None:
                self.batch_df.at[row.Index, 'remove_columns'] = False

            # Check 6.
            if row.drop_empty_rows is None:
                self.batch_df.at[row.Index, 'drop_empty_rows'] = False

            # Check 7
            if row.header_row is None:
                self.batch_df.at[row.Index, 'header_row'] = 0

    def check_data_data(self):
        pass

    def check_structure_data(self, structure_data: dict, data_headers: list):
        """
        Cycle through the dictionary and run the following data checks. title_style is not checked
        :param structure_data: The structure worksheet data.
        :param data_headers: The data worksheet's headers. This is passed to allow the data headers to be checked prior
        to parsing the data dictionary.
        :return:
        """
        for each in structure_data:
            # Check 1.
            # Check the header details are correct by checking whether the self.expected_structure_headers are a subset of
            # structure_data_keys if True we can continue and ignore any additional headers that have been provided.
            structure_data_keys = each.keys()
            if set(self.expected_structure_headers) <= set(structure_data_keys) is False:
                raise ValueError(f'The provided batch headers {structure_data_keys} do not match the required headers '
                                 f'{self.expected_batch_headers}.')
            # Check 2.
            # Confirm that the section_type data is correct.
            if each['section_type'] not in structure_keys:
                raise ValueError(f'The section_type {each["section_type"]} is not correct.')

            # Check 3.
            # Confirm that formatting data has been provided and is present.
            if each['section_style'] is None:
                raise ValueError(f'Worksheet "section_style" has not been provided.')

            # Check 3.
            if each['section_type'].lower() is 'photo' and each['path'] not in invalid:
                if not Path(each['path']).exists():
                    raise ValueError(f'The provided path {each["path"]} does not exist.')
                each['path'] = Path(each['path']).resolve()

            # Check 4.
            # Check that the data section referenced in the structure worksheet exists in the data worksheet.
            if each['section_contains'] not in data_headers:
                raise ValueError(f'{each["section_contains"]} is not defined within the structure worksheet.')

            # Check 5.
            for element in ['section_break', 'page_break']:
                if str(each[element]).lower() in ['none', 'false']:
                    each[element] = False

        self._structure.append(structure_data)

    def excel_to_dataframe(self, io, worksheet: str, header_row: int = 0, remove_col: Iterable[str] = None,
                           clean_header: bool = False, drop_empty_rows: bool = False) -> data_frame:
        """
        Open and perform basic cell_data cleaning on a single excel work worksheet.
        :param io: The Excel file to be read.
        :param worksheet: The Excel spreadsheet worksheet's name.
        :param header_row: index of the header row in the spreadsheet. Defaults to 0, i.e. assumes the headers are at
        the top of the page.
        :param remove_col: remove the column headers contained in the passed list.
        :param clean_header: If True clean the column headers
        :param drop_empty_rows: If True remove empty rows.
        :return:
        """
        df = pd.read_excel(io, worksheet, header_row)
        if remove_col is not None:
            df = df.remove_columns(remove_col)
        if clean_header is not False:
            df = df.clean_names()
        if drop_empty_rows is True:
            df = df.dropna(thresh=2)
        return df

    def prepare_row_filters(self, filters: str) -> List[Tuple[str, int]]:
        """
        Split the passed string into a list of Tuples taking the form (column_name, filter_keyword)
        :param filters:
        :return:
        """
        # filters = str(filters)  # todo: Turned off for the moment. To confirm what the data is actually passed as.
        cleaned_filters = list()
        i = filters.splitlines()
        for each in i:
            column_header, column_keyword = each.split(':')
            column_keyword = column_keyword.split(',').strip()
            cleaned_filters.append((column_header, column_keyword))
        return cleaned_filters

    def compare_lists(self, expected_list, actual_list):
        if set(expected_list) <= set(actual_list) is False:
            raise ValueError(f'The provided headers {actual_list} do not match the required headers '
                             f'{expected_list}.')
        else:
            return True

    def single_load(self) -> SingleLoad:
        """
        Create a Laundry object.
        :return:
        """
        pass
