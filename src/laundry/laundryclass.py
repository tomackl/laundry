"""Main class for laundry. This is intended to replace the original laundry script."""

from laundry.constants import data_frame, invalid, photo_formats
from typing import Dict, List, Iterable, Tuple, NamedTuple
from docx import Document
from docx.shared import Inches
from pathlib import Path, PurePath
import pandas as pd

old_structure_keys = ('sectiontype', 'sectioncontains', 'sectionstyle', 'titlestyle', 'sectionbreak', 'pagebreak',
                      'path')
# Define headers for the batch and structure worksheets. These are fixed.
expected_batch_headers = ['data_worksheet', 'structure_worksheet', 'header_row', 'remove_columns', 'drop_empty_rows',
                          'template_file', 'filter_rows', 'output_file']
expected_structure_headers = ['section_type', 'section_contains', 'section_style', 'title_style', 'section_break',
                              'page_break', 'path']


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

    def __init__(self, structure_data: pd.DataFrame, data_data: pd.DataFrame, file_template: Path,
                 file_output_path: Path):
                 # file_output_path: (Path, str), spreadsheet_fp: (Path, str)):
        """
        # The method signature is based on the laundry.single_load() function. This calls self.format_docx()
        :param structure_data: A dictionary that defines the structure of the documentation.
        :param data_data: A dictionary that contains the cell_data to be formatted.
        :param file_template: The Word .docx file that contains the formatting styles to be used.
        :param spreadsheet_fp: The path to the directory containing the spreadsheet.
        :param file_output_path: The path to the output file location.
        """
        self._structure: pd.DataFrame = structure_data
        # self._structure: dict = structure_data
        self._data: pd.DataFrame = data_data
        # self._data: dict = data_data
        self._file_template: Path = Path(file_template)  # todo: This should probably be a Path()
        # self._spreadsheet_fp: (Path, str) = spreadsheet_fp
        self._file_output: Path = Path(file_output_path)
        self._output_docx = Document()
        self._row_data: List[Dict] = list()
        # self.split_into_rows()
        self.start_wash()
        # self.format_docx()
    # def split_into_rows(self):
    #     """
    #     Split self._data and pass each row to self.format_docx.
    #     """
    #     for row in self._row_data.itertuples():
    #         self.format_docx(row)
    #
    #     # self._row_data = list()
    #     # for each in self._data:
    #     #     self._row_data.append({each: self._data[each]})

    def start_wash(self):
        """
        Start formatting the output document.
        """
        for row in self._data.intertuples():
        # for each in self._row_data:
            self.format_docx(row)
            # self.format_docx(each)

    # def format_docx(self):
    def format_docx(self, row: NamedTuple):
    # def format_docx(self, row: dict):
        """
        This is factory method that calls the appropriate the information contained within document structure.
        :param row: dictionary containing the data_str to be formatted. This is a single row from the spreadsheet.
        :return:
        """

        row = row._asdict()
        for element in self._structure.itertuples():
        # for element in self._structure:
            ele_sect_contains: str = str(element.section_contains).lower()
            # ele_sect_contains: str = str(element['sectioncontains']).lower()
            ele_sect_style: str = element.section_style
            # ele_sect_style: str = element['sectionstyle']
            ele_sect_type: str = str(element.section_type).lower()
            # ele_sect_type: str = str(element['sectiontype']).lower()
            ele_title_style: str = element.title_style
            # ele_title_style: str = element['titlestyle']
            ele_sect_break: bool = element.section_break
            # ele_sect_break: bool = element['sectionbreak']
            ele_page_break: bool = element.page_break
            # ele_page_break: bool = element['pagebreak']
            ele_row_contains = row[ele_sect_contains]

            if ele_sect_type in ['heading', 'para', 'paragraph']:
                self.insert_paragraph(str(ele_row_contains).lower(), title=ele_sect_contains.title(),
                                      section_style=ele_sect_style, title_style=ele_title_style)

            elif ele_sect_type == 'table':
                table_col_hdr = split_str(ele_sect_contains)
                sorted_row = sort_table_data(row, table_col_hdr)
                self.insert_table(len(table_col_hdr), len(sorted_row), sorted_row, section_style=ele_sect_style)

            elif ele_sect_type == 'photo':
                q = confirm_directory_path([self._spreadsheet_fp, element['path']])
                # todo: allow for photo names having to be split
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
            print(f'<-- Checking batch worksheet data -->')
            self.check_batch_worksheet_data()
        except Exception as e:
            print(f'{e}')
        print(f'--> Batch data checked <--')

        # Step 6. Convert the batch DataFrame to a dict and store.
        self._batch_dict = self.batch_df.to_dict('records')

        # Step 7.
        # Every row of the the batch DataFrame contains information regarding an output file. For each row in the
        # DataFrame produce the associated output file.
        for t_batch_row in self.batch_df.itertuples():

            t_structure_worksheet = t_batch_row.structure_worksheet
            t_data_worksheet = t_batch_row.data_worksheet
            t_filter_data_columns = t_batch_row.filter_rows
            self.t_structure_df = self.excel_to_dataframe(self._washing_basket, t_structure_worksheet, header_row=0,
                                                          clean_header=True, drop_empty_rows=False)
            self.t_structure_photo_path: Dict[str, Path] = {}
            self.t_data_df = self.excel_to_dataframe(self._washing_basket, t_data_worksheet,
                                                     header_row=t_batch_row.header_row,
                                                     remove_col=t_filter_data_columns, clean_header=True,
                                                     drop_empty_rows=True)

            # Step 8.
            # Check the structure data.
            try:
                print(f'<-- Checking structure worksheet data -->')
                self.check_structure_worksheet_data()
            except Exception as e:
                print(f'{e}')
            print(f'--> Structure data checked <--')

            # Step 9.
            # Check the data worksheet data.
            try:
                print(f'<-- Checking data worksheet data -->')
                self.check_data_worksheet_data()
            except Exception as e:
                print(f'{e}')
            print(f'--> Data worksheet data checked <--')

            SingleLoad(self.t_structure_df, self.t_data_df,
            # SingleLoad(self.t_structure_df.to_dict('records'), self.t_data_df.to_dict('records'),
                       t_batch_row.template_file, t_batch_row.output_file)
                       # t_batch_row.template_file, t_batch_row.output_file)
            del self.t_structure_photo_path


    @staticmethod
    def check_worksheets_basic(t_sheets_expected: list) -> bool:
        """

        :param t_sheets_expected:
        :return:
        """
        if len(t_sheets_expected) == 0:
            raise ValueError(f'Either the "data" and "structure" worksheets, or the "batch" worksheet must be '
                             f'provided.')
        else:
            return True

    def check_batch_worksheet_data(self):
        """
        Check the batch worksheet data is in the correct format. Data is checked as a DataFrame. The following checks
        are made.
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
        t_batch_data_worksheets_expected = self.batch_df[['data_worksheet']].to_list()
        t_batch_structure_worksheets_expected = self.batch_df[['structure_worksheet']].to_list()

        # Check 1.
        print(f'Check: Checking batch work sheet headers.')
        try:
            for expected, actual in [(expected_batch_headers, t_batch_headers)]:
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
            if row.remove_columns.notna():
                self.batch_df.at[row.Index, 'remove_columns'] = self.prepare_row_filters(row.remove_columns)

            # Check 6.
            if row.drop_empty_rows is None:
                self.batch_df.at[row.Index, 'drop_empty_rows'] = False

            # Check 7
            if row.header_row is None:
                self.batch_df.at[row.Index, 'header_row'] = 0

    def check_data_worksheet_data(self):
        """
        Check 1. Check that the photos exist in the directory. The check assumes that the first file name with the same
        name is the correct file if no filename has been provided in the worksheet. The check assumes that all photo
        names are unique regardless of photo directory.
        :return:
        """
        # Check 1. check the photo paths.
        t_photos_found: Dict[str, Path] = {}
        # Assuming that that there may be more than one directory containing photos for the worksheet loop through the
        # each folder.
        columns = self.t_structure_photo_path.keys()
        for col in columns:
            # Grab the photos stored in the folders and store their paths in dictionary. Store them in the dictionary
            # with their name minus the file extension as the key.
            for file_ext in photo_formats:
                for file in self.t_structure_photo_path[col].glob('*' + file_ext):
                    t_photos_found[file.name] = file

        # For each of the columns containing photos loop through the self.t_data_df and replace the file name with the
        # path.
            for row in pd.DataFrame(self.t_data_df[col]).itertuples():
                t_row = row._asdict()
                t_row_photo = []
                if str(t_row[col]).lower() not in ['no photo', 'none', 'nan', '-']:
                    for t in split_str(t_row[col]):
                        try:
                            t_row_photo.append(self.check_photo_paths(t, t_photos_found))
                        except Exception as e:
                            print(e)
                self.t_data_df[t_row['Index'], col] = t_row_photo

    @staticmethod
    def check_photo_paths(expected_photo: (Path, str), actual_photos: dict) -> Path:
        """
        Check 1. Check that the photo could be an image file.
        Check 2. If the expected_photo does not have a file extension then loop through photo_formats and see if the
        file does exist.
        Check 3. If the expected_photo does have a file extension then check that the file does exist.
        The method will raise exceptions if the file is not found
        :param expected_photo:
        :param actual_photos:
        :return:
        """
        # Check 1
        photo = expected_photo.strip()
        if Path(photo).suffix is not '' and Path(photo).suffix not in photo_formats:
            raise ValueError(f'The data worksheet photo {photo} is not been specified as a photo. Ensure that the'
                             f' file format is one of the following formats {photo_formats}.')
        # Check 2
        elif Path(photo).suffix is '':
            for ext in photo_formats:
                try:
                    t_photo_name = str(photo + ext)
                    if actual_photos[t_photo_name]:
                        return Path(actual_photos[t_photo_name]).resolve(strict=True)
                except KeyError:
                    pass
            raise ValueError(f'The photo {photo} does not exist in the specified directory.')
        # Check 3
        elif Path(photo).suffix in photo_formats:
            try:
                if actual_photos[photo]:
                    return Path(actual_photos[photo]).resolve(strict=True)
            except KeyError as k:
                print(f'{k}. The photo {photo} does not exist in the directory.')

        raise ValueError(f'The photo {photo} does not appear to exist in the directory.')

    def check_structure_worksheet_data(self):
        """
        Check 1: Confirm expected batch headers exist in the batch worksheet.
        Check 2: Confirm the expected section types exist.
        Check 3: Confirm the section_contains values exist in the structure document.
        Check 4: Check the photo file paths and resolve
        Check 5: Check if section_break is None, set it to False.
        Check 6: Check if page_break is None, set it to False.
        :return:
        """
        t_structure_headers = list(self.t_structure_df)
        t_structure_section_types = self.t_structure_df['section_types']
        t_structure_section_contains = self.t_structure_df['section_contains']
        t_data_section_types = list(self.t_data_df)

        # Check 1
        print(f'Check: Checking structure work sheet headers.')
        try:
            for expected, actual in [(expected_structure_headers, t_structure_headers)]:
                self.compare_lists(expected, actual)
        except Exception as e:
            print(e)
        print(f'\tStructure worksheets ok.')

        # Check 2
        print(f'Check: Checking structure worksheet section_types are correct.')
        try:
            for expected, actual in [(expected_structure_headers, t_structure_section_types)]:
                self.compare_lists(expected, actual)
        except Exception as e:
            print(e)
        print(f'\tsection_types are correct.')

        # Check 3
        print(f'Check: Checking structure worksheet section_contain details are correct.')
        try:
            for expected, actual in [(t_data_section_types, t_structure_section_contains)]:
                self.compare_lists(expected, actual)
        except Exception as e:
            print(e)
        print(f'\tsection_contains details ok.')

        for row in self.t_structure_df.itertuples():

            # Check 4
            print(f'\tChecking row {row.Index}')
            if str(row.section_type).lower() is 'photo':
                try:
                    print(f'\t\tChecking file {row.path}.')
                    t_photo_path = resolve_file_path(row.path)
                    if t_photo_path.is_dir():
                        self.batch_df.at[row.Index, 'path'] = t_photo_path
                        # Add the photo path to the dictionary with the section_contains as the key
                        self.t_structure_photo_path[row.section_contains] = t_photo_path
                except ValueError as v:
                    print(f'{v}. The path to the photos directory does not exist.')
                except Exception as e:
                    print(f'{e}')

            # Check 5.
            if row.section_break is None:
                self.batch_df.at[row.Index, 'remove_columns'] = False

            # Check 6.
            if row.page_break is None:
                self.batch_df.at[row.Index, 'drop_empty_rows'] = False

    @staticmethod
    def excel_to_dataframe(io, worksheet: str, header_row: int = 0, remove_col: Iterable[str] = None,
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

    @staticmethod
    def prepare_row_filters(filters: str) -> List[Tuple[str, int]]:
        """
        Split the passed string into a list of Tuples taking the form (column_name, filter_keyword)
        :param filters:
        :return:
        """
        cleaned_filters = list()
        i = filters.splitlines()
        for each in i:
            column_header, column_keyword = each.split(':')
            column_keyword = column_keyword.split(',').strip()
            cleaned_filters.append((column_header, column_keyword))
        return cleaned_filters

    @staticmethod
    def compare_lists(expected_list, actual_list):
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
