"""Main class for laundry. This is intended to replace the original laundry script."""

from laundry.constants import data_frame, invalid, photo_formats
from typing import Dict, List, Iterable, Tuple, NamedTuple, Any
from docx import Document
from docx.shared import Inches
from pathlib import Path
import janitor
import pandas as pd

# Define headers for the batch and structure worksheets. These are fixed.
expected_batch_headers = ['data_worksheet', 'structure_worksheet', 'header_row', 'drop_empty_rows', 'template_file',
                          'filter_rows', 'output_file']
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
    Take a dictionary and split in to a list of tuples containing the 'keys' cell_data defined in 'header' as the
    first tuple and the associated values as the second. The function will return a list containing two equal length
    tuples. The function assumes that the cell_data contained in record and header is in the correct order.
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


def resolve_file_path(path: (Path, str)) -> (Path, Exception):
    """
    Resolve the passed filepath returning as a Path() or an exception.
    :return:
    """
    q = str(path).replace('\\', '/')
    try:
        p = Path(q).resolve(strict=True)
        return p
    except Exception as e:
        print(f'{e}')


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


def strip_whitespace(wht_spc: List) -> List:
    i = []
    for each in wht_spc:
        i.append(each.strip())
    return i


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
        """
        # The method signature is based on the laundry.single_load() function. This calls self.format_docx()
        :param structure_data: A dictionary that defines the structure of the documentation.
        :param data_data: A dictionary that contains the cell_data to be formatted.
        :param file_template: The Word .docx file that contains the formatting styles to be used.
        :param file_output_path: The path to the output file location.
        """
        self._structure: pd.DataFrame = structure_data
        self._data: pd.DataFrame = data_data
        self._file_template: Document() = Document(file_template)
        self._file_output: Path = Path(file_output_path)
        self._row_data: List[Dict] = list()
        self.start_wash()
        self.issue_document()

    def start_wash(self):
        """
        Start formatting the output document.
        """
        for row in self._data.itertuples():
            self.format_docx(row)

    def format_docx(self, row: NamedTuple):
        """
        This is factory method that calls the appropriate the information contained within document structure.
        :param row: dictionary containing the data_str to be formatted. This is a single row from the spreadsheet.
        :return:
        """

        row = row._asdict()
        for structure_element in self._structure.itertuples():
            sect_contains_element: str = str(structure_element.section_contains).lower()
            sect_style_element: str = str(structure_element.section_style)
            sect_type_element: str = str(structure_element.section_type).lower()
            title_style_element: str = str(structure_element.title_style)
            sect_break_element: bool = structure_element.section_break
            page_break_element: bool = structure_element.page_break

            if sect_type_element in ['heading', 'para', 'paragraph']:
                self.insert_paragraph(str(row[sect_contains_element]).lower(), title=sect_contains_element.title(),
                                      section_style=sect_style_element, title_style=title_style_element)

            elif sect_type_element == 'table':
                table_col_hdr = split_str(sect_contains_element)
                sorted_row = sort_table_data(row, table_col_hdr)
                self.insert_table(len(table_col_hdr), len(sorted_row), sorted_row, section_style=sect_style_element)

            elif sect_type_element == 'photo':
                if row[sect_contains_element] is not 'None':
                    for each in row[sect_contains_element]:
                        self.insert_photo(each, 4)
            else:
                print('Valid section header was not found.')

            if sect_break_element is True:
                self.insert_paragraph('')

            if page_break_element is True:
                self._file_template.add_page_break()

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
        split_text = text.splitlines()
        if len(split_text) == 0:
            self._file_template.add_paragraph(split_text)
        else:
            if (title is not None) and (str(title_style) != 'nan'):
                self._file_template.add_paragraph(remove_underscore(title), style=title_style)
            for each in split_text:
                self._file_template.add_paragraph(each, style=section_style)

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
        table = self._file_template.add_table(rows=rows, cols=cols, style=section_style)
        table.autofit = autofit_table
        cell_data = enumerate(data, 0)
        for i, cell_contents in cell_data:
            for j, text in enumerate(cell_contents):
                table.rows[i].cells[j].text = str(text)

    def insert_photo(self, photo: Path, width: int = 4):
        """
        Insert a photo located at path into document and set the photo width.
        :param photo: file path to the
        :param width: width of the image in Inches
        :return:
        """
        self._file_template.add_picture(str(photo), width=Inches(width))

    def issue_document(self):
        """
        Output the file.
        :return:
        """
        self._file_template.save(self._file_output)


class Laundry:

    def __init__(self, input_fp: Path, data_worksheet: str = None, structure_worksheet: str = None,
                 batch_worksheet: str = None, header_row: int = 0, drop_empty_rows: bool = None,
                 template_file: str = None, filter_rows: str = None, output_file: (Path, str) = None):
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
        :param drop_empty_rows: An explicit tag to drop empty rows from the worksheet if they contain two or more empty
        cells. If this is left as None it will be automatically set to True for the data worksheet.
        :param template_file: 
        :param filter_rows: 
        :param output_file:
        """
        # Step 1: Basic data checking.
        t_sheets_expected = remove_from_iterable([data_worksheet, structure_worksheet, batch_worksheet], None)
        print(f'Check: Worksheets are present.')
        if len(t_sheets_expected) == 0:
            raise ValueError(f'Either the "data" and "structure" worksheets, or the "batch" worksheet must be '
                             f'provided.')
        print(f'\tWorksheets present: {t_sheets_expected}.')
        print('\n')
        # Step 2: Confirm the input file exists.
        self._input_fp: (Path, str) = ''
        try:
            print(f'Check: Resolving spreadsheet filepath.')
            self._input_fp = resolve_file_path(input_fp)
            print(f'\t{self._input_fp}')
            print('\n')
        except Exception as e:
            print(f'\t{e}: File {input_fp} does not exist.')

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
            print(f'Check: Worksheets exist in spreadsheet')
            self.compare_lists(t_sheets_expected, self._sheets_actual)
            print(f'\tWorksheets present: {self._sheets_actual}.')
            print('\n')
        except Exception as e:
            print(f'{e}')

        # Step 4. If the batching information is passed to the object at instantiation, then merge this into a
        #   dictionary. Error checking will be completed later.
        # Step 4.1. If all the expected command line parameters are set to the defaults assume that the a batch
        #   approach has been used. We do _not_ test for batch since this will be tested for later.
        input_arg = [data_worksheet, structure_worksheet, header_row, drop_empty_rows, template_file, filter_rows,
                     output_file]

        self.batch_df: pd.DataFrame = pd.DataFrame(columns=['data_worksheet', 'structure_worksheet', 'header_row',
                                                            'drop_empty_rows', 'template_file', 'output_file'])
        # Step 4.2. Since the default values for the input args are all 'None' or 0, if we remove these values from the
        #   list, if the list's length is greater than 0 then there is a chance that a single wash is required. We don't
        #   test for the input file path since this has already occurred.
        if len(remove_from_iterable(input_arg, None, 0)) > 0:

            # If the drop_empty_rows is None set it to True. This will save the user problems.
            if drop_empty_rows is None:
                drop_empty_rows = True

            # Step 4.3. Turn the command line arguments into a dict and store temporarily.
            t_batch_dict = {'data_worksheet': [data_worksheet], 'structure_worksheet': [structure_worksheet],
                            'header_row': [header_row], 'drop_empty_rows': [drop_empty_rows],
                            'template_file': [template_file], 'filter_rows': [filter_rows],
                            'output_file': [output_file]}
            self.batch_df = pd.DataFrame.from_dict(data=t_batch_dict)
        # Step 5. If batch information passed as a worksheet clean and sort the batch data.
        else:
            self.batch_df = self.excel_to_dataframe(self._washing_basket, batch_worksheet, header_row=0,
                                                    clean_header=True)

        # Step 6. Check the batch data.
        try:
            print('\n')
            print(f'{self.batch_df}')
            print('\n')
            print(f'Check: Batch worksheet data.')
            self.check_batch_worksheet_data()
            print(f'\tBatch data checked')
            print('\n')
        except Exception as e:
            print(f'{e}')

        # Step 6. Convert the batch DataFrame to a dict and store.
        self._batch_dict = self.batch_df.to_dict('records')

        # Step 7.
        # Every row of the the batch DataFrame contains information regarding an output file. For each row in the
        # DataFrame produce the associated output file.
        for t_batch_row in self.batch_df.itertuples():
            t_structure_worksheet = t_batch_row.structure_worksheet
            t_data_worksheet = t_batch_row.data_worksheet
            # t_filter_data_columns = t_batch_row.filter_rows
            self.t_structure_df = self.excel_to_dataframe(self._washing_basket, t_structure_worksheet, header_row=0,
                                                          clean_header=True, drop_empty_rows=False)

            self.t_structure_photo_path: Dict[str, Path] = {}
            self.t_data_df = self.excel_to_dataframe(self._washing_basket, t_data_worksheet,
                                                     header_row=t_batch_row.header_row,
                                                     clean_header=True, drop_empty_rows=True)

            # Filter the data DataFrame using the filters passed.
            if str(t_batch_row.filter_rows).lower() not in invalid and t_batch_row.filter_rows is not None:
                for row_filter in t_batch_row.filter_rows:
                    self.t_data_df = self.t_data_df.loc[self.t_data_df[row_filter[0]].isin(row_filter[1])]

            # Step 8.
            # Check the structure data.
            try:
                print('\n')
                print(f'{self.t_structure_df}')
                print('\n')
                print(f'Check: Structure worksheet data')
                self.check_structure_worksheet_data()
                print(f'\tOK')
            except Exception as e:
                print(f'{e}')

            # Step 9.
            # Check the data worksheet data.
            try:
                print('\n')
                print(f'{self.t_data_df}')
                print('\n')
                print(f'Check: Data worksheet data')
                self.check_data_worksheet_data()
            except Exception as e:
                print(f'{e}')

            SingleLoad(self.t_structure_df, self.t_data_df,
                       t_batch_row.template_file, t_batch_row.output_file)

            del self.t_structure_photo_path

    def check_batch_worksheet_data(self):
        """
        Check the batch worksheet data is in the correct format. Data is checked as a DataFrame. The following checks
        are made.
        Check 1: Confirm expected batch headers exist in the batch worksheet.
        Check 2: Confirm Structure and data worksheets referenced in batch worksheet exist.
        Check 3: Confirm the template files exist and resolve the files.
        Check 4: Confirm an output filename has been provided.
        Check 5: Check if filter_rows
        Check 6: Check if drop_empty_rows is None, set it to False.
        Check 7: Check if header_row is None, set it to 0.
        :return:
        """
        # Extract the data and structure worksheet names from the batch worksheet for error checking.
        t_batch_headers = list(self.batch_df)
        t_batch_data_worksheets_expected = list(self.batch_df.loc[:, 'data_worksheet'])
        t_batch_structure_worksheets_expected = list(self.batch_df.loc[:, 'structure_worksheet'])

        # Check 1.
        print(f'\tCheck: Batch work sheet headers.')
        try:
            for expected, actual in [(expected_batch_headers, t_batch_headers)]:
                self.compare_lists(expected, actual)
            print(f'\t\tOK.')
        except Exception as e:
            print(e)

        # Check 2.
        print(f'\tCheck: Data & structure worksheets referenced correctly.')
        try:
            for expected, actual in [(t_batch_structure_worksheets_expected, self._sheets_actual),
                                     (t_batch_data_worksheets_expected, self._sheets_actual)]:
                self.compare_lists(expected, actual)
            print(f'\t\tOK')
        except Exception as e:
            print(e)

        # Check 3.
        for row in self.batch_df.itertuples():
            print(f'\tRow {row.Index}: Check file {row.template_file}.')
            try:
                if row.template_file not in invalid:
                    fp_template = resolve_file_path(row.template_file)
                    self.batch_df.at[row.Index, 'template_file'] = fp_template
                    print(f"\t\t{self.batch_df.at[row.Index, 'template_file']}")
            except ValueError as v:
                print(f'{v}. The template file does not exist.')
            except Exception as e:
                print(f'{e}')

            # Check 4.
            print(f'\tCheck output filename {row.output_file}.')
            if str(row.output_file) in invalid:
                raise ValueError(f'The name of the output file has not been provided.')
            try:
                fp_name = str(Path(row.output_file).name)
                fp_output = Path(resolve_file_path(Path(row.output_file).parent)).joinpath(fp_name)
                self.batch_df.at[row.Index, 'output_file'] = fp_output
                print(f"\t\t{self.batch_df.at[row.Index, 'output_file']}")
            except FileNotFoundError as f:
                print(f'{f}. File {row.output_file} could not be resolved.')
            except Exception as e:
                print(f'{e}')
            print(f'\t\tOK')

            # Check 5.
            if str(row.filter_rows).lower() not in invalid:
                self.batch_df.at[row.Index, 'filter_rows'] = self.prepare_row_filters(row.filter_rows)

            # Check 6.
            if row.drop_empty_columns is None:
                self.batch_df.at[row.Index, 'drop_empty_columns'] = False

            # Check 7
            if row.header_row is None:
                self.batch_df.at[row.Index, 'header_row'] = 0

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
        t_structure_section_types = list(self.t_structure_df.loc[:, 'section_type'])
        t_structure_section_contains = list(self.t_structure_df.loc[:, 'section_contains'])
        t_data_section_types = list(self.t_data_df)

        # Check 1
        print(f'\tCheck: Checking structure work sheet headers.')
        try:
            for expected, actual in [(expected_structure_headers, t_structure_headers)]:
                self.compare_lists(expected, actual)
        except Exception as e:
            print(e)
        print(f'\t\tOK')

        # Check 2
        print(f'\tCheck: Checking structure worksheet section_types are correct.')
        try:
            for expected, actual in [(expected_structure_headers, t_structure_section_types)]:
                self.compare_lists(expected, actual)
        except Exception as e:
            print(e)
        print(f'\t\tOK')

        # Check 3
        print(f'\tCheck: Checking structure worksheet section_contain details are correct.')
        try:
            for expected, actual in [(t_data_section_types, t_structure_section_contains)]:
                self.compare_lists(expected, actual)
        except Exception as e:
            print(e)
        print(f'\t\tOK')

        for row in self.t_structure_df.itertuples():
            # Check 4
            if str(row.section_type).lower() == 'photo':
                root = self._input_fp.parent
                try:
                    print(f'\tRow {row.Index}: Check file {row.path}')
                    t_photo_path = resolve_file_path(root.joinpath(row.path))
                    if t_photo_path.is_dir():
                        self.batch_df.at[row.Index, 'path'] = t_photo_path
                        # Add the photo path to the dictionary with the section_contains as the key
                        self.t_structure_photo_path[row.section_contains] = t_photo_path
                    print(f'\t\t{t_photo_path}')

                except ValueError as v:
                    print(f'{v}. The path to the photos directory does not exist.')
                except Exception as e:
                    print(f'{e}')

            # Check 5.
            if row.section_break is None:
                self.batch_df.at[row.Index, 'section_break'] = False

            # Check 6.
            if row.page_break is None:
                self.batch_df.at[row.Index, 'page_break'] = False

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
        print(f'\tPhotos are located in the following data worksheet columns:')
        print(f'\t\t{columns}')

        # Grab the photos stored in the folders and store their paths in dictionary. Store them in the dictionary
        # with their name minus the file extension as the key.
        for col in columns:
            for file_ext in photo_formats:
                for file in self.t_structure_photo_path[col].glob('*' + file_ext):
                    t_photos_found[file.name] = file
            # For each of the columns containing photos loop through the self.t_data_df and replace the file name with
            # the path.
            t_df_rows = self.t_data_df.itertuples()
            for row in t_df_rows:
                # The row below must remain in this position
                col = str(col).lower()
                t_row = row._asdict()
                t_row_photo = []
                print(f'\tRow {row.Index}:')
                if str(t_row[col]).lower() not in ['no photo', 'none', 'nan', '-']:
                    for t in split_str(t_row[col]):
                        try:
                            t_photo = self.check_photo_paths(t, t_photos_found)
                            t_row_photo.append(t_photo)
                        except Exception as e:
                            print(e)
                    self.t_data_df.at[t_row['Index'], col] = t_row_photo
                    print(f'\t\t{t_row_photo}')
                elif str(t_row[col]).lower() in ['no photo', 'none', 'nan', '-']:
                    self.t_data_df.at[t_row['Index'], col] = 'None'

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

    @staticmethod
    def excel_to_dataframe(io, worksheet: str, header_row: int = 0, clean_header: bool = False,
                           drop_empty_rows: bool = False) -> data_frame:
        """
        Open and perform basic cell_data cleaning on a single excel work worksheet.
        :param io: The Excel file to be read.
        :param worksheet: The Excel spreadsheet worksheet's name.
        :param header_row: index of the header row in the spreadsheet. Defaults to 0, i.e. assumes the headers are at
        the top of the page.
        :param clean_header: If True clean the column headers
        :param drop_empty_rows: If True remove empty rows.
        :return:
        """
        df = pd.read_excel(io, worksheet, header_row)
        if clean_header is not False:
            try:
                df = df.clean_names()
            except KeyError as k:
                print(f'{k}')
        if drop_empty_rows is True:
            try:
                df = df.dropna(thresh=2)
            except KeyError as k:
                print(f'{k}')
        return df

    @staticmethod
    def prepare_row_filters(filters: str) -> List[Tuple[Any, list]]:
        """
        Split the passed string into a list of Tuples taking the form (column_name, filter_keyword)
        :param filters:
        :return:
        """
        filtered_list = []
        i = []
        if '\n' in filters:
            i = str(filters).splitlines()
        else:
            i.append(filters)
        for each in i:
            col, b = each.split(':')
            col_kw = strip_whitespace(b.split(','))
            filtered_list.append((col, col_kw))
        return filtered_list

    @staticmethod
    def compare_lists(expected_list, actual_list):
        if set(expected_list) <= set(actual_list) is False:
            raise ValueError(f'The provided headers {actual_list} do not match the required headers '
                             f'{expected_list}.')
        else:
            return True
