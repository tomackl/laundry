from docx import Document
from docx.shared import Inches
import pandas as pd
import janitor
from pathlib import Path, PurePath

from typing import List, Iterable, Dict, Tuple, Any, NewType
import click


data_frame = NewType('data_frame', pd.DataFrame)


def clean_xlsx_table(file_path: str, sheet: str, head: int = 0,
                     rm_column: List[str] = None, clean_hdr: bool = False,
                     drop_empty: bool = False) -> data_frame:
    """
    Open and perform basic data cleaning on a single excel work sheet.
    :param file_path: path to the excel file
    :param sheet: excel spreadsheet name
    :param head: index of the header row in the spreadsheet. Defaults to 0.
    :param rm_column: remove the column headers contained in the passed list.
    :param clean_hdr: If True clean the column headers
    :param drop_empty: If True remove empty rows.
    :return:
    """
    # todo: add exception here to catch a XLRDError in the event of a misnamed work sheet.
    df = pd.read_excel(pd.ExcelFile(file_path), sheet, head)
    if rm_column is not None:
        df = df.remove_columns(rm_column)
    if clean_hdr is not False:
        df = df.clean_names()
    if drop_empty is True:
        df = df.dropna(thresh=2)
    return df


def extract_data(record: Dict, header: List[str], format_title: bool = True
                 ) -> List[Tuple]:
    """
    Take a dictionary and split in to a list of tuples containing the 'keys'
    data defined in 'header' as the first tuple and the associated values as the second.
    The function will return a list containing two equal length tuples.
    :param record: the dictionary containing the data.
    :param header: the keys that defined the key-values to be extracted.
    :param format_title: make the header string title case.
    :return: a list of tuples
    """
    hdr, data = [], []
    for each in header:
        h = remove_underscore(each)
        if format_title is True:
            hdr.append(str(h).title())
        else:
            hdr.append(str(h))
        t = record.pop(each)
        data.append(t)
    return [tuple(hdr), tuple(data)]


def remove_underscore(text: str) -> str:
    return text.replace('_', ' ')


def section_contains(sect_contains: Any) -> List[str]:
    """
    A function to split a string into a list. It will return a list.
    :param sect_contains:
    :return:
    """
    if '\n' in sect_contains:
        return list(str(sect_contains).splitlines())
    elif ',' in sect_contains:
        return list(str(sect_contains).split(','))
    else:
        _ = list()
        _.append(sect_contains)
        return _


def insert_paragraph(document: object, text: str, title: str = None,
                     section_style: str = None, title_style: str = None,
                     ):
    """
    This function defines the characteristics of a paragraph to be added to the document.
    :param document: the template_doc the paragraph will be added to.
    :param text: paragraph text
    :param title: title text
    :param section_style: paragraph style
    :param title_style: title style. Use this _or_ title_level.
    :return:
    """
    if text == '':
        document.add_paragraph(text)
        return
    text = text.splitlines()
    if (title is not None) and (str(title_style) != 'nan'):
        document.add_paragraph(remove_underscore(title), style=title_style)
    for each in text:
        document.add_paragraph(each, style=section_style)


def insert_table(document: object, cols: int, rows: int,
                 data: List[Iterable[str]], section_style: str = None,
                 autofit: bool = True):
    """
    The function takes data and uses it to create a table for the template_doc.
    The first row of data is assumed to be the table header.
    :param document: the docx file the table will be added to.
    :param rows: the number of required table rows.
    :param cols: the number of required table columns.
    :param data: The list data to be inserted into the table. The idx[0] is assumed to be the header.
    :param section_style: The style to be used for the table.
    :param autofit: autofit the table to the page width.
    :return:
    """
    # todo: add error checking for 'cols' and 'row'
    table: object = document.add_table(rows=rows, cols=cols, style=section_style)
    if autofit is True:
        table.autofit = True
    data = enumerate(data, 0)
    for i, cell_contents in data:
        insert_row(table.rows[i].cells, cell_contents)


def insert_row(row_cells, data: List[str]) -> Any:
    """
    Populate a table row. The cells are passed as a row and the contents added.
    :param row_cells:
    :param data:
    :param style: style is the text style to be applied to the individual rows.
    :return:
    """
    for i, text in enumerate(data):
        row_cells[i].text = str(text)
    return row_cells


def insert_photo(document: object, photo: str, width: int = 4):
    """
    Insert a photo located at path into document and set the photo to width.
    :param document: the docx file the photo will be added to.
    :param photo: file path to the
    :param width: width of the image in Inches
    :return:
    """
    photo_formats = ['', '.jpg', '.jpeg', '.png', '.tiff']
    # photo_path = Path(photo)
    # document.add_picture(str(photo_path), width=Inches(width))
    for ext in photo_formats:
        photo_path = Path(str(photo) + ext)
        if photo_path.exists():
            document.add_picture(str(photo_path), width=Inches(width))
            return True
    print('\nPhoto {} does not exist. Check file extension.'.format(photo))
    document.add_paragraph('PHOTO "{}" NOT FOUND\n'.format(str(photo).upper()))



def format_docx(rowdict: dict, structdict: dict, outputfile: object, file_path: str):
    """
    The function is passed a dict (data_dict) containing the data to be formatted
    (structure) based on the template (outputfile).
    :param rowdict: dictionary containing the text. It represents a single row from the spreadsheet.
    :param structdict: defines the output file's format structure.
    :param outputfile: The file which data will be inserted into.
    :param file_path: directory containing the spreadsheet
    :return:
    """
    file_path = file_path
    # todo: add error checking here.
    for element in structdict:
        element_sect_con = str(element['sectioncontains']).lower()
        if str(element['sectiontype']).lower() in ('heading', 'para', 'paragraph'):
            insert_paragraph(outputfile, str(rowdict[str(element_sect_con).lower()]),
                             title=element_sect_con.title(),
                             section_style=element['sectionstyle'],
                             title_style=element['titlestyle']
                             )

        elif str(element['sectiontype']).lower() == 'table':
            table = section_contains(element_sect_con)
            data = extract_data(rowdict, table)
            insert_table(outputfile, len(table), len(data),
                         data, section_style=element['sectionstyle']
                         )

        elif str(element['sectiontype']).lower() == 'photo':
            sect_contains = rowdict[element_sect_con]
            q = confirm_path([file_path, element['path']])
            if str(sect_contains).lower() not in ['no photo', 'none', 'nan', '-']:
                photo = section_contains(sect_contains)
                for each in photo:
                    loc = q.joinpath(each)
                    insert_photo(outputfile, loc, 4)
        else:
            print('Valid section header was not found.')

        if element['sectionbreak'] is True:
            insert_paragraph(outputfile, '')

        if element['pagebreak'] is True:
            outputfile.add_page_break()


def confirm_path(filepath: List[str]) -> Path:
    """
    Convert the contents of the passed list into a Path
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


"""
the following has been implemented:

* sectioncontains: list[str] -> containing the column titles
* sectiontype: str -> paragraph/table/heading/photo
* sectionstyle: str -> Word paragraph style
* titlestyle: str -> Word title style. This does not apply to tables.
* sectionbreak: bool -> is a break required after each section?
* pagebreak: bool -> is a page break required after the section?
    - => Headings should be a single column per paragraph
    - => Paragraphs should be a single column per paragraph
    - => Tables can be any number of columns (there will be practical limits)
    - => section_styles _can_ only be a single value
    - => title_styles _can_ only be a single value
    - => pagebreak is a True/False value
"""


@click.command()
@click.option('--data-worksheet', '-dw', 'data',
              default='Master List',
              help='Name of the worksheet containing the data to be converted into a word document. '
                   'The default is "Master List".'
              )
@click.option('--template', '-t', 'template',
              help='Name of the template file to be used used as the basis of the converted file.',
              type=click.Path(exists=True)
              )
@click.option('--structure-worksheet', '-sw', '-s', 'structure',
              default='_structure',
              help='Name of the worksheet containing the data to format the structure of the outfile document. '
                   'The default is _structure.'
              )
@click.option('--data-header', '-dh', 'data_head',
              default=0,
              type=int,
              help="The row number of the data worksheet's row containing the column headers. The default is 0."
              )
@click.argument('input_file',
                type=click.Path(exists=True)
                )
@click.argument('output-file')
def cli(input_file, output_file, data, structure, template, data_head):
    """
    This is the command line interface (cli) for the Laundry app. For details regarding the operation of the app type
    `laundry --help`.

    The relative path for each file should be provided with each of the options if non-default file names are provided.

    NOTE: If output files are intended to be saved in a separate directory, that directory *must* exist otherwise the
    output file will not save.

    IMPORTANT: Laundry will overwrite, without prompting, any files with the same name in the directory where output
    files are saved.
    """
    file_input = Path(input_file)
    file_output = output_file
    wkst_data = data
    wkst_struct = structure
    template = template
    wash(file_input, file_output, wkst_data, wkst_struct, template, data_head)


def wash(file_input, file_output, wkst_data, wkst_struct, template, data_head):
    """
    This function acts as a common calling point for the module to allow the module to be run from the command line
    interface (cli) or from another script.
    :param file_input: the .xls file containing the data to be converted.
    :param file_output: name of the output file.
    :param wkst_data: name of the .xls worksheet containing the data to be processed
    :param wkst_struct: name of the .xls worksheet detailing how the data shall be processed
    :param template: the .docx file to be used as the template
    :param data_head: the number of the data worksheet's row containing the column headers.
    :return:
    """
    # todo: add exception to ensure that the `template` file actually exists.
    #  `docx.opc.exceptions.PackageNotFoundError' is raised if the file does not exist.
    file_template = Document(template)
    # todo: add exceptions to catch files that are missing file extensions.
    path_input_f = file_input.parents[0]

    structure_file = clean_xlsx_table(file_input,
                                      sheet=wkst_struct,
                                      head=0,
                                      clean_hdr=True,
                                      drop_empty=False
                                      )
    structure_dict = structure_file.to_dict('records')

    # todo: allow for `remove_columns` to be defined somewhere
    remove_columns = []
    data_file = clean_xlsx_table(file_input,
                                 sheet=wkst_data,
                                 head=data_head,
                                 rm_column=remove_columns,
                                 clean_hdr=True,
                                 drop_empty=True
                                 )
    data_dict = data_file.to_dict('records')

    with click.progressbar(iterable=data_dict,
                           label='Conversion progress:',
                           fill_char='|',
                           empty_char='_'
                           ) as data_dictionary:
        for row in data_dictionary:
            format_docx(row, structure_dict, file_template, file_path=str(path_input_f))
            # format_docx(row, structure_dict, file_template, file_path=str(path_input_f))

    file_template.save(file_output)
