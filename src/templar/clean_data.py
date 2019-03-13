import pandas as pd
from typing import List, Dict, Tuple, NewType
import janitor


data_frame = NewType('data_frame', pd.DataFrame)


def clean_xlsx_table(
        file_path: str,
        sheet: str,
        head: int = 0,
        remove_col: List[str] = None,
        clean_hdrs: bool = False,
        drop_empty: bool = False
) -> data_frame:
    """
    Open and perform basic data cleaning on a single excel work sheet.
    :param file_path: path to the excel file
    :param sheet: excel spreadsheet name
    :param head: index of the header row in the spreadsheet. Defaults to 0.
    :param remove_col: list of columns to be dropped from the table.
    :param clean_hdrs: If True clean the column headers
    :param drop_empty: If True remove empty rows.
    :return:
    """
    df = pd.read_excel(pd.ExcelFile(file_path), sheet, head)
    if remove_col is not None:
        df = df.remove_columns(remove_col)
    if clean_hdrs is not False:
        df = df.clean_names()
    if drop_empty is True:
        df = df.dropna()
    return df


def extract_data(
        record: Dict,
        header: List[str],
        format_title: bool = True
) -> List[Tuple]:
    """
    Take a dictionary and split in to a list of tuples containing the 'keys' data defined in 'header' as the first tuple and the associated values as the second.
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
        if format_title is False:
            hdr.append(str(h))
        t = record.pop(each)
        data.append(t)
    return [tuple(hdr), tuple(data)]


def remove_underscore(text: str) -> str:
    return text.replace('_', ' ')

