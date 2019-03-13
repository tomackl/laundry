from typing import List, Iterable, Any


def insert_table(document: object,
                 cols: int,
                 rows: int,
                 data: List[Iterable[str]],
                 tbl_style: str = None,
                 autofit: bool = True) -> Any:
    """
    The function takes data related to a table and uses it to create a table for the document.
    The first row of data is assumed to be the table header.
    :param document: the document the table will be added to.
    :param rows: the number of required table rows.
    :param cols: the number of required table columns.
    :param data: The list data to be inserted into the table. The idx[0] is assumed to be the header.
    :param tbl_style: The style to be used for the table.
    :return:
    """
    table: object = document.add_table(rows=rows, cols=cols, style=tbl_style)
    if autofit is not None:
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


def insert_paragraph(document: object,
                     text: str,
                     title: str = None,
                     para_style: str = None,
                     title_style: str = None,
                     title_level: int = 0) -> Any:
    """
    :param document: the document the paragraph will be added to.
    :param text: paragraph text
    :param title: title text
    :param para_style: paragraph style
    :param title_style: title style. Use this _or_ title_level.
    :param title_level: title indent level. Use this _or_ title_style.
    :return:
    """
    if (title is not None) and (title_style is not None):
        document.add_paragraph(title, style=title_style)
    elif (title is not None) and (title_level is not None):
        document.add_heading(str(title), level=title_level)
    document.add_paragraph(str(text), style=para_style)



