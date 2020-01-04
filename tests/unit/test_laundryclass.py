import pytest
import laundry.laundryclass as laundry
from pathlib import Path, PurePath

struct_dict = {1: 'a', 2: 'b'}
data_dict = {3: 'c', 4: 'd'}
file_template = 'this_is_a_path'
file_path = 'this_file_the_file_path'
file_output = 'output_file_path'


@pytest.mark.parametrize('data_str,expected',
                         [('1234\n5678', ['1234', '5678'])])
def test_split_str(data_str, expected):
    expected = expected
    result = laundry.split_str(data_str)
    # result = laundry.split_str('1234\n5678')
    assert result == expected


@pytest.mark.parametrize('record,header,format_title,expected',
                         [({"asset_name": "Storage shed",
                            "component": "Isolator",
                            "defect_type": "Technical Requirement"},
                           ["asset_name", "component", "defect_type"],
                           True,
                           [("Asset Name", "Component", "Defect Type"),
                            ("Storage shed", "Isolator", "Technical Requirement")]),
                          ({"asset_name": "Storage shed",
                            "component": "Isolator",
                            "defect_type": "Technical Requirement"},
                           ["asset_name", "component", "defect_type"],
                           False,
                           [("asset name", "component", "defect type"),
                            ("Storage shed", "Isolator", "Technical Requirement")])
                          ])
def test_sort_table_data(record, header, format_title, expected):
    expected = expected
    result = laundry.sort_table_data(record, header, format_title)
    assert expected == result


def test_remove_underscore():
    expected = 'this is a test'
    assert laundry.remove_underscore('this_is_a_test') == expected


# @pytest.mark.parametrize('test_path,expected', [[[r'\\..\unit'], Path(r'../unit')],
#                                                 [[r'/../../src'], Path(r'../../src')],
#                                                 [[r'/test_laundryclass.py'], r'Incorrect path.']
#                                                 ])
# def test_confirm_directory_path(test_path, expected):
#     assert laundry.resolve_file_path(test_path) == expected


def test_resolve_file_path():
    p = './test_laundryclass.py'
    q = Path(p)
    expected = q.resolve()
    assert laundry.resolve_file_path(p) == expected


def test_value_exists():
    v_expected = {'a', 'b', 'c'}
    v_actual = {'1', 'a', '2', 'b', '3', 'c'}
    assert laundry.values_exist(v_expected, v_actual) is True


@pytest.mark.parametrize('values,args,expected', [[('123', 'abc', None), [None], ['123', 'abc']],
                                                  [('123', 'abc', None), ['abc', None], ['123']],
                                                  ])
def test_remove_from_list(values, args, expected):
    assert laundry.remove_from_iterable(values, *args) == expected


# def test_singleload__init__():
#     """This test is limited to testing that the object is created successfully."""
#     obj = laundry.SingleLoad(struct_dict, data_dict, file_template, file_path, file_output)
#     assert type(obj) is laundry.SingleLoad


# def test_singleload_split_into_rows():
#     obj = laundry.SingleLoad(struct_dict, data_dict, file_template, file_path, file_output)
#     result = [{3: 'c'}, {4: 'd'}]
#     obj.split_into_rows()
#     assert obj._row_data == result


def start_wash():
    pass


def test_format_docx():
    pass


def test_insert_table():
    pass


def test_insert_paragraph():
    pass


def test_insert_row():
    pass


def test_photo():
    pass


def test_laundry__init__worksheet():
    test_obj = laundry.Laundry(Path('../../resources/input_files/test_spreadsheet_laundryclass.xlsm'),
                               data_worksheet='Master List', structure_worksheet='_structure',
                               output_file=Path('../../resources/output_files/test_laundry_output.docx'),
                               template_file=Path('../../resources/templates/CONVERSION_TEMPLATE.docx'), header_row=5)
    template_file = laundry.resolve_file_path('../../resources/templates/CONVERSION_TEMPLATE.docx')
    output_file = laundry.resolve_file_path('../../resources/output_files/test_laundry_output.docx')
    test_obj.wash_load(template_file, output_file)
    assert isinstance(test_obj, laundry.Laundry)


def test_laundry_check_worksheets_basic():
    """Test to confirm that an exception is raised if three worksheets are not provided at object instantiation."""
    with pytest.raises(TypeError) as excinfo:
        laundry.Laundry.check_worksheets_basic([])
    exception_msg = excinfo.value.args[0]
    assert exception_msg == 'Either the "data" and "structure" worksheets, or the "batch" worksheet must be provided.'


def test_laundry_resolve_file_path():
    """Test the file not found exception."""
    with pytest.raises(FileNotFoundError) as excinfo:
        laundry.resolve_file_path('this_file_does_not_exist')
    exception_msg = excinfo.value.args[1]
    assert exception_msg == 'No such file or directory'

def test_laundry_check_batch_data():
    pass


def test_laundry_excel_to_dataframe():
    pass


def test_laundry_prepare_row_filters():
    pass
