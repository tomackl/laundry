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
    result = laundry.split_str('1234\n5678')
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


@pytest.mark.parametrize('test_path,expected', [[[r'\\..\unit'], Path(r'../unit')],
                                                [[r'/../../src'], Path(r'../../src')],
                                                [[r'/test_laundryclass.py'], r'Incorrect path.']
                                                ])
def test_confirm_directory_path(test_path, expected):
    assert laundry.confirm_directory_path(test_path) == expected


def test_file_path_exists():
    p = './test_laundryclass.py'
    q = Path(p)
    expected = q.resolve()
    assert laundry.file_path_exists(p) == expected


def test_value_exists():
    v_expected = {'a', 'b', 'c'}
    v_actual = {'1', 'a', '2', 'b', '3', 'c'}
    assert laundry.values_exist(v_expected, v_actual) is True


@pytest.mark.parametrize('values,args,expected', [[('123', 'abc', None), [None], ['123', 'abc']],
                                                  [('123', 'abc', None), ['abc', None], ['123']],
                                                  ])
def test_remove_from_list(values, args, expected):
    assert laundry.remove_from_iterable(values, *args) == expected


def test_singleload__init__():
    """This test is limited to testing that the object is created successfully."""
    obj = laundry.SingleLoad(struct_dict, data_dict, file_template, file_path, file_output)
    assert type(obj) is laundry.SingleLoad


def test_singleload_split_into_rows():
    obj = laundry.SingleLoad(struct_dict, data_dict, file_template, file_path, file_output)
    result = [{3: 'c'}, {4: 'd'}]
    obj.split_into_rows()
    assert obj._row_data == result


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


def test_laundry__init__TypeError():
    """Test to confirm that an exception is raised if three worksheets are not provided at object instantiation."""
    pass


def test_laundry__init__FileNotFoundError():
    """Test the file not found exception."""
    pass

