"""
Units tests for laundry.py
"""
import pytest
from laundry import *


def test_clean_xlsx_table():
    # todo: this test will take a .xlsx file and return a Pandas dataframe.
    #   TEST: a dataframe is returned.
    # todo: what other tests are required for this function?
    pass


def test_remove_underscore():
    test = laundry.remove_underscore('this_is_a_test')
    expected = 'this is a test'
    assert test == expected


def test_section_contains():
    # todo: This test should be parametarised
    test = laundry.section_contains('1234\n5678')
    expected =['1234', '5678']
    assert test == expected


@pytest.mark.parametrize('record,header,format_title,expected',
                         [
                             ({"asset_name": "Storage shed", "component": "Isolator",
                               "defect_type": "Technical Requirement"
                               },
                              ["asset_name", "component", "defect_type"],
                              True,
                              [("Asset Name", "Component", "Defect Type"),
                               ("Storage shed", "Isolator", "Technical Requirement")
                               ]),
                             ({"asset_name": "Storage shed", "component": "Isolator",
                               "defect_type": "Technical Requirement"
                               },
                              ["asset_name", "component", "defect_type"],
                              False,
                              [("asset name", "component", "defect type"),
                               ("Storage shed", "Isolator", "Technical Requirement")]),
                         ])
def test_extract_data(record, header, format_title, expected):
    expected = expected
    result = laundry.extract_data(record, header, format_title)
    assert expected == result


def test_insert_table():
    pass


def test_insert_row():
    pass


def test_insert_photo():
    pass


def test_structure_docs():
    pass