"""
Constants associated with Laundry/
"""
from typing import NewType
import pandas as pd

laundry_version = '2020.1.1'

data_frame = NewType('data_frame', pd.DataFrame)
invalid = ['nan', 'None', 'NA', 'N/A', 'False', 'Nil']
photo_formats = ['.jpg', '.jpeg', '.png', '.tiff']
