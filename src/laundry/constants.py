"""
Constants associated with Laundry/
"""
from typing import NewType
import pandas as pd

laundry_version = '2019.0.8'

data_frame = NewType('data_frame', pd.DataFrame)
invalid = ['nan', 'None', 'NA', 'N/A', 'False', 'Nil']
photo_formats = ['.jpg', '.jpeg', '.png', '.tiff']
# photo_formats = ['', '.jpg', '.jpeg', '.png', '.tiff']

