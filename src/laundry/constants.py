"""
Constants associated with Laundry/
"""
from typing import NewType
import pandas as pd
import numpy as np
laundry_version = '2020.1.2b'

data_frame = NewType('data_frame', pd.DataFrame)
invalid = ['nan', 'None', 'NA', 'N/A', 'False', 'Nil', np.nan]
photo_formats = ['.jpg', '.jpeg', '.png', '.tiff']
