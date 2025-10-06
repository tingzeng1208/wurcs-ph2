import pandas as pd
import sys

import os

current_dir = os.path.dirname(os.path.abspath(__file__))
parent_dir = os.path.abspath(os.path.join(current_dir, '../src'))
if parent_dir not in sys.path:
    sys.path.insert(0, parent_dir)

def test_db_data_singleton(reset_singletons):
    from data.db_data import db_data
    a = db_data()
    b = db_data()
    assert a is b

def test_db_data_save_and_read_csv(tmp_path, reset_singletons):
    from data.db_data import db_data
    d = db_data()
    d.dt_trans = pd.DataFrame({"x": [1, 2]})
    d.dt_dictionary = pd.DataFrame({"y": ["a", "b"]})
    d.dt_railroads_to_process = pd.DataFrame({"z": [10]})
    d.save_to_csv(tmp_path)

    d2 = db_data()
    d2.dt_trans = None
    d2.dt_dictionary = None
    d2.dt_railroads_to_process = None
    d2.read_from_csv(tmp_path)

    assert list(d2.dt_trans.columns) == ["x"]
    assert d2.dt_trans.shape == (2, 1)
    assert list(d2.dt_dictionary.columns) == ["y"]
    assert list(d2.dt_railroads_to_process.columns) == ["z"]
