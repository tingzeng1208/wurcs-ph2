import os, sys
import pandas as pd
import importlib
current_dir = os.path.dirname(os.path.abspath(__file__))
parent_dir = os.path.abspath(os.path.join(current_dir, '../src/core'))
if parent_dir not in sys.path:
    sys.path.insert(0, parent_dir)

# Adjust this import path to your file location as needed:
# e.g., from core.Phase2InputData import Phase2InputData
from Phase2InputData import Phase2InputData as _Phase2InputData


def _fresh_module():
    """Re-import module for clean monkeypatching target each time."""
    import sys
    sys.modules.pop("core.Phase2InputData", None)
    return importlib.import_module("core.Phase2InputData")


def test_scale_and_return_decimal(fake_dbm):
    mod = _fresh_module()
    Phase2InputData = mod.Phase2InputData

    p = Phase2InputData(current_year="2024", db_manager=fake_dbm)
    assert p._scale(0, 123) == 123
    assert p._scale(1, 123) == 12.3
    assert p._scale(2, 123) == 1.23
    assert p._scale(6, 1_000_000) == 1.0

    assert p._return_decimal("3.5") == 3.5
    assert p._return_decimal(None) == 0.0
    assert p._return_decimal("bad") == 0.0


def test_get_report_sheet_mapping(fake_dbm):
    mod = _fresh_module()
    Phase2InputData = mod.Phase2InputData
    p = Phase2InputData(current_year="2024", db_manager=fake_dbm)

    assert p._get_report_sheet("A1", 150) == "A1P1"
    assert p._get_report_sheet("A1", 505) == "A1P5A"
    assert p._get_report_sheet("A2", 210) == "A2P2"
    assert p._get_report_sheet("A3", 520) == "A3P5"
    assert p._get_report_sheet("A4", 202) == "A4P3"
    assert p._get_report_sheet("E2", 123) == "E2P1"
    assert p._get_report_sheet("ZZ", 999) == ""


def test_get_value_filters_and_scales(fake_dbm):
    mod = _fresh_module()
    Phase2InputData = mod.Phase2InputData
    p = Phase2InputData(current_year="2024", db_manager=fake_dbm)

    # Arrange to hit the extra row returned by FakeDBManager.get_trans_table()
    # row: year=2024, rricc=111111, sch=5, line=7, c1=10
    fake_dbm.db_data.dt_trans = fake_dbm.get_trans_table("2024")
    p.i_sch = 5
    p.i_line = 7
    p.i_col = 1        # maps to c1
    p.i_scale = 2      # divide by 100
    p.i_load_code = 3  # i_rricc_code = self.i_rricc
    p.i_rricc = 111111
    p.i_rricc_region = 910000
    p.i_rricc_nation = 990000

    val = p.get_value("2024")
    # c1=10, scaler 2 => 0.1
    assert val == 0.1


def test_trans_modifications_updates_and_calls_adjust(fake_dbm):
    mod = _fresh_module()
    Phase2InputData = mod.Phase2InputData
    p = Phase2InputData(current_year="2024", db_manager=fake_dbm)

    fake_dbm.db_data.dt_trans = fake_dbm.get_trans_table("2024")

    # before
    pre = fake_dbm.db_data.dt_trans.copy()
    p._trans_modifications("2024")

    # some values should have been updated; verify adjust() was called
    assert len(fake_dbm.adjust_calls) > 0

    # spot-check: the sch==420 row had c12 recomputed as c2+c3 = 2+3=5
    df = fake_dbm.db_data.dt_trans
    r420 = df[(df["year"] == 2021) & (df["sch"] == 420)].iloc[0]
    assert float(r420["c12"]) == 5.0

    # sch in 101..147 should trigger c12 = 2*c1 + c3 + c5 for the first row
    r120 = df[(df["year"] == 2020) & (df["sch"] == 120)].iloc[0]
    assert float(r120["c12"]) == (2*1 + 3 + 5)  # =10.0


def _fresh_import(modname: str):
    sys.modules.pop(modname, None)
    return importlib.import_module(modname)

def test_other_processing_end_to_end_upload_and_batch(fake_dbm, patch_s3_layer):
    calls, pre, post = patch_s3_layer
    pre()

    mod = _fresh_import("core.Phase2InputData")  # adjust path if needed
    Phase2InputData = mod.Phase2InputData
    post(mod)

    p = Phase2InputData(current_year="2024", db_manager=fake_dbm)

    # 👇 set frames AFTER constructing p (constructor resets them to None)
    p.db_data.dt_trans = fake_dbm.get_trans_table("2024")
    p.db_data.dt_dictionary = fake_dbm.get_data_dictionary("2024")
    p.db_data.dt_railroads_to_process = fake_dbm.get_class1_rail_data_to_prepare("2024")

    ok = p.other_processing()
    assert ok is True

    assert len(calls["upload"]) == 1
    assert len(calls["copy"]) == 1
    assert len(calls["delete"]) == 1
    assert fake_dbm.wrote_ur_acode_batch is not None
    assert len(fake_dbm.wrote_ur_acode_batch) == len(p.db_data.dt_dictionary)


def test_prepare_happy_path(fake_dbm, patch_s3_layer):
    calls, pre, post = patch_s3_layer

    pre()  # install fake boto3 tree before import

    mod = _fresh_module()  # or your existing fresh-import helper
    Phase2InputData = mod.Phase2InputData

    post(mod)  # patch module-level symbols (upload_file_to_s3, copyS3ToRedShift, boto3.client)

    p = Phase2InputData(current_year="2024", db_manager=fake_dbm)
    ok = p.prepare()
    assert ok is True
    assert len(calls["upload"]) == 1
    assert len(calls["copy"]) == 1
    assert len(calls["delete"]) == 1
