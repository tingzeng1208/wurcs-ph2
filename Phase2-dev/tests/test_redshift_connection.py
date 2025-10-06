

import pandas as pd
import os
import sys
import importlib
import pytest

current_dir = os.path.dirname(os.path.abspath(__file__))
parent_dir = os.path.abspath(os.path.join(current_dir, '../src'))
if parent_dir not in sys.path:
    sys.path.insert(0, parent_dir)
from tests.conftest import FakePsycopg2Module, FakePsycopg2Connection

def test_redshift_connection_singleton(reset_singletons, monkeypatch):
    from services.RedShiftConnection import RedShiftConnection
    a = RedShiftConnection()
    b = RedShiftConnection()
    assert a is b
def test_connect_and_execute_query_success(reset_singletons, monkeypatch):
    # 1) Install fake psycopg2 BEFORE importing RedShiftConnection
    from tests.conftest import FakePsycopg2Connection, FakePsycopg2Module
    fake_conn = FakePsycopg2Connection(rows=[(1, "A")], columns=["id", "name"])
    monkeypatch.setitem(sys.modules, "psycopg2", FakePsycopg2Module(fake_conn))

    # 2) Ensure a clean re-import of the module under test
    sys.modules.pop("services.RedShiftConnection", None)
    RedShiftConnection = importlib.import_module(
        "services.RedShiftConnection"
    ).RedShiftConnection

    # 3) Now use it
    db = RedShiftConnection()
    rows = db.execute_query("SELECT 1, 'A'")
    assert rows == [(1, "A")]

def test_execute_non_query_commit_and_rollback(reset_singletons, monkeypatch):
    from tests.conftest import FakePsycopg2Connection, FakePsycopg2Module

    # ok path
    ok_conn = FakePsycopg2Connection()
    monkeypatch.setitem(sys.modules, "psycopg2", FakePsycopg2Module(ok_conn))
    sys.modules.pop("services.RedShiftConnection", None)
    RedShiftConnection = importlib.import_module(
        "services.RedShiftConnection"
    ).RedShiftConnection

    db = RedShiftConnection()
    db.execute_non_query("UPDATE t SET x=1")
    assert ok_conn._committed is True
    assert ok_conn._rolled_back is False

    # error path -> rollback
    bad_conn = FakePsycopg2Connection(fail_on={"BAD"})
    monkeypatch.setitem(sys.modules, "psycopg2", FakePsycopg2Module(bad_conn))
    sys.modules.pop("services.RedShiftConnection", None)
    RedShiftConnection = importlib.import_module(
        "services.RedShiftConnection"
    ).RedShiftConnection

    db2 = RedShiftConnection()
    with pytest.raises(Exception):
        db2.execute_non_query("UPDATE t SET y=2 -- BAD")
    assert bad_conn._rolled_back is True

def test_execute_query_set_dataframe(reset_singletons, monkeypatch):
    from tests.conftest import FakePsycopg2Connection, FakePsycopg2Module
    fake_conn = FakePsycopg2Connection(rows=[(10, "x"), (20, "y")], columns=["n", "c"])
    monkeypatch.setitem(sys.modules, "psycopg2", FakePsycopg2Module(fake_conn))
    sys.modules.pop("services.RedShiftConnection", None)
    RedShiftConnection = importlib.import_module(
        "services.RedShiftConnection"
    ).RedShiftConnection

    db = RedShiftConnection()
    df = db.execute_query_set("SELECT n, c FROM t")
    assert list(df.columns) == ["n", "c"]
    assert df.shape == (2, 2)
    


def _reload_redshift_with(fake_conn, monkeypatch):
    # 1) Install fake psycopg2 FIRST
    monkeypatch.setitem(sys.modules, "psycopg2", FakePsycopg2Module(fake_conn))
    # 2) Re-import the module fresh
    sys.modules.pop("services.RedShiftConnection", None)
    mod = importlib.import_module("services.RedShiftConnection")
    # 3) Reset singleton
    mod.RedShiftConnection._instance = None
    return mod.RedShiftConnection
    
def test_execute_sql_head(reset_singletons, monkeypatch):
    # Make a fake connection that returns the columns you expect
    fake_conn = FakePsycopg2Connection(rows=[(7, "z")], columns=["a", "b"])
    RedShiftConnection = _reload_redshift_with(fake_conn, monkeypatch)

    db = RedShiftConnection()
    cols, rows = db.execute_sql_head("SELECT a,b FROM t")
    assert cols == ["a", "b"]
    assert rows == [(7, "z")]
