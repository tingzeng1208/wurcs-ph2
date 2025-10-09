import pandas as pd
import pytest
import os
import sys


current_dir = os.path.dirname(os.path.abspath(__file__))
parent_dir = os.path.abspath(os.path.join(current_dir, '../src'))
if parent_dir not in sys.path:
    sys.path.insert(0, parent_dir)
from services.DBManager import DBManager

def _mk_dbm(fake_db_class):
    
    # reset singleton for this module-level helper
    DBManager._instance = None
    return DBManager(connection_class=fake_db_class)

def test_dbmanager_singleton(reset_singletons, fake_db_class):
    
    a = DBManager(connection_class=fake_db_class)
    b = DBManager()  # still returns the existing instance
    assert a is b

def test_execute_sql_delegates(reset_singletons, fake_db_class):
    dbm = _mk_dbm(fake_db_class)
    result = dbm.execute_sql("SELECT COUNT(*) FROM t")
    # Fake returns [(42,)]
    assert result == [(42,)]

def test_get_table_name_caches_and_uses_execute_sql(reset_singletons, fake_db_class):
    # Make a custom fake that counts calls and returns a stable table name
    class CountingFake(fake_db_class):
        def __init__(self):
            super().__init__()
            self.count = 0

        def execute_query(self, query, params=None):
            self.count += 1
            return [("R_TRANS",)]

    dbm = _mk_dbm(CountingFake)
    # First call populates the cache
    name1 = dbm._get_table_name_from_sql("1", "TRANS")
    assert name1 == "R_TRANS"
    # Second call should use cache (no new DB hit)
    name2 = dbm._get_table_name_from_sql("1", "TRANS")
    assert name2 == "R_TRANS"
    assert dbm._get_connection().count == 1

    
def test_get_class1_rail_list_builds_expected_query(reset_singletons, fake_db_class):
    import pandas as pd

    # build DBManager first
    dbm = _mk_dbm(fake_db_class)

    # warm the locator cache ON THIS INSTANCE
    dbm.db_data.table_dict[("1", "CLASS1RAILLIST")] = "R_CLASS1"
    dbm.db_data.database_dict[("1", "CLASS1RAILLIST")] = "urcs_control"

    df = dbm.get_class1_rail_list()
    assert isinstance(df, pd.DataFrame)

    q, _ = dbm._get_connection().query_set_calls[-1]
    assert "FROM urcs_control.R_CLASS1" in q
    assert "WHERE RR_ID<>0" in q
def test_get_custom_data_uses_db_and_table_from_locator(reset_singletons, fake_db_class):
    from data.db_data import db_data
    
    dbm = _mk_dbm(fake_db_class)    
    d = dbm.db_data
    
    d.table_dict[("1", "MyTable")] = "R_MY_TABLE"
    d.database_dict[("1", "MyTable")] = "my_db"

    
    df = dbm.get_custom_data("MyTable")
    assert isinstance(df, pd.DataFrame)
    q, _ = dbm._get_connection().query_set_calls[-1]
    assert q.strip().startswith("SELECT * FROM my_db.R_MY_TABLE")

def test_records_in_ua_values(reset_singletons, fake_db_class):
    from data.db_data import db_data
    
    dbm = _mk_dbm(fake_db_class)
    
    d = dbm.db_data
    d.table_dict[("2024", "AVALUES")] = "U_AVALUES"
    d.database_dict[("2024", "AVALUES")] = "db_a"
    count = dbm.records_in_ua_values("2024")
    assert count == 42  # default fake SELECT returns 42

def test_clear_substitutions_sends_delete(reset_singletons, fake_db_class):
    from data.db_data import db_data
    
    dbm = _mk_dbm(fake_db_class)    
    d = dbm.db_data
    d.table_dict[("2024", "SUBSTITUTIONS")] = "U_SUBS"
    d.database_dict[("2024", "SUBSTITUTIONS")] = "db_s"
    
    dbm.clear_substitutions("2024", "123")
    non_queries = dbm._get_connection().non_query_calls
    assert any("DELETE FROM db_s.U_SUBS WHERE Year = '2024' AND RR_ID = 123" in q for q, _ in non_queries)

def test_insert_substitutions_calls_non_query_for_rows(reset_singletons, fake_db_class, monkeypatch):
    from data.db_data import db_data
    dbm = _mk_dbm(fake_db_class)    
    d = dbm.db_data
    d.table_dict[("2024", "SUBSTITUTIONS")] = "U_SUBS"
    d.database_dict[("2024", "SUBSTITUTIONS")] = "db_s"
    d.table_dict[("1", "ECODES")] = "R_ECODES"
    values = [
        ["eCode", "Value"],   # header row (skipped)
        ["E100", 5],
        ["E200", None],       # becomes "n/a" -> 0
    ]
    dbm.insert_substitutions("2024", "123", values)
    # First call is the clear_substitutions delete; then two INSERTS
    assert len(dbm._get_connection().non_query_calls) >= 3

def test_add_columns_to_df(reset_singletons, fake_db_class):
    dbm = _mk_dbm(fake_db_class)
    row = pd.Series({"a": 1, "b": 2})
    df = dbm.add_columns_to_df(row, ["a", "b"])
    assert df.shape == (1, 2)
    assert list(df.columns) == ["a", "b"]

def test_delete_all_records(reset_singletons, fake_db_class):
    dbm = _mk_dbm(fake_db_class)
    dbm.delete_all_records("dbx", "tblx")
    q, _ = dbm._get_connection().non_query_calls[-1]
    assert q == "DELETE FROM dbx.tblx"
    

def test_generate_e_values_xml_success(reset_singletons, monkeypatch, fake_db_class):
    class CtxConn:
        def __enter__(self): return self
        def __exit__(self, exc_type, exc, tb): return False
        def cursor(self):
            class C:
                def execute(self, q, y): pass
                def fetchone(self): return ("<xml>ok</xml>",)
            return C()

    DBManager._instance = None
    dbm = DBManager(connection_class=fake_db_class)

    # 🔑 Prewarm the locator so _get_database_name_from_sql doesn't query
    dbm.db_data.database_dict[("2024", "EVALUES")] = "dummy_db"

    # Use our context-manageable connection
    monkeypatch.setattr(dbm, "_get_connection", lambda: CtxConn())

    xml = dbm.generate_e_values_xml("2024")
    assert xml == "<xml>ok</xml>"


# def test_generate_e_values_xml_failure_returns_empty(reset_singletons, monkeypatch, fake_db_class):
#     class BadCtxConn:
#         def __enter__(self): return self
#         def __exit__(self, exc_type, exc, tb): return False
#         def cursor(self):
#             class C:
#                 def execute(self, q, y): raise Exception("boom")
#                 def fetchone(self): return None
#     DBManager._instance = None
#     dbm = DBManager(connection_class=fake_db_class)
#     monkeypatch.setattr(dbm, "_get_connection", lambda: BadCtxConn())
#     assert dbm.generate_e_values_xml("2024") == ""
