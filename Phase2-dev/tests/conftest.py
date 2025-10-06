import os
import sys, types
import types
import pandas as pd
import pytest
import importlib

# ---- Existing Helpers / Fakes -------------------------------------------------

class FakeCursor:
    def __init__(self, rows=None, columns=None, fail_on=None):
        self._rows = rows or []
        self._columns = columns or []
        self._executed = []
        self._fail_on = fail_on or set()

    @property
    def description(self):
        if not self._columns:
            return None
        # mimic psycopg2: description is tuple objects whose first item is name
        return [(c, None, None, None, None, None, None) for c in self._columns]

    def execute(self, query, params=None):
        self._executed.append((query, params))
        if any(key in query for key in self._fail_on):
            raise Exception("FAKE_DB_ERROR")

    def fetchall(self):
        return list(self._rows)

    def fetchone(self):
        return self._rows[0] if self._rows else None

    def close(self):
        pass


class FakePsycopg2Connection:
    def __init__(self, rows=None, columns=None, fail_on=None):
        self.closed = 0
        self._rows = rows or []
        self._columns = columns or []
        self._fail_on = fail_on or set()
        self._committed = False
        self._rolled_back = False
        self._cursors = []

    def cursor(self):
        cur = FakeCursor(self._rows, self._columns, self._fail_on)
        self._cursors.append(cur)
        return cur

    def commit(self):
        self._committed = True

    def rollback(self):
        self._rolled_back = True

    def close(self):
        self.closed = 1

    # Support context manager to mimic psycopg2 connections
    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, tb):
        self.close()
        return False  # don't suppress exceptions


class FakePsycopg2Module(types.SimpleNamespace):
    def __init__(self, connection: FakePsycopg2Connection):
        super().__init__()
        self._connection = connection
        self.Error = Exception

    def connect(self, **kwargs):
        return self._connection


class FakeDBConnection:
    """A very small facade that mimics RedShiftConnection public API."""
    def __init__(self):
        self.query_calls = []
        self.non_query_calls = []
        self.query_set_calls = []

        # default return payloads
        self._result_rows = [(42,)]
        self._df = pd.DataFrame({"Line": [1], "C1": [100]})

    def execute_query(self, query, params=None):
        self.query_calls.append((query, params))
        return self._result_rows

    def execute_non_query(self, query, params=None):
        self.non_query_calls.append((query, params))
        return None

    def execute_query_set(self, query, params=None):
        self.query_set_calls.append((query, params))
        return self._df


# ---- New Helper / Fake for Phase2InputData -----------------------------------

class FakeDBManager:
    """DBManager stub for Phase2InputData tests."""
    def __init__(self):
        from data.db_data import db_data
        self.db_data = db_data()
        # reset frames per instance
        self.db_data.dt_trans = None
        self.db_data.dt_dictionary = None
        self.db_data.dt_railroads_to_process = None
        self.adjust_calls = []
        self.wrote_ur_acode_batch = None
        self.non_query_calls = []

    # counts/clears used in clear_previous_data()
    def records_in_ur_acode_data(self, year): return 123
    def records_in_ua_values(self, year): return 456
    def clear_ur_acode_data(self, year): pass
    def clear_UAValues(self, year): pass

    # called by _trans_modifications()
    def adjust_u_trans_values(self, col, new_val, year, rricc, sch, line):
        self.adjust_calls.append((col, float(new_val), year, rricc, sch, line))

    # called by other_processing()
    def write_ur_acode_data_batch(self, data_list):
        self.wrote_ur_acode_batch = list(data_list)

    # prepare() pulls these three
    def get_trans_table(self, year):
        cols = ["year","rricc","sch","line"] + [f"c{i}" for i in range(1,16)]
        df = pd.DataFrame([
            [int(year)-4, 910000, 120, 10] + [1,2,3,4,5,6,7,8,9,10,11, 0,0,0,14],
            [int(year)-4, 910000, 120, 11] + [1,1,1,1,1,1,1,1,1, 1, 1, 0,0,0,15],
            [int(year)-3, 910000, 420, 3 ] + [1,2,3,4,5,6,7,8,9,10,11, 0,0,99,13],
            [int(year)-2, 910000, 33, 57] + [1,2,3,4,5,6,7,8,9,10,11, 0,0,0,12],
            [int(year),   111111, 5,  7 ] + [10,0,0,0,0,0,0,0,0,0,0,  0,0,0,0],
        ], columns=cols)
        return df

    def get_data_dictionary(self, year):
        return pd.DataFrame([
            {"wtall":"A1C1L101","sch":5,"line":7,"column":1,"scaler":2,"loadcode":3},
        ])

    def get_class1_rail_data_to_prepare(self, year):
        return pd.DataFrame([
            {"rr_id": 1, "rricc": 111111, "regionrricc": 910000, "nationrricc": 990000},
            {"rr_id": 2, "rricc": 222222, "regionrricc": 910000, "nationrricc": 990000},
        ])

    def execute_non_query(self, q, db=None):
        self.non_query_calls.append(q)


# ---- Fixtures ----------------------------------------------------------------

@pytest.fixture(autouse=True)
def env_vars(monkeypatch, tmp_path):
    # Minimal Redshift env so RedShiftConnection reads clean values
    monkeypatch.setenv("REDSHIFT_HOST", "fake-host")
    monkeypatch.setenv("REDSHIFT_USER", "fake-user")
    monkeypatch.setenv("REDSHIFT_PASSWORD", "fake-pass")
    monkeypatch.setenv("REDSHIFT_DATABASE", "fake-db")
    monkeypatch.setenv("REDSHIFT_PORT", "5439")
    # For Phase2InputData file writes
    monkeypatch.setenv("LOCAL_CSV_PATH", str(tmp_path))
    monkeypatch.setenv("S3_BUCKET", "unit-test-bucket")
    monkeypatch.setenv("IAM_ROLE_ARN", "arn:aws:iam::123:role/fake")
    monkeypatch.setenv("aws_access_key_id", "AKIA_FAKE")
    monkeypatch.setenv("aws_secret_access_key", "SECRET_FAKE")
    yield


@pytest.fixture
def reset_singletons():
    from data.db_data import db_data
    from services.DBManager import DBManager
    # DO NOT import services.RedShiftConnection here (prevents early psycopg2 import)
    db_data._instance = None
    DBManager._instance = None
    yield
    db_data._instance = None
    DBManager._instance = None


@pytest.fixture
def fake_db_class():
    """Factory returning a connection_class you can pass into DBManager."""
    return FakeDBConnection


@pytest.fixture
def fake_db():
    return FakeDBConnection()


# ---- Extras for RedShift tests ------------------------------------------------

@pytest.fixture
def reload_redshift_with(monkeypatch):
    """
    Returns a helper: lambda fake_conn -> RedShiftConnection class
    Patches sys.modules['psycopg2'] BEFORE importing services.RedShiftConnection,
    reloads the module, and resets its singleton.
    """
    def _loader(fake_conn):
        monkeypatch.setitem(sys.modules, "psycopg2", FakePsycopg2Module(fake_conn))
        sys.modules.pop("services.RedShiftConnection", None)
        mod = importlib.import_module("services.RedShiftConnection")
        mod.RedShiftConnection._instance = None
        return mod.RedShiftConnection
    return _loader


# ---- Patching S3 layer (for Phase2InputData.upload_to_redshift) ---------------

@pytest.fixture
def patch_s3_layer(monkeypatch):
    """
    Provides (calls, pre, post):
      - pre(): install a package-like fake boto3 (and its submodules) BEFORE importing modules
               that might do "from boto3.s3.transfer import S3Transfer".
      - post(module): patch the Phase2InputData module's imported symbols
                      (upload_file_to_s3, copyS3ToRedShift, boto3.client) AFTER import.
    """
    calls = {"upload": [], "copy": [], "delete": []}

    def pre():
        # Build a minimal package tree: boto3, boto3.s3, boto3.s3.transfer
        fake_boto3 = types.ModuleType("boto3")

        class _FakeS3Client:
            def delete_object(self, Bucket, Key):
                calls["delete"].append((Bucket, Key))

        def client(name):
            assert name == "s3"
            return _FakeS3Client()

        fake_boto3.client = client

        fake_boto3_s3 = types.ModuleType("boto3.s3")
        fake_boto3_s3_transfer = types.ModuleType("boto3.s3.transfer")

        class S3Transfer:
            pass  # only needed if some helper imports it

        fake_boto3_s3_transfer.S3Transfer = S3Transfer

        # Register the package and submodules
        sys.modules["boto3"] = fake_boto3
        sys.modules["boto3.s3"] = fake_boto3_s3
        sys.modules["boto3.s3.transfer"] = fake_boto3_s3_transfer

    def post(phase2_module):
        # Replace symbols imported into Phase2InputData module
        monkeypatch.setattr(
            phase2_module, "upload_file_to_s3",
            lambda path, bucket, key: calls["upload"].append((path, bucket, key)),
            raising=True
        )
        monkeypatch.setattr(
            phase2_module, "copyS3ToRedShift",
            lambda bucket, key, odb, table, ak, sk, year: calls["copy"].append(
                (bucket, key, table, year)
            ),
            raising=True
        )
        # Ensure its boto3 reference uses our fake's client()
        # (only needed if the module captured boto3 at import time)
        import types as _types
        if not hasattr(phase2_module, "boto3"):
            # If the module didn’t bind boto3 name, nothing to do
            return
        # ensure .client exists and points to our fake
        def _client(name):
            class _FakeS3Client:
                def delete_object(self, Bucket, Key):
                    calls["delete"].append((Bucket, Key))
            return _FakeS3Client()
        phase2_module.boto3.client = _client

    return calls, pre, post


# ---- Public fixtures for Phase2InputData tests --------------------------------

@pytest.fixture
def fake_dbm():
    """A fresh FakeDBManager for each test."""
    return FakeDBManager()
