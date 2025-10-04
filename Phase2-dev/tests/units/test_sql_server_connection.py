import pytest
from unittest.mock import Mock, patch, MagicMock
import pyodbc
from src.services.SqlServerConnection import SqlServerConnection


class TestSqlServerConnection:
    
    @patch.dict('os.environ', {'SQLSERVER_SERVER': 'test_server', 'SQLSERVER_DATABASE': 'test_db'})
    def test_init(self):
        conn = SqlServerConnection()
        assert conn.server == 'test_server'
        assert conn.database == 'test_db'
        assert conn.connection is None

    @patch('src.services.SqlServerConnection.pyodbc.connect')
    @patch.dict('os.environ', {'SQLSERVER_SERVER': 'test_server', 'SQLSERVER_DATABASE': 'test_db'})
    def test_connect_success(self, mock_connect):
        mock_connection = Mock()
        mock_connect.return_value = mock_connection
        
        conn = SqlServerConnection()
        conn.connect()
        
        assert conn.connection == mock_connection
        mock_connect.assert_called_once()

    @patch('src.services.SqlServerConnection.pyodbc.connect')
    @patch.dict('os.environ', {'SQLSERVER_SERVER': 'test_server', 'SQLSERVER_DATABASE': 'test_db'})
    def test_connect_failure(self, mock_connect):
        mock_connect.side_effect = pyodbc.Error("Connection failed")
        
        conn = SqlServerConnection()
        with pytest.raises(pyodbc.Error):
            conn.connect()

    def test_close(self):
        conn = SqlServerConnection()
        mock_connection = Mock()
        conn.connection = mock_connection
        
        conn.close()
        
        mock_connection.close.assert_called_once()

    def test_execute_query_no_connection(self):
        conn = SqlServerConnection()
        
        with pytest.raises(Exception, match="Connection is not established"):
            conn.execute_query("SELECT 1")

    def test_execute_query_success(self):
        conn = SqlServerConnection()
        mock_connection = Mock()
        mock_cursor = Mock()
        mock_connection.cursor.return_value = mock_cursor
        mock_cursor.fetchall.return_value = [('result',)]
        conn.connection = mock_connection
        
        result = conn.execute_query("SELECT 1")
        
        assert result == [('result',)]
        mock_cursor.execute.assert_called_once_with("SELECT 1")
        mock_cursor.close.assert_called_once()

    def test_execute_query_with_params(self):
        conn = SqlServerConnection()
        mock_connection = Mock()
        mock_cursor = Mock()
        mock_connection.cursor.return_value = mock_cursor
        mock_cursor.fetchall.return_value = [('result',)]
        conn.connection = mock_connection
        
        result = conn.execute_query("SELECT * WHERE id = ?", (1,))
        
        assert result == [('result',)]
        mock_cursor.execute.assert_called_once_with("SELECT * WHERE id = ?", (1,))
        mock_cursor.close.assert_called_once()

    def test_execute_query_error(self):
        conn = SqlServerConnection()
        mock_connection = Mock()
        mock_cursor = Mock()
        mock_connection.cursor.return_value = mock_cursor
        mock_cursor.execute.side_effect = pyodbc.Error("Query failed")
        conn.connection = mock_connection
        
        with pytest.raises(pyodbc.Error):
            conn.execute_query("SELECT 1")
        
        mock_cursor.close.assert_called_once()