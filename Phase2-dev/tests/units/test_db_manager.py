import pytest
from unittest.mock import Mock, patch, MagicMock
from src.services.DBManager import DBManager


class TestDBManager:
    
    def test_init_default_connection(self):
        manager = DBManager()
        assert manager.connection_class is not None
        assert manager.db_connection is None

    def test_init_custom_connection(self):
        mock_connection_class = Mock()
        manager = DBManager(mock_connection_class)
        assert manager.connection_class == mock_connection_class

    def test_get_connection_creates_new(self):
        mock_connection_class = Mock()
        manager = DBManager(mock_connection_class)
        
        connection = manager._get_connection()
        
        mock_connection_class.assert_called_once()
        assert manager.db_connection == mock_connection_class.return_value

    def test_get_connection_reuses_existing(self):
        mock_connection_class = Mock()
        manager = DBManager(mock_connection_class)
        existing_connection = Mock()
        manager.db_connection = existing_connection
        
        connection = manager._get_connection()
        
        mock_connection_class.assert_not_called()
        assert connection == existing_connection

    @patch('src.services.DBManager.logging')
    def test_execute_sql_success(self, mock_logging):
        mock_connection = Mock()
        mock_connection.execute_query.return_value = [('result',)]
        mock_connection_class = Mock(return_value=mock_connection)
        
        manager = DBManager(mock_connection_class)
        
        result = manager.execute_sql("SELECT 1", ('param',))
        
        assert result == [('result',)]
        mock_connection.connect.assert_called_once()
        mock_connection.execute_query.assert_called_once_with("SELECT 1", ('param',))
        mock_connection.close.assert_called_once()

    @patch('src.services.DBManager.logging')
    def test_execute_sql_no_params(self, mock_logging):
        mock_connection = Mock()
        mock_connection.execute_query.return_value = [('result',)]
        mock_connection_class = Mock(return_value=mock_connection)
        
        manager = DBManager(mock_connection_class)
        
        result = manager.execute_sql("SELECT 1")
        
        assert result == [('result',)]
        mock_connection.execute_query.assert_called_once_with("SELECT 1", None)

    @patch.object(DBManager, 'execute_sql')
    def test_get_table_name_from_sql_success(self, mock_execute_sql):
        mock_execute_sql.return_value = [('test_table',)]
        
        manager = DBManager()
        
        result = manager.get_table_name_from_sql('2023', 'TEST_TYPE')
        
        assert result == 'test_table'
        mock_execute_sql.assert_called_once_with(
            "SELECT Table_Name FROM U_TABLE_LOCATOR WHERE Year = ? AND Data_Type = ?",
            ('2023', 'TEST_TYPE')
        )

    @patch.object(DBManager, 'execute_sql')
    def test_get_table_name_from_sql_no_results(self, mock_execute_sql):
        mock_execute_sql.return_value = []
        
        manager = DBManager()
        
        with pytest.raises(Exception, match="No entry found for year 2023, Data_Type TEST_TYPE"):
            manager.get_table_name_from_sql('2023', 'TEST_TYPE')

    @patch.object(DBManager, 'execute_sql')
    def test_get_table_name_from_sql_uppercase_data_type(self, mock_execute_sql):
        mock_execute_sql.return_value = [('test_table',)]
        
        manager = DBManager()
        
        result = manager.get_table_name_from_sql('2023', 'test_type')
        
        mock_execute_sql.assert_called_once_with(
            "SELECT Table_Name FROM U_TABLE_LOCATOR WHERE Year = ? AND Data_Type = ?",
            ('2023', 'TEST_TYPE')
        )

    @patch('src.services.DBManager.logging')
    def test_execute_sql_connection_error(self, mock_logging):
        mock_connection = Mock()
        mock_connection.connect.side_effect = Exception("Connection failed")
        mock_connection_class = Mock(return_value=mock_connection)
        
        manager = DBManager(mock_connection_class)
        
        with pytest.raises(Exception, match="Connection failed"):
            manager.execute_sql("SELECT 1")
        
        mock_connection.close.assert_called_once()

    @patch('src.services.DBManager.logging')
    def test_execute_sql_logging(self, mock_logging):
        mock_connection = Mock()
        mock_connection.execute_query.return_value = [('result',)]
        mock_connection_class = Mock(return_value=mock_connection)
        
        manager = DBManager(mock_connection_class)
        manager.execute_sql("SELECT 1", ('param',))
        
        mock_logging.info.assert_called_once()