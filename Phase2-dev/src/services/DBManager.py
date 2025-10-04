import logging
import pandas as pd
import os
import sys
from datetime import datetime
from typing import List, Any

# Ensure the parent directory is in sys.path for module imports
current_dir = os.path.dirname(os.path.abspath(__file__))
parent_dir = os.path.abspath(os.path.join(current_dir, '..'))
if parent_dir not in sys.path:
    sys.path.insert(0, parent_dir)

from data.db_data import db_data
from services.RedShiftConnection import RedShiftConnection  # Make sure this import path is correct

class DBManager:

    _instance = None

    def __new__(cls, *args, **kwargs):
        if cls._instance is None:
            cls._instance = super().__new__(cls)
        return cls._instance
    
    def __init__(self, connection_class=RedShiftConnection):
        """
        Initialize DBManager with flexible connection class for future changes.
        """
        self.connection_class = connection_class
        self.db_connection = None
        self.db_data = db_data()
        
    def _get_connection(self):
        """Get or create database connection."""
        if not self.db_connection:
            self.db_connection = self.connection_class()
        return self.db_connection
    
    def execute_sql(self, query, params=None):
        """
        Execute SQL query and return results.
        
        Args:
            query (str): SQL query to execute
            params (tuple, optional): Parameters for the query
            
        Returns:
            list: Query results
        """
        db = self._get_connection()
        
        try:
            logging.info("Executing SQL: %s with params %s", query, params)
            return db.execute_query(query, params)
        except Exception as e:
            logging.error("Error executing SQL: %s with params %s", query, params, exc_info=True)
            raise Exception(f"Error executing SQL: {e}")
                
    def _get_database_name_from_sql(self, year, data_type):
        """
        Gets the database name value from Table Locator table in URCS Controls database.
        
        Args:
            year (str): Year value
            data_type (str): Data type value
            
        Returns:
            str: Database Name
            
        Raises:
            Exception: If no entry found for the given year and data_type
        """
        
        if  self.db_data.database_dict.get((year, data_type)) is not None:
            return self.db_data.database_dict.get((year, data_type))
        query = f"SELECT Database_Name from {self.db_data.schema}.U_TABLE_LOCATOR WHERE Year = %s AND Data_Type = %s"
        params = (year, data_type.upper())
        
        results = self.execute_sql(query, params)
        
        if not results:
            raise Exception(f"No entry found for year {year}, Data_Type {data_type}")
        
        self.db_data.database_dict[(year, data_type)] = results[0][0]
        return results[0][0]
                
    def get_trans_table(self, year):
#         2025-07-31 16:24:29.453 - Get_Database_Name_From_SQL: SELECT Database_Name FROM U_TABLE_LOCATOR WHERE Year = 1 AND Data_Type = 'TRANS'
# 2025-07-31 16:24:29.495 - Get_Table_Name_From_SQL: SELECT Table_Name FROM U_TABLE_LOCATOR WHERE Year = 1 AND Data_Type = 'TRANS'
        dbname = self._get_database_name_from_sql("1", "TRANS")
        tablename = self._get_table_name_from_sql("1", "TRANS")
        query = f"SELECT * FROM {self.db_data.schema}.{tablename} where year between %s and %s"
        prev_year = int(year) - 4
        params = (prev_year, year)
        df = self.execute_sql_query_set(query, params)
        num_rows, num_columns = df.shape

        print(f"GetTransTable: {year} Number of records: {num_rows}")
        print(f"GetTransTable: {year} Number of columns: {num_columns}")
        
        return df
    
    def get_data_dictionary(self, year):    
        
        tablename = self._get_table_name_from_sql(1, "DATA_DICTIONARY")
        # SELECT * FROM R_Data_Dictionary WHERE Effective_Dt <= 2023 And Expiration_Dt > 2023 ORDER BY URCSID
        query = f"SELECT * FROM {self.db_data.schema}.{tablename} where Effective_Dt <= %s and Expiration_Dt > %s ORDER BY URCSID"
        params = (year, year)        
        df = self.execute_sql_query_set(query, params)
        num_rows, num_columns = df.shape

        print(f"GetDataDictionary: {year} Number of records: {num_rows}")
        print(f"GetDataDictionary: {year} Number of columns: {num_columns}")
        
        return df
        
    def _get_table_name_from_sql(self, year, data_type):
        """
        Gets the table name value from Table Locator table in URCS Controls database.
        
        Args:
            year (str): Year value
            data_type (str): Data type value
            
        Returns:
            str: Table Name
            
        Raises:
            Exception: If no entry found for the given year and data_type
        """
        
        if self.db_data.table_dict.get((year, data_type)) is not None:
            return self.db_data.table_dict.get((year, data_type))
        
        query = f"SELECT Table_Name FROM {self.db_data.schema}.U_TABLE_LOCATOR WHERE Year = %s AND UPPER(Data_Type) = %s"
        params = (year, data_type.upper())
        
        results = self.execute_sql(query, params)
        
        if not results:
            raise Exception(f"No entry found for year {year}, Data_Type {data_type}")
        
        self.db_data.table_dict[(year, data_type)] = results[0][0]
        logging.info(f"Table Name for Year {year}, Data_Type {data_type}")
        return results[0][0]
    
    def execute_sql_query_set(self, query, params=None):
        """
        Execute SQL query and return results as a list of dictionaries.

        Args:
            query (str): SQL query to execute
            params (tuple, optional): Parameters for the query

        Returns:
            list: Query results as a list of dictionaries
        """
        db = self._get_connection()

        try:
            logging.info(f"Executing SQL: {query} with params {params}")
            df = db.execute_query_set(query, params)
            return df
        except Exception as e:
            logging.error(f"Error executing SQL: {query} with params {params}", exc_info=True)
            raise Exception(f"Error executing SQL: {e}")

    def get_class1_rail_list(self) -> pd.DataFrame:
        try:
            database_name = self._get_database_name_from_sql("1", "CLASS1RAILLIST")
            table_name = self._get_table_name_from_sql("1", "CLASS1RAILLIST")
            
            query = f"SELECT * FROM {self.db_data.schema}.{table_name} WHERE RR_ID<>0 ORDER BY RR_ID"
            df = self.execute_sql_query_set(query)
            return df
            
        except Exception as e:
            raise Exception("Error when retrieving CLASS1RAILLIST table values", e)

    def get_class1_rail_data_to_prepare(self, current_year: str) -> pd.DataFrame:
        try:
            database_name = self._get_database_name_from_sql("1", "CLASS1RAILLIST")
            table_name = self._get_table_name_from_sql("1", "CLASS1RAILLIST")
            
            query = f"""
            SELECT *, 
                   (SELECT TOP 1 RRICC FROM {self.db_data.schema}.{table_name} WHERE REGION_ID = R.REGION_ID And RRICC > 900000) RegionRRICC,
                   (Select TOP 1 RRICC FROM {self.db_data.schema}.{table_name} WHERE REGION_ID = 0 And RRICC > 900000) NationRRICC 
            FROM {self.db_data.schema}.{table_name} R 
            WHERE EFFECTIVE_YEAR <={current_year} And EXPIRATION_YEAR > {current_year}
            """
            
            df = self.execute_sql_query_set(query)
            return df
            
        except Exception as e:
            raise Exception("Error When retrieving CLASS1RAILLIST table values", e)

    def get_custom_data(self, table_name: str) -> pd.DataFrame:
        """
        Retrieves all data from the specified table, similar to the VB.NET GetCustomData function.
        Args:
            table_name (str): The name of the table to query.
        Returns:
            pd.DataFrame: DataFrame containing all rows from the table.
        Raises:
            Exception: If SQL error occurs.
        """
        try:
            # Get the actual table and database names from the locator
            actual_table_name = self._get_table_name_from_sql("1", table_name)
            actual_database_name = self._get_database_name_from_sql("1", table_name)
            query = f"SELECT * FROM {actual_database_name}.{actual_table_name}"
            logging.info(f"GetCustomData: {query}")
            df = self.execute_sql_query_set(query)
            df.name = table_name  # Optionally set a name attribute for the DataFrame
            return df
        except Exception as e:
            raise Exception(f"SQL Error When executing query for {table_name}", e)

    def get_a_value(self, railroad_number: int, current_year: str) -> pd.DataFrame:
        try:
            avalues_table = self._get_table_name_from_sql(current_year, "AValues")
            avalues_db = self._get_database_name_from_sql(current_year, "AValues")
            acode_table = self._get_table_name_from_sql("1", "ACodes") 
            acode_db = self._get_database_name_from_sql("1", "ACodes")
            
            query = f"""
            Select at.Year,at.aCode_id,at.Value,ac.aLine, ac.Rpt_sheet, ac.aColumn
            FROM {avalues_db}.{avalues_table} at 
            JOIN {acode_db}.{acode_table} ac On at.aCode_id = ac.aCode_id
            WHERE RR_Id = {railroad_number} ORDER BY acode_id
            """
            print(f"get_a_value query is {str(query)}")
            # add column to df
            columns = ["Year", "aCode_id", "Value", "aLine", "Rpt_sheet", "aColumn"]
            df = self.execute_sql_query_set(query)
            return df
            
            
        except Exception as e:
            raise Exception(f"SQL Error When retrieving U_AVALUES. SQL statement: {query}", e)

    def get_a_value0_rr(self, current_year: str) -> pd.DataFrame:
        try:
            avalues_table = self._get_table_name_from_sql(current_year, "AValues")
            avalues_db = self._get_database_name_from_sql(current_year, "AValues")
            acode_table = self._get_table_name_from_sql("1", "ACodes")
            controls_db = self.db_data.schema
            
            query = f"""
            Select at.[Year],at.[aCode_id],at.[Value],ac.[aLine], ac.[Rpt_sheet] 
            FROM {avalues_db}.{avalues_table} as at 
            JOIN {controls_db}.{acode_table} as ac on at.aCode_id = ac.aCode_id 
            WHERE RR_Id = 0 ORDER BY acode_id
            """
            # add column to df
            
            print(f"get_a_value0_rr query is {str(query)}")
            
            df = self.execute_sql_query_set(query)
            return df
            
        except Exception as e:
            raise Exception(f"SQL error when retrieving U_AVALUES. SQL statement: {query}", e)

    def _get_controls_database_name(self):
        return self.db_data.schema
    
    def get_a_value_region_rr(self, railroad_number: int, current_year: str) -> pd.DataFrame:
        try:
            avalues_table = self._get_table_name_from_sql(current_year, "AValues")
            avalues_db = self._get_database_name_from_sql(current_year, "AValues")
            acode_table = self._get_table_name_from_sql("1", "ACodes")
            acode_db = self._get_database_name_from_sql("1", "ACodes")
            region_table = self._get_table_name_from_sql("1", "Region")
            class1_table = self._get_table_name_from_sql("1", "Class1RailList")
            controls_db = self._get_controls_database_name()
            
            query = f"""
            SELECT at.[Year],at.[aCode_id],at.[Value],ac.[aLine], ac.[Rpt_sheet], ac.[Code] 
            FROM {avalues_db}.{avalues_table} at 
            JOIN {acode_db}.{acode_table} ac on at.aCode_id = ac.aCode_id 
            WHERE RR_Id = (SELECT TOP 1 RR_ID FROM {controls_db}.{class1_table} 
                           WHERE SHORT_NAME = (SELECT TOP 1 Description 
                                             FROM {controls_db}.{region_table} reg INNER JOIN {controls_db}.{class1_table} rr ON reg.id = rr.REGION_ID 
                                             WHERE RR_ID = {railroad_number})) ORDER BY acode_id
            """
            
            # add column to df
            columns = ["Year", "aCode_id", "Value", "aLine", "Rpt_sheet", "Code"]
            df = self.execute_sql_query_set(query)
            return df
            
        except Exception as e:
            raise Exception(f"SQL error when retrieving U_AVALUES. SQL statement: {query}", e)

    def get_price_indexes(self, railroad_number: int, year: str) -> pd.DataFrame:
        try:
            price_index_table = self._get_table_name_from_sql("1", "Index")
            class1_table = self._get_table_name_from_sql("1", "Class1RailList")
            controls_db = self._get_controls_database_name()
            
            query = f"""
            SELECT * 
            FROM {controls_db}.{price_index_table} p 
            INNER JOIN {controls_db}.{class1_table} rr ON p.Region = rr.REGION_ID 
            WHERE rr.RR_ID = {railroad_number} AND YEAR = {year}
            """
            
            df = self.execute_sql_query_set(query)
            return df
            
        except Exception as e:
            raise Exception(f"SQL error when retrieving U_PRICE_INDEX. SQL statement: {query}", e)

    def get_car_type_statistics(self, railroad_number: int) -> pd.DataFrame:
        try:
            car_type_table = self._get_table_name_from_sql("1", "Op_Stats_By_Car_Type")
            region_table = self._get_table_name_from_sql("1", "Region")
            class1_table = self._get_table_name_from_sql("1", "Class1RailList")
            controls_db = self._get_controls_database_name()
            
            query = f"""
            SELECT [Line],[C1],[C2],[C3],[C4],[C5],[C6],[C7],[C8],[C9],[C10],[C11] 
            FROM {controls_db}.{car_type_table} c 
            INNER JOIN {controls_db}.{region_table} r ON c.Region = r.description 
            INNER JOIN {controls_db}.{class1_table} rr ON r.id = rr.REGION_ID 
            WHERE rr.RR_ID = {railroad_number}
            """
            
            df = self.execute_sql_query_set(query)
            return df
            
        except Exception as e:
            raise Exception(f"SQL error when retrieving R_OP_STATS_BY_CAR_TYPE. SQL statement: {query}", e)

    def get_car_type_statistics_part2(self, railroad_number: int) -> pd.DataFrame:
        try:
            car_type2_table = self._get_table_name_from_sql("1", "Op_Stats_By_Car_Type_2")
            region_table = self._get_table_name_from_sql("1", "Region")
            class1_table = self._get_table_name_from_sql("1", "Class1RailList")
            controls_db = self._get_controls_database_name()
            
            query = f"""
            SELECT [Line],[C1],[C2],[C3],[C4],[C5],[C6],[C7],[C8],[C9],[C10],[C11],[C12],[C13],[C14] 
            FROM {controls_db}.{car_type2_table} c 
            INNER JOIN {controls_db}.{region_table} r ON c.Region = r.description 
            INNER JOIN {controls_db}.{class1_table} rr ON r.id = rr.REGION_ID 
            WHERE rr.RR_ID = {railroad_number}
            """
            
            df = self.execute_sql(query)
            return df
            
        except Exception as e:
            raise Exception(f"SQL error when retrieving R_OP_STATS_BY_CAR_TYPE_2. SQL statement: {query}", e)

    def get_car_type_statistics_part3(self) -> pd.DataFrame:
        try:
            car_type3_table = self._get_table_name_from_sql("1", "Op_Stats_By_Car_Type_3")
            controls_db = self._get_controls_database_name()
            
            query = f"SELECT [Line],[C1],[C2],[C3],[C4],[C5] FROM {controls_db}.{car_type3_table} ORDER BY Line"
            
            df = self.execute_sql_query_set(query)
            return df
            
        except Exception as e:
            raise Exception(f"SQL error when retrieving R_OP_STATS_BY_CAR_TYPE_3. SQL statement: {query}", e)

    def get_line_source_text(self) -> pd.DataFrame:
        try:
            line_source_table = self._get_table_name_from_sql("1", "Line_Source_Text")
            controls_db = self._get_controls_database_name()
            
            query = f"SELECT * FROM {controls_db}.{line_source_table} ORDER BY Rpt_sheet,Line"
            
            print(f"GetLineSourceText:{str(query)}")
            df = self.execute_sql_query_set(query)
            return df
        except Exception as e:
            raise Exception(f"SQL error when retrieving R_LINE_SOURCE_TEXT. SQL statement: {query}", e)
                
    def handle_error(self, current_year: str, data: str, message: str, stack: str, location: str):
        try:
            timestamp = datetime.now().strftime("%m/%d/%Y %I:%M:%S.%f %p")
            message = message.replace("'", "").replace('"', "")
            stack = stack.replace("'", "").replace('"', "")
            
            errors_table = self._get_table_name_from_sql(current_year, "ERRORS")
            errors_db = self._get_database_name_from_sql(current_year, "ERRORS")
            
            query = f"""
            INSERT INTO {errors_db}.{errors_table} VALUES ('{data}','{timestamp}','{message}','{stack}','{location}')
            """
            
            self.execute_non_query(query)
            logging.info(f"HandleError: {query}")
            
        except Exception as e:
            logging.error(f"Error in handle_error: {str(e)}")

    def clear_substitutions(self, year: str, rr_id: str):
        try:
            substitutions_table = self._get_table_name_from_sql(year, "SUBSTITUTIONS")
            substitutions_db = self._get_database_name_from_sql(year, "SUBSTITUTIONS")
            
            query = f"DELETE FROM {substitutions_db}.{substitutions_table} WHERE Year = '{year}' AND RR_ID = {rr_id}"
            
            self.execute_non_query(query)
            logging.info(f"ClearSubstitutions: {query}")
            
        except Exception as e:
            logging.error(f"Error in clear_substitutions: {str(e)}")

    def insert_substitutions(self, year: str, rr_id: str, values: List[List[Any]]):
        try:
            self.clear_substitutions(year, rr_id)
            
            substitutions_table = self._get_table_name_from_sql(year, "SUBSTITUTIONS")
            ecode_table = self._get_table_name_from_sql("1", "ECODES")
            substitutions_db = self._get_database_name_from_sql(year, "SUBSTITUTIONS")
            controls_db = self._get_controls_database_name()
            
            timestamp = datetime.now().strftime("%m/%d/%Y %I:%M:%S %p")
            
            for i in range(1, len(values)):
                if len(values[i]) >= 2:
                    ecode = values[i][0] if values[i][0] is not None else ""
                    value = values[i][1] if values[i][1] is not None else "n/a"
                    
                    final_value = 0 if value == "n/a" else value
                    
                    query = f"""
                    INSERT INTO {substitutions_table} 
                    (Year, RR_id, eCode_id, Value, Final_Value, entry_dt) VALUES ('{year}', {rr_id}, 
                    (SELECT TOP 1 eCode_id FROM {controls_db}.dbo.{ecode_table} WHERE eCode = '{ecode}'), 
                    {final_value},{final_value}, '{timestamp}')
                    """
                    self.execute_non_query(query)
                    logging.info(f"InsertSubstitutions: {query}")
                    
        except Exception as e:
            raise Exception(f"SQL error when inserting SUBSTITUTIONS records", e)

    def run_substitutions(self, year: str, rr_ids: List[int]):
        try:
            substitutions_db = self._get_database_name_from_sql(year, "SUBSTITUTIONS")
            
            rr_ids_sorted = sorted(rr_ids, reverse=True)
            
            for rr_id in rr_ids_sorted:
                query = f"EXECUTE usp_RunSubstitutions '{year}',{rr_id}"
                self.execute_non_query(query)
                logging.info(f"RunSubstitutions: {query}")
                
        except Exception as e:
            raise Exception("Error in run_substitutions", e)
        
    def clear_UAValues(self, year: str):
        try:
            avalues_table = self._get_table_name_from_sql(year, "AVALUES")
            avalues_db = self._get_database_name_from_sql(year, "AVALUES")
            
            query = f"DELETE FROM {avalues_db}.{avalues_table}"
            logging.info(f"ClearUAValues: {query}")
            self.execute_non_query(query)
            
        except Exception as e:
            raise Exception("Error when deleting records from U_AVALUES table", e)

    def clear_ur_acode_data(self, year: str):
        try:
            acode_table = self._get_table_name_from_sql(year, "ACODES")
            acode_db = self._get_database_name_from_sql(year, "ACODES")
            
            query = f"DELETE FROM {acode_db}.{acode_table}"
            logging.info(f"ClearURACodeData: {query}")
            self.execute_non_query(query, acode_db)
            
        except Exception as e:
            raise Exception("Error when deleting records from U_ACODES table", e)
        
    def records_in_ur_acode_data(self, year: str) -> int:
        try:
            acode_table = self._get_table_name_from_sql(year, "ACODES")
            acode_db = self._get_database_name_from_sql(year, "ACODES")

            query = f"SELECT COUNT(*) FROM {acode_db}.{acode_table}"
            logging.info(f"RecordsInURACodeData: {query}")

            return self.execute_sql(query)[0][0]
        except Exception as e:
            raise Exception("Error when counting records in U_ACODES table", e)
    
    def records_in_ua_values(self, year: str) -> int:
        try:
            avalues_table = self._get_table_name_from_sql(year, "AVALUES")
            avalues_db = self._get_database_name_from_sql(year, "AVALUES")

            query = f"SELECT COUNT(*) FROM {avalues_db}.{avalues_table}"
            logging.info(f"RecordsInUAValues: {query}")

            return self.execute_sql(query)[0][0]

        except Exception as e:
            raise Exception("Error when counting records in U_AVALUES table", e)
        
    def create_e_values(self, year: str):
        try:
            self.clear_e_values(year)
            
            substitutions_table = self._get_table_name_from_sql(year, "SUBSTITUTIONS")
            evalues_table = self._get_table_name_from_sql(year, "EVALUES")
            evalues_db = self._get_database_name_from_sql(year, "EVALUES")
            
            query = f"""
            INSERT INTO {evalues_table} 
            SELECT Year, RR_Id, eCode_id, Final_Value, entry_dt 
            FROM {substitutions_table} WHERE Year = '{year}'
            """
            self.execute_non_query(query)
            logging.info(f"CreateEValues: {query}")
            
        except Exception as e:
            raise Exception(f"SQL error when writing E_VALUES. SQL statement: {query}", e)

    def clear_e_values(self, year: str):
        try:
            evalues_table = self._get_table_name_from_sql(year, "EVALUES")
            evalues_db = self._get_database_name_from_sql(year, "EVALUES")
            
            query = f"DELETE FROM {evalues_table} WHERE Year = '{year}'"
            self.execute_non_query(query)
            logging.info(f"ClearEValues: {query}")
            
        except Exception as e:
            raise Exception("Error when deleting records from U_EVALUES table", e)

    def adjust_u_trans_values(self, col: int, new_val: float, year: str, rricc: str, sch: str, line: str):
        try:
            trans_table = self._get_table_name_from_sql("1", "TRANS")
            trans_db = self._get_database_name_from_sql("1", "TRANS")
            
            query = f"""
            UPDATE {trans_table} SET C{col} = {new_val} 
            WHERE Year = {year} AND RRICC = {rricc} AND SCH = {sch} AND LINE = {line}
            """
            
            self.execute_non_query(query, trans_db)
            
        except Exception as e:
            raise Exception("Error in adjust_u_trans_values", e)

    def write_ur_acode_data_batch(self, data_list: List[tuple]):
        try:
            acode_table = self._get_table_name_from_sql("1", "ACODES")
            
            values_str = ','.join([f"('{row[0]}',{row[1]},{row[2]},'{row[3]}','{row[4]}','{row[5]}','{row[6]}')" for row in data_list])
            
            
            query = f"""
            INSERT INTO {self.db_data.schema}.{acode_table} (aCode, aColumn, aLine, LineA, APart, Code, Rpt_sheet)
            VALUES {values_str}
            """
            
            self.execute_non_query(query, None)
            logging.info(f"Batch inserted {len(data_list)} records")
            
        except Exception as e:
            logging.error(f"Error in batch insert: {str(e)}")

    def write_ua_values_batch(self, data_list: List[tuple], current_year: str):
        try:
            if not data_list:
                return
                
            year = data_list[0][0]
            avalues_table = self._get_table_name_from_sql(current_year, "AVALUES")              
            db_name_ua_values = self._get_database_name_from_sql(current_year, "AVALUES")
        
            timestamp = datetime.now().strftime("%m/%d/%Y %I:%M:%S %p")
            
            values_str = ','.join([f"('{row[0]}',{row[1]},{row[2]},{row[3]},'{timestamp}')" for row in data_list])
            
            query = f"""
            INSERT INTO {db_name_ua_values}.{avalues_table} (Year, RR_id, aCode_id, Value, entry_dt)
            VALUES {values_str}
            """
            
            self.execute_non_query(query, None)
            logging.info(f"Batch inserted {len(data_list)} UA values")
            
        except Exception as e:
            logging.error(f"Error in batch UA values: {str(e)}")
            
    def execute_non_query(self, query, db_name=None):
        """
        Execute a non-query SQL command (e.g., INSERT, UPDATE, DELETE).
        
        Args:
            query (str): SQL command to execute
            db_name (str): Database name
            
        Raises:
            Exception: If an error occurs during execution
        """
        db = self._get_connection()
        
        try:
            logging.info(f"Executing non-query SQL: {query}")
            db.execute_non_query(query)
        except Exception as e:
            logging.error(f"Error executing non-query SQL: {query}", exc_info=True)
            raise Exception(f"Error executing non-query SQL: {e}")
        
    def _get_current_year(self, current_year: str = "1"):
        """
        Get the current year value from the database.
        
        Args:
            current_year (str): Year value, default is "1"
            
        Returns:
            str: Current year value
        """
        return current_year

    def add_columns_to_df(self, row: pd.Series, columns: List[str]) -> pd.DataFrame:
        return pd.DataFrame([row], columns=columns)
        # return pd.DataFrame.from_records(row, columns=columns)

    def truncate_a_tables(self, current_year: str = "1"):
        try:
            acode_table = self._get_table_name_from_sql("1", "ACODES")
            acode_db = self._get_database_name_from_sql("1", "ACODES")
            
            query1 = f"DELETE {acode_table}"
            logging.info(f"TruncateATables (1): {query1}")
            
            query2 = f"DBCC CHECKIDENT('{acode_table}', RESEED, 0)"
            logging.info(f"TruncateATables (2): {query2}")
            
            current_year = self._get_current_year(current_year)
            avalues_table = self._get_table_name_from_sql(current_year, "AVALUES") 
            avalues_db = self._get_database_name_from_sql(current_year, "AVALUES")
            
            query3 = f"DELETE {avalues_table}"
            logging.info(f"TruncateATables (3): {query3}")
            
        except Exception as e:
            logging.error(f"Error in truncate_a_tables: {str(e)}")

    def generate_e_values_xml(self, year: str) -> str:
        try:
            evalues_db = self._get_database_name_from_sql(year, "EVALUES")
            
            with self._get_connection() as conn:
                cursor = conn.cursor()
                cursor.execute("usp_GenerateEValuesXML ?", year)
                
                result = cursor.fetchone()
                return result[0] if result else ""
                
        except Exception as e:
            logging.error(f"Error generating XML: {str(e)}")
            return ""

    def delete_all_records(self, database_name: str, table_name: str):
        """
        Delete all records from a table in the specified database.
        Args:
            database_name (str): The name of the database.
            table_name (str): The name of the table.
        """
        try:
            query = f"DELETE FROM {database_name}.{table_name}"
            self.execute_non_query(query, database_name)
            logging.info(f"Deleted all records from {database_name}.{table_name}")
        except Exception as e:
            logging.error(f"Error deleting all records from {database_name}.{table_name}: {str(e)}")

if __name__ == "__main__":
    dbManager = DBManager()
    tablename = dbManager._get_table_name_from_sql("1", "URCS_YEARS")
    print(f"Table Name: {tablename}")
    dbname = dbManager._get_database_name_from_sql("1", "URCS_YEARS")
    print(f"Database Name: {dbname}")
    dbManager.get_trans_table("2023")
    dbManager.get_data_dictionary("2023")
    dbManager._get_table_name_from_sql("1", "URCS_YEARS")

