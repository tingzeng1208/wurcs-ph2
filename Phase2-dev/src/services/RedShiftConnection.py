import os
import psycopg2
from dotenv import load_dotenv
import pandas as pd

current_dir = os.path.dirname(os.path.abspath(__file__))
dotenv_path = os.path.join(current_dir, '..', '..', '.env')  # Adjust the relative path as needed

# Load environment variables from .env file
load_dotenv()

class RedShiftConnection:
    
    _instance = None
    """
    RedShiftConnection is a class for managing connections to an Amazon Redshift database.
    Attributes:
        host (str): The Redshift cluster endpoint.
        database (str): The name of the database to connect to.
        user (str): The username for the database.
        password (str): The password for the database.
        connection (psycopg2.Connection): The connection object to the Redshift database.
    Methods:
        __init__(database=None): Initializes the RedShiftConnection instance, reading connection details from environment variables.
        connect(): Establishes a connection to the Redshift database.
        close(): Closes the connection to the Redshift database.
        execute_query(query, params=None): Executes a SQL query and returns the results as a list of tuples.
        execute_query_set(query, params=None): Executes a SQL query and returns the results as a Pandas DataFrame.
    """
    
    def __new__(cls, *args, **kwargs):
        if cls._instance is None:
            cls._instance = super().__new__(cls)
        return cls._instance
    
    def __init__(self, database=None):
        # Read connection details from environment variables
        self.host = os.getenv("REDSHIFT_HOST")
        self.user = os.getenv("REDSHIFT_USER")
        self.port = os.getenv("REDSHIFT_PORT", "5439")  # Default Redshift port is 5439
        self.password = os.getenv("REDSHIFT_PASSWORD")
        if (database is None):
            self.database = os.getenv("REDSHIFT_DATABASE")
        else:
            self.database = database
        self.connection = None
        
    def _get_connection(self):
        if not self.connection or self.connection.closed:
            self.connect()
        return self.connection

    def connect(self):
        """
        Establish a connection to the Redshift database.
        """
        try:
            
            self.connection = psycopg2.connect(
                dbname=self.database,
                user=self.user,
                password=self.password,
                host=self.host,
                port=self.port  # Default Redshift port
            )
            print("Connection to Redshift established successfully.")
        except psycopg2.Error as e:
            print(f"Error connecting to Redshift: {e}")
            raise

    def close(self):
        """
        Close the connection to the Redshift database.
        """
        if self.connection:
            self.connection.close()
            print("Connection to Redshift closed.")

    def execute_query(self, query, params=None):
        """
        Execute a SQL query and return the results.
        """
        if not self._get_connection():
            raise Exception("Connection is not established. Call connect() first.")
        
        cursor = self._get_connection().cursor()
        try:
            if params:
                cursor.execute(query, params)
            else:
                cursor.execute(query)
            return cursor.fetchall()
        except psycopg2.Error as e:
            print(f"Error executing query: {e}")
            raise
        finally:
            cursor.close()
            
    def execute_non_query(self, query, params=None):
        """
        Execute a SQL query and return the results.
        """
        conn = self._get_connection()
        if conn is None:
            raise Exception("Connection is not established. Call connect() first.")
        
        cursor = conn.cursor()
        try:
            if params:
                cursor.execute(query, params)
            else:
                cursor.execute(query)
            conn.commit()
        except psycopg2.Error as e:
            conn.rollback()
            print(f"Error executing query: {e}")
            raise
        finally:
            cursor.close()
            
    def execute_query_set(self, query, params=None):
        """
        Execute a SQL query and return the results.
        """
        conn = self._get_connection()
        if conn is None:
            raise Exception("Connection is not established. Call connect() first.")
        
        cursor = conn.cursor()
        
        try:
            if params:
                cursor.execute(query, params)
            else:
                cursor.execute(query)
            rows = cursor.fetchall()

            # Get column names
            columns = [column[0] for column in cursor.description]

            # Create a Pandas DataFrame
            df = pd.DataFrame.from_records(rows, columns=columns)
            return df
        except psycopg2.Error as e:
            print(f"Error executing query: {e}")
            raise
        finally:
            cursor.close()
            
    def execute_sql_head(self, query, params=None):
        """
        Execute a SQL query and return both field names and field data.
        Returns:
            tuple: (columns, rows)
        """
        conn = self._get_connection()
        if conn is None:
            raise Exception("Connection is not established. Call connect() first.")
        cursor = conn.cursor()
        try:
            if params:
                cursor.execute(query, params)
            else:
                cursor.execute(query)
            rows = cursor.fetchall()
            columns = [desc[0] for desc in cursor.description]
            return columns, rows
        except psycopg2.Error as e:
            print(f"Error executing query: {e}")
            raise
        finally:
            cursor.close()

# Example usage
if __name__ == "__main__":
    
    db = RedShiftConnection()
    try:        
        # Example query
        results = db.execute_query("SELECT Database_Name FROM URCS_CONTROL.U_TABLE_LOCATOR WHERE Year = 1 AND Data_Type = 'TRANS'")
        print(results)
        query1 = "SELECT Table_Name FROM URCS_CONTROL.U_TABLE_LOCATOR WHERE Year =%s AND Data_Type = %s"
        param = (1, "TRANS")
        results = db.execute_query(query1, param)
        print(results)
        query3 = "SELECT * FROM URCS_CONTROL.U_TABLE_LOCATOR"
        results = db.execute_query_set(query3)
        print(results)

    finally:
        db.close()