from typing import Optional
import pandas as pd
import os
from dotenv import load_dotenv

current_dir = os.path.dirname(os.path.abspath(__file__))
dotenv_path = os.path.join(current_dir, '..', '..', '.env')  # Adjust the relative path as needed

# Load environment variables from .env file
load_dotenv()

class db_data:
  
  _instance = None
  
  def __new__(cls, *args, **kwargs):
    if not cls._instance:
      cls._instance = super().__new__(cls)
    return cls._instance
  
  def __init__(self):
    self.table_dict = {}
    self.database_dict = {}
    self.dt_trans: Optional[pd.DataFrame] = None
    self.dt_dictionary: Optional[pd.DataFrame] = None
    self.dt_railroads_to_process: Optional[pd.DataFrame] = None
    self.schema = "urcs_control"
    self.csv_path = os.getenv("LOCAL_CSV_PATH")
    self.s_current_year = 2024
    self.RUN_YEARS = 5
    self.CapitalCost = False
    self.VarialabilityFlow = False
    self.Account76 = True
    self.Account80 = False
    self.Account90 = False
    self.SSAC = False
    self.CreateLog = True
    self.NullStringReplace = "N/A"
  
  def save_to_csv(self, dir_path: str = "."):
    """
    Save dt_trans, dt_dictionary, and dt_railroads_to_process to CSV files in the specified directory.
    Args:
        dir_path (str): Directory path to save CSV files. Defaults to current directory.
    """
    try:
      if self.dt_trans is not None:
        self.dt_trans.to_csv(f"{dir_path}/dt_trans.csv", index=False)
        print(f"dt_trans saved to {dir_path}/dt_trans.csv")
      if self.dt_dictionary is not None:
        self.dt_dictionary.to_csv(f"{dir_path}/dt_dictionary.csv", index=False)
        print(f"dt_dictionary saved to {dir_path}/dt_dictionary.csv")
      if self.dt_railroads_to_process is not None:
        self.dt_railroads_to_process.to_csv(f"{dir_path}/dt_railroads_to_process.csv", index=False)
        print(f"dt_railroads_to_process saved to {dir_path}/dt_railroads_to_process.csv")
    except Exception as e:
      print(f"errors at csv output {e}")
      
  def read_from_csv(self, dir_path: str = "."):
    """
    Read dt_trans, dt_dictionary, and dt_railroads_to_process from CSV files in the specified directory.
    Args:
        dir_path (str): Directory path to read CSV files from. Defaults to current directory.
    """
    
    trans_path = os.path.join(dir_path, "dt_trans.csv")
    dict_path = os.path.join(dir_path, "dt_dictionary.csv")
    rr_path = os.path.join(dir_path, "dt_railroads_to_process.csv")
    if os.path.exists(trans_path):
      self.dt_trans = pd.read_csv(trans_path)
    if os.path.exists(dict_path):
      self.dt_dictionary = pd.read_csv(dict_path)
    if os.path.exists(rr_path):
      self.dt_railroads_to_process = pd.read_csv(rr_path)