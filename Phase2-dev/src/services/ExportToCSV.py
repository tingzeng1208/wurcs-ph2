import csv
from dotenv import load_dotenv
import os
from UploadtoS3 import upload_file_to_s3

current_dir = os.path.dirname(os.path.abspath(__file__))
dotenv_path = os.path.join(current_dir, '..', '..', '.env')  # Adjust the relative path as needed

# Load environment variables from .env file
load_dotenv()

class ExportToCSV:
  
  def __init__(self):
    
    self.csv_path = os.getenv("LOCAL_CSV_PATH")
  
  
  def export_data_to_csv(self, data, filename):
      """
      Exports a list of tuples to a CSV file.
      
      Args:
          data (list of tuples): Data to be exported.
          file_path (str): Path where the CSV file will be saved.
      """
      
      file_path = os.path.join(self.csv_path, filename)
      if not data:
          print("No data to export.")
          return
        
      try:
          with open(file_path, 'w', newline='') as csvfile:
              writer = csv.writer(csvfile)
              writer.writerows(data)
          print(f"Data successfully exported to {file_path}")
          upload_file_to_s3(file_path, os.getenv("S3_BUCKET"), filename)
      except Exception as e:
          print(f"An error occurred while exporting data to CSV: {e}")
# Example tuple list


