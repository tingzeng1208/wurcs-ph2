import os
import sys
import pandas as pd
from typing import Dict, Optional, Callable, Any
import time
from datetime import datetime            
import csv
import uuid
import boto3
from dotenv import load_dotenv

load_dotenv()

aws_access_key_id = os.getenv("aws_access_key_id")
aws_secret_access_key = os.getenv("aws_secret_access_key")

current_dir = os.path.dirname(os.path.abspath(__file__))
parent_dir = os.path.abspath(os.path.join(current_dir, '..'))
if parent_dir not in sys.path:
    sys.path.insert(0, parent_dir)
    
from services.DBManager import DBManager
from services.UploadtoS3 import upload_file_to_s3
from services.CopyS3RedShift import copyS3ToRedShift
from data.db_data import db_data
from model.Report import ReportSingleton



class Phase2InputData:
    
    def __init__(self, current_year: str, db_manager: DBManager = None):
        
        if db_manager is None:
            db_manager = DBManager()
        self.s_processing_year: str = ""
        self.i_acode_id: int = 0
        self.s_acode: str = ""
        self.i_acode_counter: int = 0
        self.s_a_column: str = ""
        self.s_a_line: str = ""
        self.s_line_a: str = ""
        self.s_a_part: str = ""
        self.s_code: str = ""
        self.s_car_type_id: str = ""
        self.i_car_type_id: int = 0
        self.i_rricc: int = 0
        self.i_rricc_region: int = 0
        self.i_rricc_nation: int = 0
        self.i_rr_id: int = 0
        self.d_value: float = 0.0
        self.i_sch: int = 0
        self.i_line: int = 0
        self.i_col: int = 0
        self.i_scale: int = 0
        self.i_load_code: int = 0
        self.s_rpt_sheet: str = ""
        self.i_this_row: int = 0
        self.db_data = db_manager.db_data
        self.report_singleton = ReportSingleton()
        
        if db_manager is None:
            db_manager = DBManager()
        
        self.db_data.dt_trans: Optional[pd.DataFrame] = None
        self.db_data.dt_dictionary: Optional[pd.DataFrame] = None
        self.db_data.dt_railroads_to_process: Optional[pd.DataFrame] = None
        
        self.o_db = db_manager
        self.s_current_year = current_year
        
        self.status_updated_callback: Optional[Callable[[str, bool], None]] = None
        self.error_occurred_callback: Optional[Callable[[str], None]] = self.error_callback

    def error_callback(self, error_message: str):
        # """
        # Callback function to handle errors during processing.
        # Args:
        #     error_message (str): The error message to handle.
        # """
        print(f"Error occurred: {error_message}")
    
    def clear_previous_data(self):
        
        ur_acode_records = self.o_db.records_in_ur_acode_data("1")
        ua_values_records = self.o_db.records_in_ua_values(self.s_current_year)
        
        print(f"\nRecords in ur_acode_data to remove: {ur_acode_records}")
        print(f"Records in ua_values to remove: {ua_values_records}")   
                
        self.o_db.clear_ur_acode_data("1")
        self.o_db.clear_UAValues(self.s_current_year)
        
    def other_processing(self) -> bool:
        
        self.i_acode_id = 1
        self.i_acode_counter = len(self.db_data.dt_dictionary)
        
        batch_ur_acode_data = []
        batch_ua_values = []

        self.clear_previous_data()
      
        for dr_index, dr in self.db_data.dt_dictionary.iterrows():
            self.i_this_row += 1
            
            if self.i_this_row % 300 == 0:
                print(f"Processing row {self.i_this_row} of {self.i_acode_counter} of dr_index: {dr_index}")

            self.s_acode = str(dr["wtall"])
            c_index = self.s_acode.find("C")
            l_index = self.s_acode.find("L")
            
            if c_index >= 0:
                self.s_a_column = self.s_acode[c_index + 1:c_index + 2]
            if l_index >= 0:
                self.s_a_line = self.s_acode[l_index + 1:l_index + 4]
                self.s_line_a = self.s_acode[l_index:l_index + 4]
            
            self.s_a_part = self.s_acode[0:2]
            if c_index >= 0:
                self.s_code = self.s_acode[c_index:c_index + 2]

            self.s_rpt_sheet = self._get_report_sheet(self.s_a_part, int(self.s_a_line))

            # self.o_db.write_ur_acode_data(
            #     self.s_acode, self.s_a_column, self.s_a_line, 
            #     self.s_line_a, self.s_a_part, self.s_code, self.s_rpt_sheet
            # )

            batch_ur_acode_data.append((self.s_acode, self.s_a_column, self.s_a_line,
                                        self.s_line_a, self.s_a_part, self.s_code, self.s_rpt_sheet))
            
            self.i_sch = int(dr["sch"])
            self.i_line = int(dr["line"])
            self.i_col = int(dr["column"])
            self.i_scale = int(dr["scaler"])
            self.i_load_code = int(dr["loadcode"])

            i = int(self.s_current_year) - 4

            # print(f"i_sch: {self.i_sch}, i_line: {self.i_line}, i_col: {self.i_col}, i_scale: {self.i_scale}, i_load_code: {self.i_load_code}")
            # self.o_db.clear_e_values(self.s_current_year)
            while i <= int(self.s_current_year):
                o_region_values: Dict[int, float] = {}
                
                for dr_rr in self.db_data.dt_railroads_to_process.itertuples(index=False):
                    self.s_processing_year = str(i)
                    self.i_rr_id = int(dr_rr.rr_id)
                    self.i_rricc = int(dr_rr.rricc)
                    self.i_rricc_region = int(dr_rr.regionrricc)
                    self.i_rricc_nation = int(dr_rr.nationrricc)
                    # print(f"Processing year: {self.s_processing_year}, rr_id: {self.i_rr_id}, acode_id: {self.i_acode_id}, rricc: {self.i_rricc}, region: {self.i_rricc_region}, nation: {self.i_rricc_nation}")

                    if self.i_rricc_region not in o_region_values:
                        o_region_values[self.i_rricc_region] = 0.0

                    if self.i_rricc > 900000 and self.i_load_code == 0:
                        self.d_value = o_region_values[self.i_rricc_region]
                    else:
                        self.d_value = self._get_value(self.s_processing_year)
                        o_region_values[self.i_rricc_region] += self.d_value
                        
                    # if self.d_value > 0:
                    #     print(f"Value found for year {self.s_processing_year}, rr_id {self.i_rr_id}, acode_id {self.i_acode_id}: {self.d_value}")

                    # self.o_db.write_ua_values(
                    #     self.s_processing_year, str(self.i_rr_id), 
                    #     str(self.i_acode_id), str(self.d_value)
                    # )
                    
                    now_str = datetime.now().strftime("%m/%d/%Y %I:%M:%S %p")
                    batch_ua_values.append((self.s_processing_year, str(self.i_rr_id),
                                            str(self.i_acode_id), str(self.d_value), now_str))

                i += 1

            # batch insertion into database    
            # if ((dr_index + 1) % 100 ==0):
            #         if not batch_ua_values:
            #             print("No ua_values to write for this batch.")
            #         else:
            #             print(f"dr_index: {dr_index+1}, Batch size for ua_values: {len(batch_ua_values)}")
            #             self.o_db.write_ua_values_batch(batch_ua_values, self.s_current_year)
            #             batch_ua_values = []
                        
            self.i_acode_id += 1

      
        
        # Export batch_ua_values to CSV, upload to S3, and copy into Redshift
        if batch_ua_values:
           upload_success = self.upload_to_redshift(batch_ua_values)  
        else:
            print("No ua_values to write for this batch.")
            raise ValueError("No ua_values to write for this batch.")
            return False
        
        print(f"Batch size for ur_acode_data: {len(batch_ur_acode_data)}")
        self.o_db.write_ur_acode_data_batch(batch_ur_acode_data)     
        
        return True
    
    def get_value(self, s_year: str) -> float:
        """
        Python equivalent of VB.NET GetValue function.
        """
        d = 0.0
        i_rricc_code = 0

        # Determine i_rricc_code based on i_load_code
        if self.i_load_code in [1, 4, 5]:
            i_rricc_code = self.i_rricc_region
        elif self.i_load_code == 2:
            i_rricc_code = self.i_rricc_nation
        elif self.i_load_code == 3:
            i_rricc_code = self.i_rricc
        else:
            i_rricc_code = self.i_rricc

        # Filter DataFrame similar to DataRow.Select(s)
        filtered_trans = self.db_data.dt_trans[
            (self.db_data.dt_trans["year"] == int(s_year)) &
            (self.db_data.dt_trans["rricc"] == i_rricc_code) &
            (self.db_data.dt_trans["sch"] == self.i_sch) &
            (self.db_data.dt_trans["line"] == self.i_line)
        ]

        # Map i_col to column name
        col_map = {
            1: "c1", 2: "c2", 3: "c3", 4: "c4", 5: "c5",
            6: "c6", 7: "c7", 8: "c8", 9: "c9", 10: "c10",
            11: "c11", 12: "c12", 13: "c13", 14: "c14", 15: "c15"
        }
        col_name = col_map.get(self.i_col, None)

        for _, dr_trans in filtered_trans.iterrows():
            if col_name and col_name in dr_trans:
                d = self._scale(self.i_scale, self._return_decimal(dr_trans[col_name]))
            break  # Only process the first matching row, as in VB.NET

        return float(d)

    def upload_to_redshift(self, batch_ua_values):
        """
        Uploads the batch of ua_values to Redshift.
        
        Args:
            batch_ua_values (list of tuples): The batch of ua_values to upload.
        """
        # Generate a unique filename for the batch
        
        try:
            csv_filename = f"ua_values_batch_{uuid.uuid4().hex}.csv"
            csv_path = os.path.join(os.getenv("LOCAL_CSV_PATH"), csv_filename)
            
            # Write batch_ua_values to CSV
            with open(csv_path, mode='w', newline='') as csvfile:
                writer = csv.writer(csvfile)
                writer.writerow(["year", "rr_id", "acode_id", "value", "entry_dt"])
                writer.writerows(batch_ua_values)
            print(f"Exported batch_ua_values to {csv_path} with {len(batch_ua_values)} records")
            
            # Upload CSV to S3
            s3_bucket = os.getenv("S3_BUCKET")  # TODO: set your bucket name
            s3_key = f"ua_values_batches/{csv_filename}"
            upload_file_to_s3(csv_path, s3_bucket, s3_key)
            print(f"Uploaded {csv_path} to s3://{s3_bucket}/{s3_key}")
            
            iam_role = os.getenv("IAM_ROLE_ARN")
            
            # Copy from S3 into Redshift
            redshift_table = f"{self.o_db.db_data.schema}.U_AVALUES"
            copyS3ToRedShift(s3_bucket, s3_key, self.o_db, redshift_table, aws_access_key_id, aws_secret_access_key, self.s_current_year)

            print(f"Copied data from s3://{s3_bucket}/{s3_key} into Redshift table {redshift_table}")

            # Remove the S3 object after successful copy
            
            s3 = boto3.client('s3')
            try:
                s3.delete_object(Bucket=s3_bucket, Key=s3_key)
                print(f"Deleted {s3_key} from S3 bucket {s3_bucket}")
            except Exception as del_e:
                print(f"Warning: Failed to delete {s3_key} from S3: {del_e}")

            # Optionally, remove the local CSV file after upload
            os.remove(csv_path)
        except Exception as e:
            print(f"An error occurred during upload to Redshift: {e}")
            raise e
        
        return True
    
    
    def temp_copy_to_redshift(self, csv_filename: str):
        
        try:
            
            s3_bucket = os.getenv("S3_BUCKET")  # TODO: set your bucket name
            s3_key = f"ua_values_batches/{csv_filename}"
            csv_path = os.path.join(os.getenv("LOCAL_CSV_PATH"), csv_filename)
            # upload_file_to_s3(csv_path, s3_bucket, s3_key)
            print(f"Uploaded {csv_path} to s3://{s3_bucket}/{s3_key}")
            
            iam_role = os.getenv("IAM_ROLE_ARN")
            
            databasename = self.o_db._get_database_name_from_sql(self.s_current_year, "AVALUES")
            
            # Copy from S3 into Redshift
            
            # IAM_ROLE '{iam_role}'  -- TODO: set your Redshift IAM role ARN
            redshift_table = f"{databasename}.U_AVALUES"  # Adjust table name as needed
            copy_sql = f"""
            COPY {redshift_table}
            FROM 's3://{s3_bucket}/{s3_key}'
            CREDENTIALS 'aws_access_key_id={aws_access_key_id};aws_secret_access_key={aws_secret_access_key}'
            FORMAT AS CSV
            IGNOREHEADER 1
            TIMEFORMAT 'auto';
            """
            self.o_db.execute_non_query(copy_sql, None)
            print(f"Copied data from s3://{s3_bucket}/{s3_key} into Redshift table {redshift_table}")
         
        except Exception as e:
            print(f"An error occurred during upload to Redshift: {e}")
            raise e   
    
    
     
    def prepare(self) -> bool:
        
        try: 
 
            # start preparing data for Phase 2
            print("Starting preparation of Phase 2 input data...")
            start_save = time.time()
            
            self.db_data.dt_trans = self.o_db.get_trans_table(self.s_current_year)
            self.db_data.dt_dictionary = self.o_db.get_data_dictionary(self.s_current_year)
            self.db_data.dt_railroads_to_process = self.o_db.get_class1_rail_data_to_prepare(self.s_current_year)

            self._trans_modifications(self.s_current_year)
            
            self.db_data.save_to_csv()
            
            self.other_processing()
            
            end_save2 = time.time()
            print(f"Elapsed time for all processing: {end_save2 - start_save:.2f} seconds")
                        
            return True

        except Exception as ex:
            if self.error_occurred_callback:
                self.error_occurred_callback(str(ex))
            return False

    def _trans_modifications(self, current_year: str):
        
        print(self.db_data.dt_trans.columns.tolist())
        for i_process_year in range(int(current_year) - 4, int(current_year) + 1):
            filtered_trans = self.db_data.dt_trans[
                (self.db_data.dt_trans["year"] == i_process_year) & 
                (self.db_data.dt_trans["sch"] > 100) & 
                (self.db_data.dt_trans["sch"] < 148)
            ]
            print("size of filtered trans for year", i_process_year, ":", len(filtered_trans))
            
            numberRowsModified = 0
            
            for idx, dr in filtered_trans.iterrows():
                f_value = (dr["c1"] * 2) + dr["c3"] + dr["c5"]
                if f_value != float(dr["c12"]):
                    self.db_data.dt_trans.at[idx, "c12"] = f_value
                    numberRowsModified += 1
                    self.o_db.adjust_u_trans_values(12, f_value, str(i_process_year), 
                                                  str(dr["rricc"]), str(dr["sch"]), str(dr["line"]))

                f_value = dr["c1"] + dr["c3"] + dr["c5"] + dr["c7"]
                if f_value != float(dr["c13"]):
                    numberRowsModified += 1
                    self.db_data.dt_trans.at[idx, "c13"] = f_value
                    self.o_db.adjust_u_trans_values(13, f_value, str(i_process_year), 
                                                  str(dr["rricc"]), str(dr["sch"]), str(dr["line"]))

                f_value = dr["c3"] + dr["c5"] + (dr["c7"] * 2)
                if f_value != float(dr["c14"]):
                    numberRowsModified += 1
                    self.db_data.dt_trans.at[idx, "c14"] = f_value
                    self.o_db.adjust_u_trans_values(14, f_value, str(i_process_year), 
                                                  str(dr["rricc"]), str(dr["sch"]), str(dr["line"]))

            filtered_trans_420 = self.db_data.dt_trans[
                (self.db_data.dt_trans["year"] == i_process_year) & 
                (self.db_data.dt_trans["sch"] == 420)
            ]
            
            print("size of filtered trans 420 for year", i_process_year, ":", len(filtered_trans_420))
            
            for idx, dr in filtered_trans_420.iterrows():
                f_value = dr["c2"] + dr["c3"]
                if f_value != float(dr["c12"]):
                    numberRowsModified += 1
                    self.db_data.dt_trans.at[idx, "c12"] = f_value
                    self.o_db.adjust_u_trans_values(12, f_value, str(i_process_year), 
                                                  str(dr["rricc"]), str(dr["sch"]), str(dr["line"]))

            filtered_trans_33_57 = self.db_data.dt_trans[
                (self.db_data.dt_trans["year"] == i_process_year) & 
                (self.db_data.dt_trans["sch"] == 33) & 
                (self.db_data.dt_trans["line"] == 57)
            ]
            
            print("size of filtered trans 33_57 for year", i_process_year, ":", len(filtered_trans_33_57))
            
            
            for idx, dr in filtered_trans_33_57.iterrows():
                f_value = dr["c1"] + dr["c2"] + dr["c3"] + dr["c4"]
                if f_value != float(dr["c8"]):
                    numberRowsModified += 1
                    self.db_data.dt_trans.at[idx, "c8"] = f_value
                    self.o_db.adjust_u_trans_values(8, f_value, str(i_process_year), 
                                                  str(dr["rricc"]), str(dr["sch"]), str(dr["line"]))

                f_value = dr["c5"] + dr["c6"]
                if f_value != float(dr["c9"]):
                    numberRowsModified += 1
                    self.db_data.dt_trans.at[idx, "c9"] = f_value
                    self.o_db.adjust_u_trans_values(9, f_value, str(i_process_year), 
                                                  str(dr["rricc"]), str(dr["sch"]), str(dr["line"]))

                f_value = dr["c5"] + dr["c8"]
                if f_value != float(dr["c10"]):
                    numberRowsModified += 1
                    self.db_data.dt_trans.at[idx, "c10"] = f_value
                    self.o_db.adjust_u_trans_values(10, f_value, str(i_process_year), 
                                                  str(dr["rricc"]), str(dr["sch"]), str(dr["line"]))
            print(f"Number of rows modified for year {i_process_year}: {numberRowsModified}")
    
    def _get_value(self, s_year: str) -> float:
        d = 0.0
        i_rricc_code = 0

        if self.i_load_code in [1, 4, 5]:
            i_rricc_code = self.i_rricc_region
        elif self.i_load_code == 2:
            i_rricc_code = self.i_rricc_nation
        elif self.i_load_code == 3:
            i_rricc_code = self.i_rricc
        else:
            i_rricc_code = self.i_rricc

        filtered_trans = self.db_data.dt_trans[
            (self.db_data.dt_trans["year"] == int(s_year)) &
            (self.db_data.dt_trans["rricc"] == i_rricc_code) &
            (self.db_data.dt_trans["sch"] == self.i_sch) &
            (self.db_data.dt_trans["line"] == self.i_line)
        ]

        for _, dr_trans in filtered_trans.iterrows():
            column_name = f"c{self.i_col}"
            if column_name in dr_trans:
                d = self._scale(self.i_scale, self._return_decimal(dr_trans[column_name]))
            break

        # if (float(d) > 0):
        #     print(f"Value found for year {s_year}, rr_id {self.i_rr_id}, acode_id {self.i_acode_id}: {d}")
        return float(d)

    def _scale(self, scaler: int, value: float) -> float:
        if scaler == 1:
            return value / 10
        elif scaler == 2:
            return value / 100
        elif scaler == 3:
            return value / 1000
        elif scaler == 4:
            return value / 10000
        elif scaler == 5:
            return value / 100000
        elif scaler == 6:
            return value / 1000000
        else:
            return value

    def _return_decimal(self, value: Any) -> float:
        try:
            return float(value) if value is not None else 0.0
        except (ValueError, TypeError):
            return 0.0

    def _get_report_sheet(self, s_a_part: str, i_line: int) -> str:
        if s_a_part == "A1":
            if 101 <= i_line <= 160:
                return "A1P1"
            elif 201 <= i_line <= 216:
                return "A1P2A"
            elif 217 <= i_line <= 235:
                return "A1P2B"
            elif 236 <= i_line <= 254:
                return "A1P2C"
            elif 301 <= i_line <= 324:
                return "A1P3A"
            elif 341 <= i_line <= 364:
                return "A1P3B"
            elif 401 <= i_line <= 482:
                return "A1P4"
            elif 501 <= i_line <= 516:
                return "A1P5A"
            elif 521 <= i_line <= 536:
                return "A1P5B"
            elif 541 <= i_line <= 556:
                return "A1P6"
            elif 561 <= i_line <= 576:
                return "A1P7"
            elif 580 <= i_line <= 595:
                return "A1P8"
            elif 901 <= i_line <= 918:
                return "A1P9"
        elif s_a_part == "A2":
            if 101 <= i_line <= 184:
                return "A2P1"
            elif 201 <= i_line <= 262:
                return "A2P2"
            elif 301 <= i_line <= 366:
                return "A2P3"
            elif 401 <= i_line <= 422:
                return "A2P4"
        elif s_a_part == "A3":
            if 101 <= i_line <= 178:
                return "A3P1"
            elif 201 <= i_line <= 224:
                return "A3P2"
            elif 301 <= i_line <= 344:
                return "A3P3"
            elif 401 <= i_line <= 444:
                return "A3P4"
            elif 501 <= i_line <= 543:
                return "A3P5"
            elif 601 <= i_line <= 643:
                return "A3P6"
            elif 701 <= i_line <= 728:
                return "A3P7"
            elif 801 <= i_line <= 829:
                return "A3P8"
        elif s_a_part == "A4":
            if 101 <= i_line <= 145:
                return "A4P1"
            elif 170 <= i_line <= 178:
                return "A4P2"
            elif 201 <= i_line <= 205:
                return "A4P3"
        elif s_a_part == "E2":
            return "E2P1"
        
        return ""
    
if __name__ == "__main__":    
    
    phase2_input_data = Phase2InputData(current_year="2023")
    # phase2_input_data.temp_copy_to_redshift("ua_values_batch_a3ff198649ed41b1ae98e267c74c9c03.csv")
    success = phase2_input_data.prepare()
    if success:
        print("Phase 2 input data prepared successfully.")
    else:
        print("Error occurred while preparing Phase 2 input data.")
