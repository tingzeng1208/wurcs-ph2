from services.DBManager import DBManager

def copyS3ToRedShift(s3_bucket:str, s3_key: str, db_manager: DBManager, redshift_table: str, aws_access_key_id: str, aws_secret_access_key: str, s_current_year: str):
        
        try:
            
            # s3_bucket = os.getenv("S3_BUCKET")  # TODO: set your bucket name
            # s3_key = f"ua_values_batches/{csv_filename}"
            # csv_path = os.path.join(os.getenv("LOCAL_CSV_PATH"), csv_filename)
            # upload_file_to_s3(csv_path, s3_bucket, s3_key)
            # print(f"Uploaded {csv_path} to s3://{s3_bucket}/{s3_key}")
            
            # iam_role = os.getenv("IAM_ROLE_ARN")
            
            databasename = db_manager._get_database_name_from_sql(s_current_year, "AVALUES")
            
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
            db_manager.execute_non_query(copy_sql, None)
            print(f"Copied data from s3://{s3_bucket}/{s3_key} into Redshift table {redshift_table}")
         
        except Exception as e:
            print(f"An error occurred during upload to Redshift: {e}")
            raise e   
    