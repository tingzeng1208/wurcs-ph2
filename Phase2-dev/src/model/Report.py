import json
import boto3

class ReportSingleton:
    _instance = None
    _initialized = False

    def __new__(cls, *args, **kwargs):
        if not cls._instance:
            cls._instance = super().__new__(cls)
        return cls._instance

    def __init__(self):
        if not self._initialized:
            self.report_data = {}
            self._initialized = True

    def add_report_data(self, key, value):
        self.report_data[key] = value

    def get_report_data(self, key):
        return self.report_data.get(key)

    def clear_report_data(self):
        self.report_data.clear()

    def serialize_to_json(self, file_path):
        """
        Serializes the report_data to a JSON file at the given path.
        """
        with open(file_path, 'w', encoding='utf-8') as f:
            json.dump(self.report_data, f, ensure_ascii=False, indent=4)

    def upload_json_to_s3(self, file_path, bucket, key):
        """
        Uploads the JSON file at file_path to the specified S3 bucket with the given key.
        """
        s3 = boto3.client('s3')
        s3.upload_file(file_path, bucket, key)

    def serialize_and_upload_to_s3(self, bucket, key):
        """
        Serializes report_data to a JSON string and uploads it directly to S3.
        """
        s3 = boto3.client('s3')
        json_data = json.dumps(self.report_data, ensure_ascii=False, indent=4)
        s3.put_object(Bucket=bucket, Key=key, Body=json_data.encode('utf-8'))

    def deserialize_from_json(self, file_path):
        """
        Loads report_data from a JSON file at the given path.
        """
        with open(file_path, 'r', encoding='utf-8') as f:
            self.report_data = json.load(f)

    def download_json_from_s3(self, bucket, key):
        """
        Downloads a JSON file from S3 and loads report_data from it.
        """
        s3 = boto3.client('s3')
        response = s3.get_object(Bucket=bucket, Key=key)
        json_data = response['Body'].read().decode('utf-8')
        self.report_data = json.loads(json_data)