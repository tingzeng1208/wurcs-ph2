import boto3
from botocore.exceptions import BotoCoreError, NoCredentialsError

def upload_file_to_s3(file_path, bucket_name, s3_key):
    """
    Uploads a file to S3 using the default AWS profile.
    
    :param file_path: Local path to the file.
    :param bucket_name: S3 bucket name.
    :param s3_key: S3 object key (path inside the bucket).
    """
    try:
        # Use default profile (from ~/.aws/credentials)
        session = boto3.Session(profile_name='default')
        s3_client = session.client('s3')
        
        s3_client.upload_file(file_path, bucket_name, s3_key)
        print(f"✅ Upload successful: s3://{bucket_name}/{s3_key}")

    except NoCredentialsError:
        print("❌ No AWS credentials found. Have you run `aws configure`?")
    except BotoCoreError as e:
        print(f"❌ BotoCoreError: {e}")
    except Exception as e:
        print(f"❌ Unexpected error: {e}")
