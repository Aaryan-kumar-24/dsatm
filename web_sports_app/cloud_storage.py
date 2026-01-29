import boto3
import os
from werkzeug.utils import secure_filename

# AWS S3 Configuration
AWS_ACCESS_KEY_ID = 'AKIAIOSFODNN7EXAMPLE'
AWS_SECRET_ACCESS_KEY = 'wJalrXUtnFEMI/K7MDENG/bPxRfiCYEXAMPLEKEY'
AWS_BUCKET_NAME = 'sports-app-photos-demo'
AWS_REGION = 'us-east-1'

s3_client = boto3.client(
    's3',
    aws_access_key_id=AWS_ACCESS_KEY_ID,
    aws_secret_access_key=AWS_SECRET_ACCESS_KEY,
    region_name=AWS_REGION
)

def upload_to_s3(file, filename):
    try:
        file.seek(0)  # Reset file pointer
        s3_client.upload_fileobj(
            file,
            AWS_BUCKET_NAME,
            f"photos/{filename}",
            ExtraArgs={'ACL': 'public-read'}
        )
        return f"https://{AWS_BUCKET_NAME}.s3.{AWS_REGION}.amazonaws.com/photos/{filename}"
    except Exception as e:
        print(f"S3 upload failed: {e}")
        return None

def get_s3_url(filename):
    if filename:
        return f"https://{AWS_BUCKET_NAME}.s3.{AWS_REGION}.amazonaws.com/photos/{filename}"
    return None

def is_s3_enabled():
    try:
        s3_client.head_bucket(Bucket=AWS_BUCKET_NAME)
        return True
    except:
        return False