#!/usr/bin/env python3
"""Test script to verify cloud storage setup"""

from cloud_storage import is_s3_enabled, upload_to_s3, get_s3_url
import io

def test_cloud_storage():
    print("Testing cloud storage setup...")
    
    # Test S3 connection
    if is_s3_enabled():
        print("[OK] S3 connection successful")
        
        # Test file upload
        test_file = io.BytesIO(b"test image data")
        test_filename = "test_photo.jpg"
        
        url = upload_to_s3(test_file, test_filename)
        if url:
            print(f"[OK] File upload successful: {url}")
        else:
            print("[FAIL] File upload failed")
            
    else:
        print("[INFO] S3 connection failed - using local storage fallback")
        print("Note: This is normal if AWS credentials are not configured")
        print("Note: This is normal if AWS credentials are not configured")

if __name__ == "__main__":
    test_cloud_storage()