
import boto3

def s3_create_bucket(bucket_name,s3):
    session = boto3.session.Session()
    current_region = session.region_name
    bucket_response = s3.create_bucket(
			Bucket=bucket_name,
			CreateBucketConfiguration={
				'LocationConstraint':current_region
			}
		     )
    print(bucket_response)
    return bucket_name, bucket_response

def s3_delete_bucket(bucket_name,s3):
    session = boto3.session.Session()
    current_region = session.region_name
    bucket_response = s3.meta.client.delete_bucket(Bucket=bucket_name)
   # print(bucket_response)
    return bucket_name, bucket_response

def s3_upload_file(bucket_name,s3,filename)
    session = boto3.session.Session()
    current_region = session.region_name
    
