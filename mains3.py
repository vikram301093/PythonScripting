
import boto3,s3_module

s3 = boto3.resource('s3')

#bucket_response = s3_module.s3_create_bucket('vikramsinghkushwahas3',s3)

bucket_response = s3_module.s3_delete_bucket('vikramsinghkushwahas3',s3)

print(bucket_response)

