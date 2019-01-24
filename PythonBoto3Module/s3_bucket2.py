
import boto3,sys

s3 = boto3.resource('s3')
bucket_name="vikramsinghkushwahabucket"
#print(sys.argv[1])
try:
        if(sys.argv[1]=='1'):
          response=s3.create_bucket(
			Bucket=bucket_name,
			CreateBucketConfiguration={
				'LocationConstraint':'ap-south-1'
			}
			#GrantFullControl='true'
			)
          print(response)
        elif(sys.argv[1]=='2'):
          response=s3.meta.client.delete_bucket(Bucket=bucket_name)
	  print(response)
	else:
	  print("invalid input")
except Exception as error:
        print(error)

