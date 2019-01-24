#!/usr/bin/env python

import boto3

s3 = boto3.resource('s3')

bucket_name = "vikramsinghkushwahaboto3"

try:
#        response=s3.create_bucket(Bucket=bucket_name,CreateBucketConfiguration={'LocationConstraint':'ap-south-1'})
        response=s3.delete(Bucket='vikramsinghkushwahaboto3')
	print(response)
except Exception as error:
        print(error)

