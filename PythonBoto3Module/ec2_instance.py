#!/usr/bin/env python

import boto3

ec2 = boto3.resource('ec2')
'''
instance = ec2.create_instances(
	ImageId='ami-0ad42f4f66f6c1cc9',
 	MinCount=1,
	MaxCount=1,
	InstanceType='t2.micro'
)

print(instance[0].id)
'''

instance_id = 'i-0f77317f515dcbed5'
instance = ec2.Instance(instance_id)
response = instance.terminate()
print(response)

