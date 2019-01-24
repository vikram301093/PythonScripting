#!/usr/bin/env python

import os

#print(os.getcwd())

list_dir = os.listdir("/home/ec2-user/Python_Tutorial")
for x in list_dir:
 if x == "os":
   list_file = os.listdir("/home/ec2-user/Python_Tutorial/os")
   os.chdir("/home/ec2-user/Python_Tutorial/os")
   for x in list_file:
     os.remove(x)
   os.rmdir("/home/ec2-user/Python_Tutorial/os")

os.mkdir("/home/ec2-user/Python_Tutorial/os")
os.mknod("/home/ec2-user/Python_Tutorial/os/first_python.py",755)
os.mknod("/home/ec2-user/Python_Tutorial/os/second_python.py",755)

