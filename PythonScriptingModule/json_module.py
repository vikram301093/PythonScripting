#!/usr/bin/env python

import json

employee_data = '''
{
"people":[
{
"name":"vikram",
"email":"vikramsinghkushwaha8@gmail.com",
"mobile_no":"9028544625"
},
{
"name":"nishtha",
"email":"nishthagoel4sep@gmail.com",
"mobile_no":"9028544625"
}
]
}
'''

print(employee_data)
data = json.loads(employee_data)
print(data)
