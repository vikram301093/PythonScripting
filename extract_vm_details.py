#!/usr/bin/python

import re
import csv
import copy
import sys, os
import time
import string
import email, smtplib, ssl
import adal
import xlsxwriter
import base64
import pandas as pd
import logging

from azure.mgmt.compute import ComputeManagementClient
from azure.mgmt.resource import ResourceManagementClient, SubscriptionClient
from msrestazure.azure_active_directory import AdalAuthentication
#from azure.common.credentials import ServicePrincipalCredentials
from os.path import basename
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import COMMASPACE, formatdate
from email.mime.base import MIMEBase
from datetime import datetime
from ansible.constants import DEFAULT_VAULT_ID_MATCH
from ansible.parsing.vault import VaultLib
from ansible.parsing.vault import VaultSecret

#Create and configure logger 
logging.basicConfig(filename="/root/vm_count_script/vm_count.log",
                    datefmt='%d-%m-%Y %H:%M:%S',
                    format='%(asctime)s %(levelname)s %(message)s', 
                    filemode='w') 
  
#Creating an object 
logger=logging.getLogger() 
  
#Setting the threshold of logger to DEBUG 
logger.setLevel(logging.INFO)

# Creating date variable
date = time.strftime("%d-%m-%Y")

# Using ansible vault for our credentials
vault_data = {}
vault = VaultLib([(DEFAULT_VAULT_ID_MATCH, VaultSecret(base64.b64decode('UHl0aG9uQDEyMw==')))])
vault_data = eval(vault.decrypt(open('/root/vm_count_script/consumption_vault').read()))

# Service Principal for ECE & AMAP Azure
AZURE_CLIENT_ID = vault_data['AZURE_CLIENT_ID']
AZURE_SECRET = vault_data['AZURE_SECRET']
AZURE_TENANT = vault_data['AZURE_TENANT']

# End points for ECE & AMAP Azure
authentication_endpoint = 'https://login.microsoftonline.com/'
azure_endpoint = 'https://management.azure.com/'

# Service Principal for CN
AZURE_CLIENT_ID_CN = vault_data['AZURE_CLIENT_ID_CN']
AZURE_SECRET_CN = vault_data['AZURE_SECRET_CN']
AZURE_TENANT_CN = vault_data['AZURE_TENANT_CN']

# End points for CN Azure
authentication_endpoint_CN = 'https://login.chinacloudapi.cn/'
azure_endpoint_CN = 'https://management.chinacloudapi.cn/'

# Block for ECE & AMAP Azure function

# Fetches subscription id
def getSubsIDs(cred, base_url=azure_endpoint):
    logger.info("Inside subscription id fetching function")
    subscription = SubscriptionClient(cred, base_url=azure_endpoint)
    subs_dict = {}
    for subs in subscription.subscriptions.list():
        subs_dict[subs.display_name] = subs.subscription_id
    return subs_dict

# Fetches vmss details for a perticular subscription id
def getVMSSdetails(cred, subscription_id):
    logger.info("Inside vmss details fetching function")
    rg_client = ResourceManagementClient(cred, subscription_id, base_url=azure_endpoint)
    rglist = []
    for rg in rg_client.resource_groups.list():
        rglist.append(rg.name)
    return rglist

# Get the list of all subscription id for the customer which was passed in the argument
def getCustEnv(subs_dict):
    logger.info("Inside customer env fetching function")
    cust_subs = {}
    #envlist = ['DEVDAP','DEVDAP2','BCCSSCONFP','VESTAP']
    for sub in subs_dict.keys():
        for env in envlist:
            z = re.match(".*" + env + "$", sub)
            if z:
                cust_subs[sub] = subs_dict[sub]
    return cust_subs

# Get core count for the vm
def getCoreCount(cred, subscription_id, skusize, location):
    logger.info("Inside counting core count function")
    core_client = ComputeManagementClient(cred, subscription_id, base_url=azure_endpoint)
    vm_cores_list = core_client.virtual_machine_sizes.list(location)
    core_count = 0
    #print(skusize)
    for i in vm_cores_list:
        if i.name == skusize:
            core_count += i.number_of_cores
    return core_count

# Block for CN Azure function

# Fetches subscription id
def getSubsIDs_CN(cred_CN, base_url=azure_endpoint_CN):
    logger.info("Inside subscription id fetching function")
    subscription_CN = SubscriptionClient(cred_CN, base_url=azure_endpoint_CN)
    subs_dict_CN = {}
    for subs in subscription_CN.subscriptions.list():
        subs_dict_CN[subs.display_name] = subs.subscription_id
    return subs_dict_CN

# Fetches vmss details for a perticular subscription id
def getVMSSdetails_CN(cred_CN, subscription_id_CN):
    logger.info("Inside vmss details fetching function")
    rg_client_CN = ResourceManagementClient(cred_CN, subscription_id_CN, base_url=azure_endpoint_CN)
    rglist_CN = []
    for rg in rg_client_CN.resource_groups.list():
        rglist_CN.append(rg.name)
    return rglist_CN

# Get the list of all subscription id for the customer which was passed in the argument
def getCustEnv_CN(subs_dict_CN):
    logger.info("Inside customer env fetching function")
    cust_subs_CN = {}
    #envlist = ['DEVDAP','DEVDAP2','BCCSSCONFP','VESTAP']
    for sub in subs_dict_CN.keys():
        for env in envlist:
            z = re.match(".*" + env + "$", sub)
            if z:
                cust_subs_CN[sub] = subs_dict_CN[sub]
    return cust_subs_CN

# Get core count for the vm
def getCoreCount_CN(cred_CN, subscription_id_CN, skusize, location):
    logger.info("Inside counting core count function")
    core_client_CN = ComputeManagementClient(cred_CN, subscription_id_CN, base_url=azure_endpoint_CN)
    vm_cores_list_CN = core_client_CN.virtual_machine_sizes.list(location)
    core_client_CN = 0
    #print(skusize)
    for i in vm_cores_list_CN:
        if i.name == skusize:
            core_client_CN += i.number_of_cores
    return core_client_CN

# Common Block function which are needed by all the regions

# Function for cleaning up of old data files
def cleanup():
    logger.info("Inside cleanup function")
    path = "/root/vm_count_script/data"
    now = time.time()
    for f in os.listdir(path):
        fpath = os.path.join(path, f)
        if os.stat(fpath).st_mtime < now - 15 * 86400:
            if os.path.isfile(fpath):
                os.remove(fpath)

# Function for sending data file over email
def sendanemail(datafile):
    logger.info("Inside send email function")
    SUBJECT = "VM and Core count data for the day."
    CC = "ARPITA.MARATHE@t-systems.com"

    #TO = "mayur.naik@t-systems.com"
    TO = "DL-DL-ConfPlus_MCSPaaS@t-systems.com"
    #TO = "parthik.ghosh@t-systems.com"
    
    FROM = "no-reply@mcs-paas.io"
    text = """Hello Team,
                        
                        Please find attached VM and core counts
    
Thanks!!"""

    msg = MIMEMultipart()
    msg['From'] = FROM
    msg['To'] = TO
    msg['Cc'] = CC
    msg['Date'] = formatdate(localtime=True)
    msg['Subject'] = SUBJECT

    msg.attach(MIMEText(text))

    with open(datafile, "rb") as fil:
        part = MIMEApplication(
            fil.read(),
            Name=basename(datafile)
        )

    part['Content-Disposition'] = 'attachment; filename="%s"' % basename(datafile)
    msg.attach(part)
    server = smtplib.SMTP('mailin12.telekom.de:25')
    server.sendmail(FROM, TO, msg.as_string())

# Calling clean up function to clean old files
cleanup()

# Passing customer name to env list from the arguments
envlist = sys.argv[1:]

# Logic block for ECE & AMAP Azure region which creates the data file
logger.info("Inside Logic block for ECE & AMAP Azure region")

# Authentication block
context = adal.AuthenticationContext(authentication_endpoint+AZURE_TENANT)
credentials = AdalAuthentication(
    context.acquire_token_with_client_credentials,
    azure_endpoint,
    AZURE_CLIENT_ID,
    AZURE_SECRET
)

# Subs_dict -- dictionary would contain the all the subscription names and IDs
subs_dict = getSubsIDs(credentials)

# Cust_subs -- dictionary would contain the all names and IDs of provided customers (as arguments)
cust_subs = getCustEnv(subs_dict)

# Cust_subs_nest -- dictionary would contain subscription name as key and list of subscrID and list of resource groups in subscription
cust_subs_nest = {}

# Convert the value to list
for i in cust_subs:
    cust_subs_nest[i] = cust_subs[i].split()

# Add the resource groups to the value list
id_rg_dict = {}
for key, value in cust_subs_nest.items():
    # Get the list of resource groups
    rglist = getVMSSdetails(credentials, value[0])
    cust_subs_nest[key].extend(rglist)

# Id_rg_dict -- -- dictionary would contain subscription and list of subsID, resourceg
id_rg_dict = cust_subs_nest

# Alter id_rg_dict so that value will contain only subsID and resourcegroup containing the VM/VMSS
for sub in id_rg_dict:
    for env in envlist:
        for rg in id_rg_dict[sub][1:]:
            z = re.match(".*" + env + "$", rg)
            if z:
                id_rg_dict[sub][1:] = []
                id_rg_dict[sub][1:] = rg.split()

# Assign dummy value if no customer resource group present
for i in id_rg_dict:
    if len(id_rg_dict[i]) <= 1:
        id_rg_dict[i].append('NA')

id_rg_dict_core = copy.deepcopy(id_rg_dict)

print('---ECE & AMAP------------------------------------------------------------------------------------------------------')

# Prints subscription name & id
for i in id_rg_dict:
    print 'id_rg_dict', i, id_rg_dict[i]

print('---ECE & AMAP------------------------------------------------------------------------------------------------------')

vmss_type = ['master','worker','infra']

# Logic for fetching vm count in a vmss and also counting the number of cores
logger.info("Inside Logic for fetching vm count in a vmss and also counting the number of cores")
for i in id_rg_dict:
    count_total = 0
    total_core = 0
    count_vm = 0
    vm_client = ComputeManagementClient(credentials, id_rg_dict[i][0], base_url=azure_endpoint)
    for j in vmss_type:
        compute_client = ComputeManagementClient(credentials, id_rg_dict[i][0], base_url=azure_endpoint)
        vmss = compute_client.virtual_machine_scale_set_vms.list(id_rg_dict[i][1], j)
        vm = compute_client.virtual_machines.list(id_rg_dict[i][1], j)
        count = 0
        core_per_vmss = 0
        total_vmss_core = 0
        total_core_per_vmss = 0
        try:
            for k in vmss:
                count += 1
                print(i, k.name, k.sku.name, k.location)
                core_counts = getCoreCount(credentials, id_rg_dict[i][0], k.sku.name, k.location)
                # print(core_counts)
                total_vmss_core += core_counts
                total_core_per_vmss += core_counts
            id_rg_dict[i].append(count)
            core_per_vmss = total_core_per_vmss/count
            total_core_per_vmss = 0
            id_rg_dict[i].append(core_per_vmss)
            count_total += count
        except:
            logger.error("There is some error while fetching vm count in a vmss and core count in a resorce group in ECE & AMAP region") 
            id_rg_dict[i].append(0)
            count_total += 0
        total_core += total_vmss_core
        #id_rg_dict_core[i].append(total_vmss_core)
    vm = vm_client.virtual_machines.list(id_rg_dict[i][1])
    try:
        for name in vm:
            count_vm += 1
        id_rg_dict[i].append(count_vm)
        count_total+=count_vm
    except:
        logger.error("There is some error while fetching vm count in a resorce group in ECE & AMAP region")
        id_rg_dict[i].append(0)

    print(i, total_core)
    print('---ECE & AMAP------------------------------------------------------------------------------------------------------')
    id_rg_dict[i].append(count_total)
    id_rg_dict_core[i].append(total_core)

print('---ECE & AMAP------------------------------------------------------------------------------------------------------')

# Logic for compiling all data into one data stucture
logger.info("Inside Logic for compiling all data into one data stucture")

ds = [id_rg_dict, id_rg_dict_core]
dict_total = {}

for k in id_rg_dict.iterkeys():
    dict_total[k] = list(dict_total[k][2:] for dict_total in ds)

for i in dict_total:
    dict_total[i] =  dict_total[i][0] + dict_total[i][1]

for key, value in dict_total.iteritems():
    value.insert(0, key)
    value.insert(0, date)
    print(value)

print('---ECE & AMAP------------------------------------------------------------------------------------------------------')

# This is block is needed to format the output table
dict_total_excel = {}
for k in id_rg_dict.iterkeys():
    dict_total_excel[k] = list(dict_total_excel[k][2:] for dict_total_excel in ds)

for i in dict_total_excel:
    dict_total_excel[i] =  dict_total_excel[i][0] + dict_total_excel[i][1]

for key, value in dict_total_excel.iteritems():
    value.insert(0, key)
    value.insert(0, date)
    print(value)

print('---ECE & AMAP-Excel-----------------------------------------------------------------------------------------------------')

excel_values = {}
excel_values = dict_total

for key, value in excel_values.iteritems():
    # Removing values to create a new order sequence which match manual excel sheet
    del excel_values[key][3]
    del excel_values[key][4]
    del excel_values[key][5]
    del excel_values[key][7]
    print(value)

print('---ECE & AMAP-Excel-----------------------------------------------------------------------------------------------------')

excel_values_new = {}
excel_values_new = dict_total_excel

logger.info("Inside Logic which gives proper structured data which can be put in excel for ECE & AMAP region")

for (key1, value1), (key2, value2) in zip(excel_values.iteritems(), excel_values_new.iteritems()):
    # Appending values to create a new order sequence which match manual excel sheet
    excel_values[key1].append(excel_values_new[key2].pop(3))
    excel_values[key1].append(excel_values_new[key2].pop(4))
    excel_values[key1].append(excel_values_new[key2].pop(5))
    excel_values[key1].append(excel_values_new[key2].pop(7))
    print(value1)

# The output of the above logics give us a proper structured data which can be put in excel
print('---ECE & AMAP-Excel-----------------------------------------------------------------------------------------------------')

# Logic block for CN Azure region which creates the data file
logger.info("Inside Logic block for CN Azure region")

# Authentication block
context_CN = adal.AuthenticationContext(authentication_endpoint_CN+AZURE_TENANT_CN)

credentials_CN = AdalAuthentication(
    context_CN.acquire_token_with_client_credentials,
    azure_endpoint_CN,
    AZURE_CLIENT_ID_CN,
    AZURE_SECRET_CN
)

# Subs_dict_CN -- dictionary would contain the all the subscription_CN names and IDs
subs_dict_CN = getSubsIDs_CN(credentials_CN)
# Cust_subs_CN -- dictionary would contain the all names and IDs of provided customers (as arguments)
cust_subs_CN = getCustEnv_CN(subs_dict_CN)
# Cust_subs_nest_CN -- dictionary would contain subscription_CN name as key and list of subscrID and list of resource groups in subscription_CN
cust_subs_nest_CN = {}

# Convert the value to list
for i in cust_subs_CN:
    cust_subs_nest_CN[i] = cust_subs_CN[i].split()

# Add the resource groups to the value list
id_rg_dict_CN = {}
for key, value in cust_subs_nest_CN.items():
    # Get the list of resource groups
    rglist_CN = getVMSSdetails_CN(credentials_CN, value[0])
    cust_subs_nest_CN[key].extend(rglist_CN)

# Id_rg_dict_CN -- -- dictionary would contain subscription_CN and list of subsID, resourceg
id_rg_dict_CN = cust_subs_nest_CN

# Alter id_rg_dict_CN so that value will contain only subsID and resourcegroup containing the VM/VMSS
for sub in id_rg_dict_CN:
    for env in envlist:
        for rg in id_rg_dict_CN[sub][1:]:
            z = re.match(".*" + env + "$", rg)
            if z:
                id_rg_dict_CN[sub][1:] = []
                id_rg_dict_CN[sub][1:] = rg.split()

# Assign dummy value if no customer resource group present
for i in id_rg_dict_CN:
    if len(id_rg_dict_CN[i]) <= 1:
        id_rg_dict_CN[i].append('NA')

id_rg_dict_core_CN = copy.deepcopy(id_rg_dict_CN)

print('----CN-----------------------------------------------------------------------------------------------------')

# Prints subscription name & id
for i in id_rg_dict_CN:
    print 'id_rg_dict_CN', i, id_rg_dict_CN[i]

print('----CN-----------------------------------------------------------------------------------------------------')

vmss_type_CN = ['master','worker','infra']

# Logic for fetching vm count in a vmss and also counting the number of cores
logger.info("Inside Logic for fetching vm count in a vmss and also counting the number of cores")
for i in id_rg_dict_CN:
    count_total_CN = 0
    total_core_CN = 0
    count_vm_CN = 0
    vm_client_CN = ComputeManagementClient(credentials_CN, id_rg_dict_CN[i][0], base_url=azure_endpoint_CN)
    for j in vmss_type_CN:
        compute_client_CN = ComputeManagementClient(credentials_CN, id_rg_dict_CN[i][0], base_url=azure_endpoint_CN)
        vmss_CN = compute_client_CN.virtual_machine_scale_set_vms.list(id_rg_dict_CN[i][1], j)
        vm_CN = compute_client_CN.virtual_machines.list(id_rg_dict_CN[i][1], j)
        count_CN = 0
        core_per_vmss_CN = 0
        total_vmss_core_CN = 0
        total_core_per_vmss_CN = 0
        try:
            for k in vmss_CN:
                count_CN += 1
                print(i, k.name, k.sku.name, k.location)
                core_counts_CN = getCoreCount_CN(credentials_CN, id_rg_dict_CN[i][0], k.sku.name, k.location)
                #print(core_counts_CN)
                total_vmss_core_CN += core_counts_CN
                total_core_per_vmss_CN += core_counts_CN
            id_rg_dict_CN[i].append(count_CN)
            core_per_vmss_CN = total_core_per_vmss_CN/count_CN
            total_core_per_vmss_CN = 0
            id_rg_dict_CN[i].append(core_per_vmss_CN)
            count_total_CN += count_CN
        except:
            logger.error("There is some error while fetching vm count in a vmss and core count in a resorce group in CN region") 
            id_rg_dict_CN[i].append(0)
            count_total_CN += 0
        total_core_CN += total_vmss_core_CN
        #id_rg_dict_core_CN[i].append(total_vmss_core_CN)

    vm_CN = vm_client_CN.virtual_machines.list(id_rg_dict_CN[i][1])
    try:
        for name in vm_CN:
            count_vm_CN += 1
        id_rg_dict_CN[i].append(count_vm_CN)
        count_total_CN+=count_vm_CN
    except:
        logger.error("There is some error while fetching vm count in a resorce group in CN region")
        id_rg_dict_CN[i].append(0)

    print(i, total_core_CN)
    print('----CN-----------------------------------------------------------------------------------------------------')
    id_rg_dict_CN[i].append(count_total_CN)
    id_rg_dict_core_CN[i].append(total_core_CN)

print('----CN-----------------------------------------------------------------------------------------------------')

# Logic for compiling all data into one data stucture
logger.info("Inside Logic for compiling all data into one data stucture")

ds_CN = [id_rg_dict_CN, id_rg_dict_core_CN]
dict_total_CN = {}

for k in id_rg_dict_CN.iterkeys():
    dict_total_CN[k] = list(dict_total_CN[k][2:] for dict_total_CN in ds_CN)

for i in dict_total_CN:
    dict_total_CN[i] =  dict_total_CN[i][0] + dict_total_CN[i][1]

for key, value in dict_total_CN.iteritems():
    value.insert(0, key)
    value.insert(0, date)
    print(value)

print('----CN-----------------------------------------------------------------------------------------------------')

#This is block is needed to format the output table
dict_total_excel_CN = {}
for k in id_rg_dict_CN.iterkeys():
    dict_total_excel_CN[k] = list(dict_total_excel_CN[k][2:] for dict_total_excel_CN in ds_CN)

for i in dict_total_excel_CN:
    dict_total_excel_CN[i] =  dict_total_excel_CN[i][0] + dict_total_excel_CN[i][1]

for key, value in dict_total_excel_CN.iteritems():
    value.insert(0, key)
    value.insert(0, date)
    print(value)

print('----CN-Excel----------------------------------------------------------------------------------------------------')

excel_values_CN = {}
excel_values_CN = dict_total_CN

for key, value in excel_values_CN.iteritems():
    # Removing values to create a new order sequence which match manual excel sheet
    del excel_values_CN[key][3]
    del excel_values_CN[key][4]
    del excel_values_CN[key][5]
    del excel_values_CN[key][7]
    print(value)

print('----CN-Excel----------------------------------------------------------------------------------------------------')

excel_values_new_CN = {}
excel_values_new_CN = dict_total_excel_CN

logger.info("Inside Logic which gives proper structured data which can be put in excel for CN region")

for (key1, value1), (key2, value2) in zip(excel_values_CN.iteritems(), excel_values_new_CN.iteritems()):
    # Appending values to create a new order sequence which match manual excel sheet
    excel_values_CN[key1].append(excel_values_new_CN[key2].pop(3))
    excel_values_CN[key1].append(excel_values_new_CN[key2].pop(4))
    excel_values_CN[key1].append(excel_values_new_CN[key2].pop(5))
    excel_values_CN[key1].append(excel_values_new_CN[key2].pop(7))
    print(value1)

# The output of the above logics give us a proper structured data which can be put in excel
print('----CN-Excel----------------------------------------------------------------------------------------------------')

print('----ECE, AMAP & CN-Excel----------------------------------------------------------------------------------------------------')

# In this block we join the result from all the refions ECE, AMAP & CN
excel_values.update(excel_values_CN)

#from itertools import chain
#sorted_dict = dict(chain(excel_values.items(), excel_values_CN.items()))
for key, value in excel_values.iteritems():
    print(value)

print('----ECE, AMAP & CN-Excel----------------------------------------------------------------------------------------------------')

# Sorting the excel sheet according to subscription name
from collections import OrderedDict
sorted_dict = OrderedDict(sorted(excel_values.items(), key=lambda x: x[1]))

for key, value in sorted_dict.iteritems():
    print(value)

print('----ECE, AMAP & CN-Excel----------------------------------------------------------------------------------------------------')

# Block for creating a excel file and writing and formating all the data we gathered
datafile = '/root/vm_count_script/data/' + 'data_' + time.strftime("%d-%m-%Y") + '.xlsx'

# Creating excel workbook and worksheet
workbook = xlsxwriter.Workbook(datafile)
worksheet = workbook.add_worksheet()

row = 0
col = 0

# Setting proper width to columns
worksheet.set_column(0, 0, 15)
worksheet.set_column(1, 1, 50)
worksheet.set_column(2, 10, 15)

# Setting format for header
header_format = workbook.add_format()
header_format.set_bg_color('#FFC000')
header_format.set_border() 
header_format.set_align('center')

# Setting format for column
column_format = workbook.add_format()
column_format.set_bg_color('#FFFF00')
column_format.set_border()

# Setting format for row
row_format = workbook.add_format()
row_format.set_border()
row_format.set_align('center')

# Setting format for output
output_format = workbook.add_format()
output_format.set_bg_color('#00B050')
output_format.set_align('center')
output_format.set_border()

# Creating Header
header =['Date', 'Subcription', 'Master Nodes', 'Worker Nodes', 'Infra Nodes', ' Postgress Nodes', 'VM Total', 'Core per Master', 'Core per Worker', ' Core per Infra', 'Core Total']
worksheet.write_row('A1', header, header_format)

row += 1

# Inserting values
for key, value in sorted_dict.iteritems():
    worksheet.write_row(row, col, value, row_format)
    worksheet.write(row, 1, key, column_format)
    worksheet.write(row, 6, value.pop(6), output_format)
    worksheet.write(row, 10, value.pop(), output_format)
    row += 1

# Closing the created workbook
workbook.close()

#Sending excel sheet over email for a perticular date
sendanemail(datafile)
print('----ECE, AMAP & CN-Excel----------------------------------------------------------------------------------------------------')


print('----Juno_Core_VM_Amount-Excel----------------------------------------------------------------------------------------------------')

# In this block we add the current date excel data to the excel file which has previous date data
main_excel = "/root/vm_count_script/Juno_Core_VM_Amount.xlsx"

excel_day = pd.read_excel('/root/vm_count_script/data/' + 'data_' + time.strftime("%d-%m-%Y") + '.xlsx')
excel_total = pd.read_excel("/root/vm_count_script/Juno_Core_VM_Amount.xlsx")

all_excel = [excel_total, excel_day]

appended_df = pd.concat(all_excel, sort=False)

appended_df.to_excel("/root/vm_count_script/Juno_Core_VM_Amount.xlsx", index=False, sheet_name='juno_data')

#sendanemail(main_excel)
print('----Juno_Core_VM_Amount-Excel----------------------------------------------------------------------------------------------------')

# Formating the final excel sheet
def xlsx_to_workbook(xlsx_in_file_url, xlsx_out_file_url, sheetname):
    logger.info("Inside function for creating final workbook excel")
    """
    Read EXCEL file into xlsxwriter workbook worksheet
    """
    workbook = xlsxwriter.Workbook(xlsx_out_file_url)
    worksheet = workbook.add_worksheet(sheetname)
    # Read my_excel into a pandas DataFrame
    df = pd.read_excel(xlsx_in_file_url)
    # A list of column headers
    list_of_columns = df.columns.values

    for col in range(len(list_of_columns)):
        # Write column headers.
        # If you don't have column headers remove the folling line and use "row" rather than "row+1" in the if/else statments below
        worksheet.write(0, col, list_of_columns[col] )
        for row in range (len(df)):
            # Test for Nan, otherwise worksheet.write throws it.
            if df[list_of_columns[col]][row] != df[list_of_columns[col]][row]:
                worksheet.write(row+1, col, "")
            else:
                worksheet.write(row+1, col, df[list_of_columns[col]][row])
    return workbook, worksheet

# Create a workbook
# Read you Excel file into a workbook/worksheet object to be manipulated with xlsxwriter
# This assumes that the EXCEL file has column headers
workbook_new, worksheet_new = xlsx_to_workbook("/root/vm_count_script/Juno_Core_VM_Amount.xlsx", "/root/vm_count_script/Juno_Core_VM_Amount.xlsx", "juno_data")

# Setting format for header
header_format_new = workbook_new.add_format()
header_format_new.set_bg_color('#FFC000')
header_format_new.set_border()
header_format_new.set_align('center')

# Setting format for column
column_format_new = workbook_new.add_format()
column_format_new.set_border()
column_format_new.set_align('center')

# Setting format for key
key_format_new = workbook_new.add_format()
key_format_new.set_border()
key_format_new.set_align('center')
key_format_new.set_bg_color('#FFFF00')

# Setting format for output
output_format_new = workbook_new.add_format()
output_format_new.set_bg_color('#00B050')
output_format_new.set_align('center')
output_format_new.set_border()

# Setting proper format to columns
worksheet_new.set_column(0, 0, 15, column_format_new)
worksheet_new.set_column(1, 1, 50, key_format_new)
worksheet_new.set_column(2, 10, 15, column_format_new)
worksheet_new.set_column(6, 6, 15, output_format_new)
worksheet_new.set_column(10, 10, 15, output_format_new)

# Setting proper format to heading
worksheet_new.set_row(0, None, header_format_new)

workbook_new.close()

# Sending the final excel which contains all dates data.
sendanemail(main_excel)
logger.info("Python script was able to send email for data consumption")
print('----Juno_Core_VM_Amount-Excel----------------------------------------------------------------------------------------------------')

