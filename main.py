import pandas as pd
import json
import requests
import base64
import win32com.client as win32

import os
from datetime import datetime

datetime_Format = "{:%Y-%m-%d}".format(datetime.now())

newpath = "C:/path/to/report/" + "Updated_Report_" + datetime_Format
if not os.path.exists(newpath):
    os.makedirs(newpath)

with open('C:/path/to/textfile/ivnt_encode.txt', 'rb') as ivnt_key:
  coded_string_ivnt = ivnt_key.read()
ivnt_api__Key = base64.b64decode(coded_string_ivnt).decode('utf-8')
iv_auth_header = {'Authorization': ivnt_api__Key}

search_ivanti_url = "https://tenant.saasit.com/api/odata/businessobject/incidents?$filter=ProjectLink ne '$NULL' and Status ne 'closed' and Status ne 'resolved'&$select=IncidentNumber,TypeOfIncident,Subject,LastModDateTime,Owner,OwnerTeam,Status&$top=100" #the first search record before loop
request_iv_get_recID = requests.get(url=search_ivanti_url, headers=iv_auth_header)
request_iv_get_recID_text = request_iv_get_recID.text
json_data_request_first = json.loads(request_iv_get_recID_text)
df = pd.json_normalize(json_data_request_first['value'])
df.to_csv(newpath + "/" + 'report_'+ datetime_Format + '.csv', index=False, mode='a') # the first df to run the first 100 records (Ivanti API has a constraint to pull 100 MAX records)

count = 1 #there is already the first 100 records to adjust this search query.
while count < 10:
    # Code to be executed in each iteration.
    # Will run 10 times to ensure max incidents are in this report.
    try: 
      search_ivanti_url = "https://tenant.saasit.com/api/odata/businessobject/incidents?$filter=ProjectLink ne '$NULL' and Status ne 'closed' and Status ne 'resolved'&$select=IncidentNumber,TypeOfIncident,Subject,LastModDateTime,Owner,OwnerTeam,Status&$skip=" +str(count) +"00&$top=100"
      request_iv_get_recID = requests.get(url=search_ivanti_url, headers=iv_auth_header)
      request_iv_get_recID_text = request_iv_get_recID.text
      json_data_request_first = json.loads(request_iv_get_recID_text)
      df = pd.json_normalize(json_data_request_first['value'])
      df.to_csv(newpath + "/" + 'xtraction_'+ datetime_Format + '.csv', index=False, mode='a', header=False) #False header as headers were already created in the first run OUTSIDE of this loop.
      count += 1
      print("Iterations: " + str(count))
    except ValueError:
       print("Decoding JSON failed (There is no more tickets to search for in this while loop.)")
       break
       

outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'emailtosend@gmail.com'
mail.Subject = 'Report: Project Linked tickets (Open) ' + datetime_Format
mail.Body = 'Project linked tickets'
attachment = newpath + "/" + 'reports_'+ datetime_Format + '.csv'
mail.Attachments.Add(attachment)

mail.Send()
