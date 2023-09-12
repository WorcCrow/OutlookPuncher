#https://medium.com/@manojkumardhakad/python-read-and-send-outlook-mail-using-oauth2-token-and-graph-api-53de606ecfa1

import requests
import json
import re
import os
import time
import datetime

tenant_id = ''
client_id = ''
refresh_token = ''
client_secret = ''

file_path = "outlook.log"
data = []
counter = 0

def addRecord(link):
  try:
    data.append(link)
    with open(file_path, 'a') as file:
        file.write(link + '\n')
  except:
    print("Error: addRecord")
  

def loadData():
  try:
    with open(file_path, 'r') as file:
      for line in file:
          data.append(line.strip())
  except:
    print("Error: loadData")     

def printData():
  try:
    print("Current Record: ")
    for line in data:
      print(line)
  except:
    print("Error: printData")   

def refreshToken():
  try:
    url = "https://login.microsoftonline.com/" + tenant_id + "/oauth2/v2.0/token"
    payload = {
        'client_id': client_id,
        'scope': 'offline_access Mail.ReadWrite Mail.send',
        'grant_type': 'refresh_token',
        'client_secret': client_secret,
        'refresh_token': refresh_token
        }
    files=[]
    headers = {
      'Cookie': 'fpc=ArzN7nqM-xNKv6hKpPUD5qn3DVvmAQAAAOqEg9wOAAAAEFVrtQMAAABohYPcDgAAAA; stsservicecookie=estsfd; x-ms-gateway-slice=estsfd'
    }
    response = requests.request("POST", url, headers=headers, data=payload, files=files)
    token = json.loads(response.text)
    return token
  except:
    print("Error: refreshToken")
  return ""

def checkPuchline():
  token = ""
  try:
    token = refreshToken()
    url = "https://graph.microsoft.com/v1.0/me/messages?$search=\"from:attendancetracking@salary.com\"&top=1"
    payload = {}
    headers = {
      'Authorization': 'Bearer ' + token["access_token"]
    }

    response = requests.request("GET", url, headers=headers, data=payload)

    responseJson = json.loads(response.text)
    content = responseJson["value"][0]["body"]["content"]

    punchlink = re.findall('href="([^"]+)"', content)
    #print(content,"\n")
    #print(punchlink,"\n")
    #print(punchlink[0],"\n")
    link = punchlink[0].replace("&amp;","&")
    #print(link,"\n")
    if link in data:
      print("Time Check:",datetime.datetime.now())
    else:
      print("\nGoing to Punch ", link)
      addRecord(link)
      os.system("start chrome \"" + link + "\"")
  except:
    print("Time of Error:", datetime.datetime.now())
    print("Error: checkPuchline ", token)


loadData()
printData()
while True:
  checkPuchline()
  time.sleep(300)
  #print("\n" * 100)







