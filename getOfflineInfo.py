import requests
import json
import datetime
from collections import defaultdict
import xlwt
from tempfile import TemporaryFile


#Asks the date range you want to scan
startVar = raw_input("What is the start date you would like to check? (Use YYYY-MM-DD format): ")
endVar = raw_input("What is the end date you would like to check? (Use YYYY-MM-DD format): ")

#how to upgrade pip
#python -m pip install --upgrade pip --trusted-host=pypi.python.org --trusted-host=pypi.org --trusted-host=files.pythonhosted.org
#how to install xlwt if module won't import
#python -m pip install xlwt --trusted-host=pypi.python.org --trusted-host=pypi.org --trusted-host=files.pythonhosted.org

#Key expires at: 2:45pm

#This block performs a GET request from Cylance to get all the devices
url = "https://protectapi.cylance.com/devices/v2"
querystring = {"page":"1","page_size":"10000"}
headers = {
    'Accept': "application/json",   #Every 30 minutes the Authorization key below expires and you must regenerate and copy + paste.
    'Authorization': "Bearer eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJzdWIiOiI2OWI2MzQyMC1iMjM2LTQ2OTktYmRlYy1kZThjOTAxYjE0MzkiLCJpYXQiOjE1NjU5Nzc1NDAsInNjcCI6WyJhcHBsaWNhdGlvbjpsaXN0IiwiZGV2aWNlOmxpc3QiLCJkZXZpY2U6cmVhZCIsImRldmljZTp0aHJlYXRsaXN0IiwiZ2xvYmFsbGlzdDpsaXN0IiwiZ2xvYmFsbGlzdDpyZWFkIiwib3B0aWNzY29tbWFuZDpsaXN0Iiwib3B0aWNzY29tbWFuZDpyZWFkIiwib3B0aWNzZGV0ZWN0Omxpc3QiLCJvcHRpY3NkZXRlY3Q6cmVhZCIsIm9wdGljc2V4Y2VwdGlvbjpsaXN0Iiwib3B0aWNzZXhjZXB0aW9uOnJlYWQiLCJvcHRpY3Nmb2N1czpsaXN0Iiwib3B0aWNzZm9jdXM6cmVhZCIsIm9wdGljc3BrZ2NvbmZpZzpsaXN0Iiwib3B0aWNzcGtnY29uZmlnOnJlYWQiLCJvcHRpY3Nwa2dkZXBsb3k6bGlzdCIsIm9wdGljc3BrZ2RlcGxveTpyZWFkIiwib3B0aWNzcG9saWN5Omxpc3QiLCJvcHRpY3Nwb2xpY3k6cmVhZCIsIm9wdGljc3J1bGU6bGlzdCIsIm9wdGljc3J1bGU6cmVhZCIsIm9wdGljc3J1bGVzZXQ6bGlzdCIsIm9wdGljc3J1bGVzZXQ6cmVhZCIsIm9wdGljc3N1cnZleTpsaXN0Iiwib3B0aWNzc3VydmV5OnJlYWQiLCJwb2xpY3k6bGlzdCIsInBvbGljeTpyZWFkIiwidGhyZWF0OmRldmljZWxpc3QiLCJ0aHJlYXQ6bGlzdCIsInRocmVhdDpyZWFkIiwidXNlcjpsaXN0IiwidXNlcjpyZWFkIiwiem9uZTpsaXN0Iiwiem9uZTpyZWFkIl0sInRpZCI6Ijk2ODU5OWQ1LTM0YmQtNDg3YS04M2U1LWFmOWViOWY4OGY5OCIsImlzcyI6Imh0dHA6Ly9jeWxhbmNlLmNvbSIsImF1ZCI6Imh0dHBzOi8vcGFwaS5jeWxhbmNlLmNvbS9hcGkiLCJleHAiOjE1NjU5NzkzNDAsIm5iZiI6MTU2NTk3NzU0MH0.TJGa6gWnYcP4W3LE5TBdjMv2r_E8k4SAk7H9tLiRKlw",
    'Cache-Control': "no-cache", #MAKE SURE YOU LEAVE THE "Bearer" TEXT BEFORE THE KEY. NOT SURE BUT IT WILL ONLY WORK IF YOU KEEP IT.
    'Host': "protectapi.cylance.com",
    'Accept-Encoding': "gzip, deflate",
    'Connection': "keep-alive",
    'cache-control': "no-cache"
    }

response = requests.request("GET", url, headers=headers, params=querystring)
rawData = response.text

#This block parses all of the devices and returns a list of the offline ones by ID name
splitData = rawData.split("{")
offlineAssets = list()
for item in splitData:
    if 'Offline' in item:
        offlineAssets.append(item)
    else:
        pass

stringOfOfflineAssets = str(offlineAssets)
splitOfflineAssets = stringOfOfflineAssets.split(",")
rawID = list()
for asset in splitOfflineAssets:
    if 'id' in asset:
        rawID.append(asset)
    else:
        pass

stringID = str(rawID)
splitID = stringID.split('"')
finalID = list()
for id in splitID:
    if '-' in id:
        finalID.append(id)
    else:
        pass

#This block performs another GET request on each individual device by ID to obtain the last date connected
deviceurl = "https://protectapi.cylance.com/devices/v2/"
deviceurlList = []
for param in finalID:
    deviceurlList.append(deviceurl + param)

deviceDict = {}
for url in deviceurlList:
    devicequerystring = {"":""}
    deviceresponse = requests.request("GET", url, headers=headers, params=devicequerystring)
    deviceRawData = deviceresponse.text
    deviceDictRaw = json.loads(deviceRawData)
    offlineDateTime = deviceDictRaw["date_offline"]
    offlineDateTime = str(offlineDateTime)
    deviceID = deviceDictRaw["id"]
    deviceID = str(deviceID)
    tempDict = {}
    for item in deviceDictRaw:
        tempDict[offlineDateTime] = deviceID
    deviceDict.update(tempDict)

#This block prompts the user for a range of dates they want to search and returns the device IDs that are in that range
dateTimeList = deviceDict.keys()
dateListRaw = list()
for item in dateTimeList:
    splitString = item.split("T")
    dateListRaw.append(splitString[0])

    #Some devices don't have an offline date, so "None" is returned. This parses those values out.
if "None" in dateListRaw:
    dateListRaw.remove("None")    
else:
    pass

dateListFinal = [datetime.datetime.strptime(date, '%Y-%m-%d').date() for date in dateListRaw]
start = datetime.datetime.strptime(startVar, "%Y-%m-%d").date()
end = datetime.datetime.strptime(endVar, "%Y-%m-%d").date()

#This block compares the offline_dates we have with the date range they requested. If the device's offline_date is in the range, it will be appeneded to a list.
deviceIDList = list()
for date in dateListFinal:
    if start <= date <= end:
        date = str(date)
        deviceIDRaw = ([v for k,v in deviceDict.items() if k.startswith(date)])
        if deviceIDRaw in deviceIDList:
            pass
        else:
            deviceIDList.append(deviceIDRaw)
    else:
        pass

#At this point in time, the deviceIDList has multiple IDs as one item. This block will split them apart and make each ID its own values
deviceIDString = str(deviceIDList)
splitIDs = deviceIDString.split("'")
finalIDList = list()
for asset in splitIDs:
    if '-' in asset:
        finalIDList.append(asset)
    else:
        pass

#This block performs another GET request on each individual device by ID that was in the specified range and return the host name, OS, offlinedate, IP, and MAC address for each one
newdeviceurl = "https://protectapi.cylance.com/devices/v2/"
newdeviceurlList = []
for param in finalIDList:
    newdeviceurlList.append(newdeviceurl + param)

deviceInfoList = list()
for url in newdeviceurlList:
    devicequerystring = {"":""}
    deviceresponse = requests.request("GET", url, headers=headers, params=devicequerystring)
    deviceRawData = deviceresponse.text
    deviceDictRaw = json.loads(deviceRawData)
    deviceInfoList.append(deviceDictRaw["host_name"])
    deviceInfoList.append(deviceDictRaw["os_version"])
    deviceInfoList.append(deviceDictRaw["date_offline"])
    deviceInfoList.append(deviceDictRaw["ip_addresses"])
    deviceInfoList.append(deviceDictRaw["mac_addresses"])
print(deviceInfoList)

#This block formats the data into an excel sheet
workbook = xlwt.Workbook()
worksheet = workbook.add_sheet('Offline Devices')

worksheet.write(0, 0, "Host Name")
worksheet.write(0, 1, "Operating System")
worksheet.write(0, 2, "Offline Date")
worksheet.write(0, 3, "IP Address(es)")
worksheet.write(0, 4, "MAC Address(es)")

iterator = 0
col = 0
row = 1
for data in deviceInfoList:
    worksheet.write(row, col, data)
    col = col + 1
    iterator = iterator + 1
    if iterator is 5:
        row = row + 1
        iterator = 0
        col = 0

name = startVar + " to " + endVar + " Offline Devices.xls"
workbook.save(name)
workbook.save(TemporaryFile())