import requests
import json
import datetime
from collections import defaultdict
import xlwt
from tempfile import TemporaryFile

#Asks the date range you want to scan
startVar = raw_input("What is the start date you would like to check? (Use YYYY-MM-DD format): ")
endVar = raw_input("What is the end date you would like to check? (Use YYYY-MM-DD format): ")

#This block performs a GET request from Cylance to get all the devices
url = "https://protectapi.cylance.com/devices/v2"
querystring = {"page":"1","page_size":"10000"}
headers = {
    'Accept': "application/json",
    'Authorization': "ENTER YOUR CYLANCE API KEY HERE",
    'Cache-Control': "no-cache",
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

#This block takes the users range of dates they want to search and returns the device IDs that are in that range
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
