from pathlib import Path
from datetime import datetime
import requests
import pprint
import xlwings as xw
import wmi
import math

conn = wmi.WMI()



 
url = "https://xltraders.online/api/v1/validate"
 
 
 

# file_path = Path(__file__).parent / 'hello.txt'

def sysinfo():
    sysdrive = conn.Win32_OperatingSystem()[0].SystemDrive
    freespace = next((f"{100 * int(i.FreeSpace) / int(i.Size):.2f}" for i in conn.Win32_LogicalDisk() if i.DeviceID == sysdrive), None)
    
    osdrivefreecheck = None
    
    try:
        osdrivefreecheck = float(freespace)>50.00
    except:
        osdrivefreecheck = None
        
    return {
            "sysID": conn.Win32_ComputerSystemProduct()[0].UUID,
            "vendor" : conn.Win32_ComputerSystemProduct()[0].vendor,
            "proc" : conn.Win32_Processor()[0].Name,
            "tmem" : ("{:.2f}".format(math.ceil(int(conn.Win32_Computersystem()[0].TotalPhysicalMemory)/1024/1024/1024))),
            "username" : conn.Win32_Computersystem()[0].UserName,
            "partofdomain" : conn.Win32_Computersystem()[0].PartOfDomain,
            "workgroup" : conn.Win32_Computersystem()[0].Workgroup,
            "freemem" : ("{:.2f}".format(int(conn.Win32_OperatingSystem()[0].FreePhysicalMemory)/1024/1024)),
            "osarch" : conn.Win32_OperatingSystem()[0].OSArchitecture,
            "sysdrive" : sysdrive,
            "serialnum" : conn.Win32_Baseboard()[0].SerialNumber,
            "manufacturer" : conn.Win32_Baseboard()[0].Manufacturer,
            "product" : conn.Win32_Baseboard()[0].Product,
            "macs" : [mac.MACAddress for mac in conn.Win32_NetworkAdapterConfiguration() if mac.MACAddress is not None],
            "freespace" : f'{freespace}%',
            "osdrivefreecheck" : osdrivefreecheck
            }

def main():
    wb = xw.Book.caller()
    sheet = wb.sheets[0]
    if sheet.range('B5').value == None:
        sheet.range('C5').value = "Please input user id"
    else:
        sheet.range('C5').value = "OK"

    if sheet.range('B6').value == None:
        sheet.range('C6').value = 'No activation key found, subscribe any plan first to get activation key'
    else:
        sheet.range('C6').value = "OK"

    if sheet.range('B5').value != None and sheet.range('B6').value != None:
        userid = sheet.range('B5').value
        token = sheet.range('B6').value
        data = {"userid": userid, "token": token }
        response = requests.post(url, json=data)

        if response.status_code == 200:
            date_string = response.json()["validity"]
            date_format = '%m-%d-%Y'
            desired_date = datetime.strptime(date_string, date_format)
            current_date = datetime.now()
            time_difference = desired_date - current_date
            days_difference = time_difference.days
            sheet.range("B9").value = days_difference
            sheet.range("B8").value = date_string
            sheet.range("E5").value = ""
        else:
            sheet.range("E5").value = response.json()["Value"]
            sheet.range("B9").value = ""
            sheet.range("B8").value = ""
    else:
        sheet.range('A1').value = "Correct the errors please"
        
# def final_data():
#     data = sysinfo()
#     data["userid"] = userid
#     data["token"] = token
#     return data

# pprint.pprint(final_data())


@xw.func
def hello(name):
    return f"Hello {name}!"


if __name__ == "__main__":
    xw.Book("demo.xlsm").set_mock_caller()
    main()
