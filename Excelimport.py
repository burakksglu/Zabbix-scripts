# -*- coding: utf-8 -*- 
import getpass
import pprint
import sys
import xlrd
from pyzabbix import ZabbixAPI

# reload(sys)
# sys.setdefaultencoding( "utf-8" )

#File input (sudo apt-get install python3-tk)
import tkinter as tk
from tkinter import filedialog

root = tk.Tk()
root.withdraw()

file_name = filedialog.askopenfilename()


#Zabbix Api connection
try:
    url=input("URL for Zabbix Server(Please enter full URL): ")
    zapi = ZabbixAPI(url)
    user=input("Username: ")
    pw=getpass.getpass("Password: ")
    zapi.login(user, pw)
except:
    print ("\n Can not login, please check URL & credentials.")
    sys.exit()

#Template ID Get Method
def get_templateid(template_name):
    template_data = {
        "host": [template_name]
    }
    result = zapi.template.get(filter=template_data)
    if result:
        return result[0]['templateid']
    else:
        return result

#Checking for existing group
def check_group(group_name):
    return zapi.hostgroup.get(output= "groupid",filter= {"name": group_name})

#Creating new group
def create_group(group_name):
    groupid=zapi.hostgroup.create({"name": group_name})
    return groupid

#Group ID Get Method
def get_groupid(group_name):
    group_id=check_group(group_name)[0]["groupid"]
    return group_id

#Create Method for Host, also checks if it exists
def create_host(host_data):
    host=zapi.host.get(output= ["host"],filter= {"host": host_data["host"]})
    if len(host) > 0 and host[0]["host"] == host_data["host"]:
      #print "host %s exists" % host_data["name"]
      print ("host exists: %s ,group: %s ,templateid: %s" % (host_data["name"],host_data["groups"][0]["groupid"],host_data["templates"][0]["templateid"]))
    else:
      zapi.host.create(host_data)
      print ("host created: %s ,group: %s ,templateid: %s" % (host_data["name"],host_data["groups"][0]["groupid"],host_data["templates"][0]["templateid"]))

#Reading XLSX 
def open_excel(file= file_name):
     try:
         data = xlrd.open_workbook(file)
         return data
     except Exception(e):
         print (str(e))

#Get Method for hosts from XLSX
def get_hosts(file):
    data = open_excel(file)
    table = data.sheets()[0]
    nrows = table.nrows
    ncols = table.ncols
    list = []
    for rownum in range(1,nrows):
      list.append(table.row_values(rownum)) 
    return list

def main():
  hosts=get_hosts(file_name)
  for host in hosts:
    host_name=host[5]
    visible_name="Dogus "+host[1]+" "+host[5]+" | "+host[6]
    host_ip=host[6]
    group=host[0]+"/"+host[2]+"/"+host[3]
    polling_method =host[7]
    if polling_method == "SNMP":
        template = "Template Net Network Generic Device SNMPv2"
    else:
        template = "Template Module ICMP Ping"
    templateid=get_templateid(template)
    if not check_group(group):
        # print (u'Added host grup: %s' % group)
        groupid=create_group(group)
    groupid=get_groupid(group)
    host_data = {
        "host": host_name,
        "name": visible_name,
        "interfaces": [
            {
                "type": 2,
                "main": 1,
                "useip": 1,
                "ip": host_ip.strip(),
                "dns": "",
                "port": "161",
                "details": {
                    "version": 2,
                    "community": "public"
                }
            }
        ],
        "groups": [
            {
                "groupid": groupid
            }
        ],
        "templates": [
            {
                "templateid": templateid
            }
        ]
    }
    create_host(host_data)

if __name__=="__main__":
    main()
