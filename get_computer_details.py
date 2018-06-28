#!/usr/bin/env python

# Getting wmic working: https://www.si458.co.uk/?p=134
# wmic -U INDIGO/<user>%<pass> //MS4CCC6AE71035 "SELECT Version FROM Win32_ComputerSystemProduct"
# wmic -U INDIGO/<user>%<pass> //MS4CCC6AE71035 "COMPUTERSYSTEM GET username"
# wmic -U INDIGO/<user>%<pass> //MS4CCC6AE71035 "SELECT UserName FROM Win32_ComputerSystem"


# https://pypi.python.org/pypi/wmicq/1.0.0
# https://www.activexperts.com/admin/scripts/wmi/python/
# https://msdn.microsoft.com/en-us/library/aa394589(v=vs.85).aspx?cs-save-lang=1&cs-lang=vb#code-snippet-1
from ldap3 import Server, Connection
import wmi_client_wrapper as wmi
import sys
import csv
import os

server_address = "e5070s01sv001.indigo.schools.internal"
domain = "indigo"
user = "<username>"
password = "<password>"
library_csv = ""

# Contacts LDAP to get a list of hostnames that are in the specified OU
def get_ou_computers(OU):

    server = Server(server_address)
    with Connection(server, user=domain + "\\" + user, password=password) as conn:

        BaseDN = OU + ",OU=Desktops,OU=School Managed,OU=Computers,OU=E5070S01,OU=Schools,DC=indigo,DC=schools,DC=internal"
        Filter = "(objectCategory=computer)"
        conn.search(BaseDN,Filter)

        hostnames = ()
        for entry in conn.entries:
            hostnames += (str(entry).split(',')[0][7:],)

    return hostnames

# Ping computers to see if they're online
def get_online_computers(hostnames):
    online = ()
    for hostname in hostnames:
        if (os.system("ping -c 1 " + hostname + " > /dev/null 2>&1") == 0):
            online += (hostname,)
    
    return online

# Contacts WIM to get a list of boardnames associated with the supplied hostnames
def get_barcodes(hostnames):
    
    barcodes = ()
    for hostname in hostnames:
        output = os.popen("/bin/wmic --user=" + domain + "/" + user +"%" + password + " //" + hostname + " \"SELECT Version FROM Win32_ComputerSystemProduct\"").read()
        barcodes += (output.split("|")[2][8:],)

    return barcodes 

# Return (barcode, serial, year)
def get_records(barcode, library_csv)
    
    records= ()
    with open(library_csv) as csv_file:
        lines = tuple(csv.reader(csv_file))

        for barcode in barcodes:
            records += filter(lambda record: barcode in record[0], lines)

    return records 

def main():

    OUs = ("OU=Room 10,OU=Block G",)
#    OUs = ("OU=Room 1,OU=Block E",
#           "OU=Room 2,OU=Block E",
#           "OU=Room 3,OU=Block E",
#           "OU=Room 4,OU=Block E",
#           "OU=Room 5,OU=Block F",
#           "OU=Room 6,OU=Block F",
#           "OU=Room 7,OU=Block F",
#           "OU=Room 8,OU=Block F",
#           "OU=Room 9,OU=Block G",
#           "OU=Room 10,OU=Block G",
#           "OU=Room 11,OU=Block G",
#           "OU=Room 12,OU=Block G",
#           "OU=Room 13,OU=Block H")

    for OU in OUs:

        hostnames = get_ou_computers(OU)
        hostnames = get_online_computers(hostnames)
        barcodes = get_barcodes(hostnames)
        records = get_records(barcodes, library_csv)

        print(OU)
        print(str(len(hostnames)) + " computers")
        for b, s, y in records:
            print("barcode: " + b + " serial: " + s + " year: " + y)



main()
