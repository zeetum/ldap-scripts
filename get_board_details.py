#!/usr/bin/env python3

# Getting wmic working: https://www.si458.co.uk/
# wmic -U INDIGO/<user>%<pass> //MS4CCC6AE71035 "SELECT Product FROM Win32_BaseBoard"

# https://www.activexperts.com/admin/scripts/wmi/python/
# https://msdn.microsoft.com/en-us/library/aa394589(v=vs.85).aspx?cs-save-lang=1&cs-lang=vb#code-snippet-1
from ldap3 import Server, Connection
#import win32com.client

server_address = "e5070s01sv001.indigo.schools.internal"
user = "indigo\\e4088746"
password = "Wonderful4"
OU_csv = "./OUs.csv"


# Contacts LDAP to get a list of hostnames that are in the specified OU
def get_ou_computers(OU):

    server = Server(server_address)
    with Connection(server, user=user, password=password) as conn:

        BaseDN = OU + "OU=Desktops,OU=School Managed,OU=Computers,OU=E5070S01,OU=Schools,DC=indigo,DC=schools,DC=internal"
        Filter = "(objectCategory=computer)"
        conn.search(BaseDN,Filter)

        hostnames = ()
        for entry in conn.entries:
            hostnames += (str(entry).split(',')[0][7:],)

    return hostnames

# Contacts WIM to get a list of boardnames associated with the supplied hostnames
def get_board_details(hostnames):
        
    boards = ()
    for hostname in hostnames:
        # Possibly change the server address and put the hostname as a WHERE clause
        objWMIService = win32com.client.Dispatch("WbemScripting.SWbemLocator")
        objSWbemServices = objWMIService.ConnectServer(hostname, "root\cimv2")
        colItems = objSWbemServices.ExecQuery("SELECT * FROM Win32_BaseBoard")


        for objItem in colItems:
            if objItem.Product != None:
                boards += objItem.Product

    return boards 


# Opens a CSV with the OUs you'd like computer details from
with open(OU_csv) as csv_read:
    #OUs = list(csv.reader(csv_read))
    OUs = ("OU=Room 10,OU=Block G,",)

    for OU in OUs:
        print(OU)
        hostnames = get_ou_computers(OU)
        # Get win32com working
        boardnames = get_board_details(hostnames)
        boardnames = hostnames

        print(OU)
        for h, b in zip(hostnames, boardnames):
            print("hostname: " + h + " boardname: " + b)

