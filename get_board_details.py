#!/usr/bin/env python3
# https://www.activexperts.com/admin/scripts/wmi/python/
from ldap3 import Server, Connection
import win32com.client

server_address = "e5070s01sv001.indigo.schools.internal"
user = "indigo\\<username>"
password = "<password>"

# Opens a CSV with the OUs you'd like computer details from
with open(OU_csv) as csv_read:
    OUs = list(csv.reader(csv_read)):

        for OU in OUs:

            hostnames = get_ou_computers(OU)
            boardnames = get_board_details(hostnames)

            print(OU)
            for h, b in zip(hostnames, boardnames):
                print("hostname: " + h + " boardname: " + b)

# Contacts LDAP to get a list of hostnames that are in the specified OU
def get_ou_computers(OU):

    server = Server(server_address)
    with Connection(server, user=user, password=password) as conn:

        BaseDN = 'DC=' + OU + ',DC=indigo,DC=schools,DC=internal'
        Filter = "(objectCategory=computer)"

        conn.search(BaseDN,Filter)
        hostnames = ()
        for entry in conn.entries:
            hostnames += entry

    return hostnames

# Contacts WIM to get a list of boardnames associated with the supplied hostnames
def get_board_details(hostnames):
        
    boards = ()
    for hostname in hostnames:
        # Possibly change the server address and put the hostname as a WHERE clause
        objWMIService = win32com.client.Dispatch("WbemScripting.SWbemLocator")
        objSWbemServices = objWMIService.ConnectServer(hostname, user + '\\' + password)
        colItems = objSWbemServices.ExecQuery("SELECT * FROM Win32_BaseBoard")


        for objItem in colItems:
            if objItem.Model != None:
                boards += objItem.Model

    return boards 
