#!/usr/bin/env python

# Getting wmic working: https://www.si458.co.uk/?p=134
# wmic -U INDIGO/<user>%<pass>//MS4CCC6AE71035 "SELECT Product FROM Win32_BaseBoard"

# https://pypi.python.org/pypi/wmicq/1.0.0
# https://www.activexperts.com/admin/scripts/wmi/python/
# https://msdn.microsoft.com/en-us/library/aa394589(v=vs.85).aspx?cs-save-lang=1&cs-lang=vb#code-snippet-1
from ldap3 import Server, Connection
import wmi_client_wrapper as wmi
import os

server_address = "e5070s01sv001.indigo.schools.internal"
domain = "indigo"
user = "<username>"
password = "<password>"
OU_csv = "./OUs.csv"


# Contacts LDAP to get a list of hostnames that are in the specified OU
def get_ou_computers(OU):

    server = Server(server_address)
    with Connection(server, user=domain + "\\" + user, password=password) as conn:

        BaseDN = OU + "OU=Desktops,OU=School Managed,OU=Computers,OU=E5070S01,OU=Schools,DC=indigo,DC=schools,DC=internal"
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
def get_computer_details(hostnames):
    
    boards = ()
    for hostname in hostnames:
        output = os.popen("/bin/wmic --user=" + domain + "/" + user +"%" + password + " //" + hostname + " \"SELECT Version FROM Win32_ComputerSystemProduct\"").read()
        boards += (output.split("|")[2][8:],)

    return boards 


# Opens a CSV with the OUs you'd like computer details from
with open(OU_csv) as csv_read:
    #OUs = list(csv.reader(csv_read))
    OUs = ("OU=Room 13,OU=Block H,",)

    for OU in OUs:

        hostnames = get_ou_computers(OU)
        hostnames = get_computer_details(hostnames)
        boardnames = get_board_details(hostnames)

        print(OU)
        for h, b in zip(hostnames, boardnames):
            print("hostname: " + h + " model: " + b)

