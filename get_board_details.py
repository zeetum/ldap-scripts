#!/usr/bin/env python

# Getting wmic working: https://www.si458.co.uk/
# wmic -U INDIGO/<user>%<pass>//MS4CCC6AE71035 "SELECT Product FROM Win32_BaseBoard"

# https://pypi.python.org/pypi/wmicq/1.0.0
# https://www.activexperts.com/admin/scripts/wmi/python/
# https://msdn.microsoft.com/en-us/library/aa394589(v=vs.85).aspx?cs-save-lang=1&cs-lang=vb#code-snippet-1
from ldap3 import Server, Connection
import wmi_client_wrapper as wmi

server_address = "e5070s01sv001.indigo.schools.internal"
user = "indigo\\<username>"
password = "<pass>"
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
# change structure to call a command instead of using a library
# https://docs.python.org/3/library/subprocess.html#subprocess.check_output

    boards = ()
    for hostname in hostnames:
        wmic = wmi.WmiClientWrapper(
            username="indigo/<username>",
            password="<pass>",
            host=hostname,
        )

        output = wmic.query("\"SELECT Product FROM Win32_BaseBoard\"")
        print(output)

    return boards 


# Opens a CSV with the OUs you'd like computer details from
with open(OU_csv) as csv_read:
    #OUs = list(csv.reader(csv_read))
    OUs = ("OU=Room 13,OU=Block H,",)

    for OU in OUs:

        hostnames = get_ou_computers(OU)
        boardnames = get_board_details(hostnames)
        boardnames = hostnames

        print(OU)
        for h, b in zip(hostnames, boardnames):
            print("hostname: " + h + " boardname: " + b)


