#!/usr/bin/env python3

from ldap3 import Server, Connection

def get_computer_details(server_address, user, password, hostname):

    server = Server(server_address)
    with Connection(server, user=user, password=password) as conn:

        BaseDN = 'DC=indigo,DC=schools,DC=internal'
        Filter = "(&(objectCategory=computer)(name=" + hostname + "))"

        conn.search(BaseDN,Filter)
        for entry in conn.entries:
            print(entry)

get_computer_details("e5070s01sv001.indigo.schools.internal","indigo\\<username>","<password>","MS1C6F6508EDBE")
