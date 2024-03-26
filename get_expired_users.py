#!/usr/bin/env python3

import csv
import datetime
from datetime import datetime, timezone
from ldap3 import Server, Connection

# Returns a list of all users locked in LDAP
def get_ldap_expired_users(site_code, user, password, group):
    ldap_users = []
    
    print("e" + site_code + "s01sv001.indigo.schools.internal")
    server = Server("e" + site_code + "s01sv001.indigo.schools.internal")
    with Connection(server, user=user, password=password) as conn:

        BaseDN = 'OU=School Users,DC=indigo,DC=schools,DC=internal'
        Filter = '(&(SAMAccountType=805306368)(memberOf=CN=' + group + ',OU=School Managed,OU=Groups,OU=E' + site_code + 'S01,OU=Schools,DC=indigo,DC=schools,DC=internal)(!(userAccountControl:1.2.840.113556.1.4.803:=2)))'
   
        conn.search(BaseDN, Filter, attributes=['cn','accountExpires'])
        print("Number of Student Accounts: " + str(len(conn.entries)))
        for entry in conn.entries:

            # Checking if the account has expired
            account_expires = entry['accountExpires'].value
            if account_expires <= datetime.now(timezone.utc):

                ldap_users.append(str(entry['cn']))
                
    print(ldap_users)
    return ldap_users


# Returns a list of all users logged as locked
# CSV format: user,lock_status,time_stamp
def get_log_expired_users(file_location):
    log_users = {}
    with open(file_location) as csv_file:
        
        lines = reversed(list(csv.reader(csv_file)))

        for line in lines:
            if len(line) == 0 or line[0] in log_users:
                continue
           
            log_users[line[0]] = line[1]

    return list(filter(lambda user: log_users[user] == 'locked', log_users))


# Appends changes in state for each user to the log file
def update_log():
    file_location = "C:/bin/ldap_users.csv"
    site_code = "5070"
    username = "e4088746"
    password = "(8)-<~king*"

    ldap_locked = get_ldap_expired_users(site_code, "indigo\\" + username,password, "E" + site_code + "S01-InternetAccess-AllStudents")
    #ldap_locked += get_ldap_locked_users(site_code," indigo\\" + username, password, "E" + site_code + "S01-InternetAccess-AllStaff")
    log_locked = get_log_expired_users(file_location)

    append_locked  = list(set(ldap_locked) - set(log_locked))
    append_unlocked = list(set(log_locked) - set(ldap_locked))

    with open(file_location, "a") as csv_file:
        for user in append_locked:
            csv_file.write(user + "," + "locked" + "," +  datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + "\n")

        for user in append_unlocked:
            csv_file.write(user + "," + "unlocked" + "," +  datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + "\n")

    print("Current Locked: " + str(ldap_locked))


update_log()
