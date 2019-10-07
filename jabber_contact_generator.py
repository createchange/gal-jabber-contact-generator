# -*- coding: utf-8 -*-
"""
Created on Fri Oct 19 17:15:23 2018

@author: jhweaver

Gather info from GAL, write to XML document for importing into Jabber.
"""

import json
import subprocess


def get_GAL_contacts():
    '''
    Assign function to variable, returns address list in json format
    '''
    subprocess.call(r'OutlookAddressBookView.exe /sjson "temp/test_contacts.json" /AddressBookName "All Users"', shell=True)
    with open(r"temp/test_contacts.json", 'r', encoding="utf16") as f:
        json_string = f.read()
        f.close()
    return json.loads(json_string)

def org_dictionary(contacts_list):
    '''
    Takes json formatted contact list and returns dictionaries in the 
    following format: ORG[DEPARTMENT][USER]:[EMAIL]
    
    e.g. INT['Clerical']['Jonathan H. Weaver]:[jhweaver@intinc.com]
    
    This filters out entries with empty "Department" field in AD.
    '''
    for x in contacts_list[:]:
        for k, v in x.items():
            if k == "Company":
                comp = v
            if k == "Display Name":
                name = v
            if k == "Email Address":
                email = v
            if k == "Department Name":
                dept = v
        if dept == "":
            pass
        else:
            if comp == "INT":
                try:
                    INT[dept][name] = [email]
                except:
                    INT[dept] = {}
                    INT[dept][name] = [email]
                
            elif comp == "MSM":
                try:
                    MSM[dept][name] = [email]
                except:
                    MSM[dept] = {}
                    MSM[dept][name] = [email]


def xml_department_writer(company_name, department, user_dict):
    '''
    Takes INT and MSM list as input, writes XML info to file 
    for importing to Jabber
    '''
    with open(r'output/%s/%s-%s-contacts.xml' % (company_name, company_name, department), 'w') as f:
        f.write('<?xml version="1.0" encoding="utf-8" ?>\n')
        f.write('<buddylist>\n')
        f.write("<group>\n")
        f.write("<gname>" + company_name + " " + department + "</gname>\n")
        for k,v in user_dict.items():
            f.write("<user>\n")
            f.write("<uname>" + v[0] + "</uname>\n")
            f.write("<fname>" + k + "</fname>\n")
            f.write("</user>\n")
            f.write("</group>\n")
        f.write("</buddylist>")
        
def xml_org_writer(company_name, company_dict):
    '''
    Expects: company_name (string), company_dict (dict)
    Generates master list of contacts, putting to 
    top: INT
    bottom: MSM
    '''
    with open(r'output/master/%s-master-contacts.xml' % company_name, 'w') as f:
        f.write('<?xml version="1.0" encoding="utf-8" ?>\n')
        f.write('<buddylist>\n')
        for k,v in company_dict.items():
            dept = k
            for k,v in company_dict[dept].items():
                f.write("<group>\n")
                f.write("<gname>" + company_name + " " + dept + "</gname>\n")
                f.write("<user>\n")
                f.write("<uname>" + v[0] + "</uname>\n")
                f.write("<fname>" + k + "</fname>\n")
                f.write("</user>\n")
                f.write("</group>\n")
        f.write("</buddylist>")
        
        
def generate_org_list():
    xml_org_writer('INT', INT)
    xml_org_writer('MSM', MSM)
        
def generate_department_list():           
    for k,v in INT.items():
        dept = k
        xml_department_writer('INT', dept, INT[dept])
    
    for k,v in MSM.items():
        dept = k
        xml_department_writer('MSM', dept, MSM[dept])
            
        
"""
Start program
"""

# Get contacts from Outlook GAL
print("Working...")
contacts_list = get_GAL_contacts()

# Initialize empty org dictionaries
INT = {}
MSM = {}

# Fills respective dictionaries with dictionaries based on dept.
org_dictionary(contacts_list)

# Generates master list for INT and MSM
generate_org_list()

# Generates contact directories, outputting to file system
generate_department_list()

print("Done. You can find the XML files in the output folder.")
input("Press any button to continue...")