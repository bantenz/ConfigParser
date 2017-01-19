
#PYTHON SCRIPT TO PULL DATA FROM CISCO ASA V0.3 beta(update August 30th, 2016

from ciscoconfparse import CiscoConfParse

import re
import os
import sys
import xlsxwriter
import iptools
import argparse

#remark = "";

def ConfigParser(fileName):
    
    p = CiscoConfParse(fileName)
    text = ""

    DN = p.find_objects("^hostname")
    for DNtmp in DN:
        DeviceName = DNtmp.replace("hostname", "").lstrip()
    print "Please wait, your file is generated.........."
    workbook = xlsxwriter.Workbook(DeviceName + '.xlsx')

#####################################
#                                   #
#   Extracting Network Object       #
#                                   #
#####################################

    worksheet1 = workbook.add_worksheet('Object Network')
    row = 0
    col = 0

    worksheet1.write(row, col + 1, "Object Name")
    worksheet1.write(row, col + 2, "IP Address")
    worksheet1.write(row, col + 3, "Netmask")
    worksheet1.write(row, col + 4, "NAT")
    worksheet1.write(row, col + 5, "Description")
    row += 1
    
    # Looking for Object Network Type #
    for parents in p.find_objects(r"^object network"):
        tmpNetOBJ = parents.text
        #print (tmp)
        worksheet1.write(row, col + 1, tmpNetOBJ.replace("object network", "").lstrip())

        if parents.re_search_children("host"):
         for strs in p.find_children_w_parents(tmpNetOBJ, 'host'):
                host = strs.replace("host", "").lstrip()
                #print (host)
                worksheet1.write(row, col + 2, host)
                worksheet1.write(row, col + 3, "255.255.255.255")

        if parents.re_search_children("subnet"):
            for strs in p.find_children_w_parents(tmpNetOBJ, 'subnet'):
                ip = strs.replace("subnet", "").lstrip()
                address,netmask = ip.split()
                #network = address + '/' + str(iptools.ipv4.netmask2prefix(netmask))
                #print (subnet)
                worksheet1.write(row, col + 2, address)
                worksheet1.write(row, col + 3, netmask)

        if parents.re_search_children("fqdn"):
            for strs in p.find_children_w_parents(tmpNetOBJ, 'fqdn'):
                fqdn = strs.replace("fqdn v4", "").lstrip()
                #print (fqdn)
                worksheet1.write(row, col + 2, fqdn)
                worksheet1.write(row, col + 3, "-")

        if parents.re_search_children("range"):
            for strs in p.find_children_w_parents(tmpNetOBJ, 'range'):
                srange = strs.replace("range", "").lstrip()
                #print (fqdn)
                worksheet1.write(row, col + 2, srange.replace("range", "").lstrip())
                worksheet1.write(row, col + 3, "-")
                
        if not parents.re_search_children("description"):
            #print ("description none ")
            worksheet1.write(row, col + 5, "-")
        elif parents.re_search_children("description"):
            arry = []
            for child in p.find_children_w_parents("^%s$" % tmpNetOBJ, 'description', ignore_ws=True):
                string = child.lstrip()
                arry.append(string.replace("description", "").lstrip())
                #print (', '.join(arry))
            worksheet1.write(row, col + 5, ', '.join(arry))

        row += 1

    print ("Extracting Network Object Configuration : Completed")

#####################################
#                                   #
#   Extracting Service Object       #
#                                   #
#####################################

    worksheet2 = workbook.add_worksheet('Object Service')
    row = 0
    col = 0
    
    worksheet2.write(row, col + 1, "Object Name")
    worksheet2.write(row, col + 2, "TCP/UDP")
    worksheet2.write(row, col + 3, "Source Port")
    worksheet2.write(row, col + 4, "Destination Port")
    worksheet2.write(row, col + 5, "Description")
    row += 1    

    for parents in p.find_objects(r"^object service"):
        tmpServiceOBJ = parents.text
        #print (tmp)
        worksheet2.write(row, col + 1, tmpServiceOBJ.replace("object service", "").lstrip())

        if parents.re_search_children("service tcp destination eq"):
         for strs in p.find_children_w_parents(tmpServiceOBJ, 'service tcp destination eq'):
                service = strs.replace("service tcp destination eq", "").lstrip()
                #print (service)
                worksheet2.write(row, col + 2, "TCP")
                worksheet2.write(row, col + 3, "1-65535")
                worksheet2.write(row, col + 4, service)

        if parents.re_search_children("service tcp destination range"):
            for strs in p.find_children_w_parents(tmpServiceOBJ, 'service tcp destination range'):
                service = strs.replace("service tcp destination range", "").lstrip()
                #print (service)
                worksheet2.write(row, col + 2, "TCP Range")
                worksheet2.write(row, col + 3, "1-65535")
                worksheet2.write(row, col + 4, service)

        if parents.re_search_children("service udp destination eq"):
            for strs in p.find_children_w_parents(tmpServiceOBJ, 'service udp destination eq'):
                service = strs.replace("service udp destination eq", "").lstrip()
                #print (service)
                worksheet2.write(row, col + 2, "UDP")
                worksheet2.write(row, col + 3, "1-65535")
                worksheet2.write(row, col + 4, service)

        if parents.re_search_children("service udp destination range"):
            for strs in p.find_children_w_parents(tmpServiceOBJ, 'service udp destination range'):
                service = strs.replace("service udp destination range", "").lstrip()
                #print (service)
                worksheet2.write(row, col + 2, "UDP Range")
                worksheet2.write(row, col + 3, "1-65535")
                worksheet2.write(row, col + 4, service)                

        if parents.re_search_children("service icmp"):
            for strs in p.find_children_w_parents(tmpServiceOBJ, 'service icmp'):
                #service = strs.replace("service tcp destination range", "").lstrip()
                #print (service)
                worksheet2.write(row, col + 2, "ICMP")
                worksheet2.write(row, col + 3, "-")
                worksheet2.write(row, col + 4, "-")
                
        row += 1

    print ("Extracting Service Object Configuration : Completed")

##########################################
#                                        #
#   Extracting Network Object Group      #
#                                        #
##########################################

    worksheet3 = workbook.add_worksheet('Network Group Object')
    row = 0
    col = 0

    worksheet3.write(row, col + 1, "Group Name")
    worksheet3.write(row, col + 2, "Member Name")
    worksheet3.write(row, col + 3, "IP Address")
    worksheet3.write(row, col + 4, "Netmask")
    worksheet3.write(row, col + 5, "Control Number")
    worksheet3.write(row, col + 6, "Description")
    row += 1
    netobject = ""
    
    for parents in p.find_objects(r"^object-group network"):
        tmpNetOBJGroup = parents.text
        #worksheet3.write(row, col + 1, tmpNetOBJGroup.replace("object-group network", "").lstrip())
        #print (tmp)

        if parents.re_search_children("description"):
         for strs in p.find_children_w_parents(tmpNetOBJGroup, 'description'):
             description = strs.replace("description", "").lstrip()
             worksheet3.write(row, col + 1, tmpNetOBJGroup.replace("object-group network", "").lstrip())
             worksheet3.write(row, col + 6, description)
             
        if parents.re_search_children("network-object"):
         for strs in p.find_children_w_parents(tmpNetOBJGroup, 'network-object'):            
             netobject = strs.replace("network-object", "").lstrip()
             worksheet3.write(row, col + 1, tmpNetOBJGroup.replace("object-group network", "").lstrip())
             
             #print netobject
             if 'host' in netobject:
                 netobject = netobject.replace("host", "").lstrip()
                 network = address + "/32"
                 worksheet3.write(row, col + 2, network)
                 worksheet3.write(row, col + 3, netobject)
                 worksheet3.write(row, col + 4, "255.255.255.255")
                 #print netobject
             elif "object" in netobject:
                 netobject = netobject.replace("object", "").lstrip()
                 worksheet3.write(row, col + 2, netobject)
             else:
                 address,netmask = netobject.split()
                 network = address + '/' + str(iptools.ipv4.netmask2prefix(netmask))
                 worksheet3.write(row, col + 2, network)
                 worksheet3.write(row, col + 3, address)
                 worksheet3.write(row, col + 4, netmask)
                 #print netobject
             row += 1                            
                
    print ("Extracting Network Object Group Configuration : Completed")
 
##########################################
#                                        #
#   Extracting Network service Group     #
#                                        #
##########################################

    worksheet4 = workbook.add_worksheet('Service Group Object')
    row = 0
    col = 0

    worksheet4.write(row, col + 1, "Group Name")
    worksheet4.write(row, col + 2, "TCP/UDP")
    worksheet4.write(row, col + 3, "Source Port")
    worksheet4.write(row, col + 4, "Destination Port")
    row += 1
    
    for parents in p.find_objects(r"^object-group service"):
        tmpSVCOBJGroup = parents.text
        #print (tmp)
        groupname = tmpSVCOBJGroup.replace("object-group service", "").lstrip()
        if "tcp" in groupname:
          groupname = groupname.replace("tcp", "").lstrip()
          worksheet4.write(row, col + 2, "TCP")
        elif "udp" in groupname:
          groupname = groupname.replace("udp", "").lstrip()
          worksheet4.write(row, col + 2, "UDP")

        #worksheet4.write(row, col + 1, groupname)        
        #firstrow = row
        
        if parents.re_search_children("port-object eq"):
         for strs in p.find_children_w_parents(tmpSVCOBJGroup, 'port-object eq'):
                serviceobject = strs.replace("port-object eq", "").lstrip()
                worksheet4.write(row, col + 1, groupname) 
                worksheet4.write(row, col + 3, "1-65535")
                worksheet4.write(row, col + 4, serviceobject)
                row += 1
                
        if parents.re_search_children("port-object range"):
         for strs in p.find_children_w_parents(tmpSVCOBJGroup, 'port-object range'):
                serviceobject = strs.replace("port-object range", "").lstrip()
                worksheet4.write(row, col + 1, groupname) 
                worksheet4.write(row, col + 3, "1-65535")
                worksheet4.write(row, col + 4, serviceobject)
                row += 1
                
        if parents.re_search_children("service-object object"):
         for strs in p.find_children_w_parents(tmpSVCOBJGroup, 'service-object object'):
                servicegroupobject = strs.replace("service-object object", "").lstrip()
                worksheet4.write(row, col + 1, groupname) 
                worksheet4.write(row, col + 3, "1-65535")
                if ('tcp' or 'TCP') in servicegroupobject:
                    worksheet4.write(row, col + 2, "TCP")
                elif ('udp' or 'UDP') in servicegroupobject:
                    worksheet4.write(row, col + 2, "UDP")                           
                                    
                worksheet4.write(row, col + 4, servicegroupobject)
                row += 1

        if parents.re_search_children("service-object tcp"):
         for strs in p.find_children_w_parents(tmpSVCOBJGroup, 'service-object tcp destination'):
               service = strs.replace("service-object tcp destination", "").lstrip()
               worksheet4.write(row, col + 1, groupname) 
               worksheet4.write(row, col + 2, "TCP")
               worksheet4.write(row, col + 3, "1-65535")
               worksheet4.write(row, col + 4, service.replace("eq", ""))
               row += 1

        if parents.re_search_children("service-object udp"):
         for strs in p.find_children_w_parents(tmpSVCOBJGroup, 'service-object udp'):
               service = strs.replace("service-object udp destination", "").lstrip()
               worksheet4.write(row, col + 1, groupname) 
               worksheet4.write(row, col + 2, "UDP")
               worksheet4.write(row, col + 3, "1-65535")
               worksheet4.write(row, col + 4, service.replace("eq", ""))
               row += 1

        if parents.re_search_children("group-object"):
         for strs in p.find_children_w_parents(tmpSVCOBJGroup, 'group-object'):
               worksheet4.write(row, col + 1, groupname) 
               servicegroupobject = strs.replace("group-object", "").lstrip()
               worksheet4.write(row, col + 3, "1-65535")
               worksheet4.write(row, col + 4, 'Group ' + servicegroupobject)
               row += 1

              
    print ("Extracting Service Object Group Configuration : Completed")
 

##########################################
#                                        #
#   Extracting Access List               #
#                                        #
##########################################

    worksheet5 = workbook.add_worksheet('Access-List')
    row = 0
    col = 0

    worksheet5.write(row, col + 1, "Access List Name")
    worksheet5.write(row, col + 2, "Source")
    worksheet5.write(row, col + 3, "Destination")
    worksheet5.write(row, col + 4, "Service")
    worksheet5.write(row, col + 5, "Action")
    worksheet5.write(row, col + 6, "Remark")
    row += 1
    
    for parents in p.find_objects(r"^access-list"):
        tmp = parents.text.split()

        if 'remark' in tmp[2]:
			global remark
			remark = parents.text.split(' ', 3)
			worksheet5.write(row, col + 6, remark[3])
        
        if 'extended' in tmp[2]:
            if 'permit' in tmp[3]:
                worksheet5.write(row, col + 5, "Permit")
            elif 'deny' in tmp[3]:
                worksheet5.write(row, col + 5, "Deny")
            
            if "ip" in tmp[4]:
                if 'any' in tmp[5]:
                    worksheet5.write(row, col + 2, "Any")
                    if 'any' in tmp[6]:
                        worksheet5.write(row, col + 3, "Any")
                    elif ('object' or 'object-group') in tmp[6]:
                        worksheet5.write(row, col + 3, tmp[7])
                    elif 'host' in tmp[6]:
                        worksheet5.write(row, col + 3, tmp[7])
                    else:
                        worksheet5.write(row, col + 3, (tmp[6]+" "+tmp[7]))
                elif ('object' or 'object-group') in tmp[5]:
                    worksheet5.write(row, col + 2, tmp[6])
                    if 'any' in tmp[7]:
                        worksheet5.write(row, col + 3, "Any")
                    elif ('object' or 'object-group') in tmp[7]:
                        worksheet5.write(row, col + 3, tmp[8])
                    elif 'host' in tmp[7]:
                        worksheet5.write(row, col + 3, tmp[8])
                    else:
                        worksheet5.write(row, col + 3, (tmp[7]+" "+tmp[8]))
                elif 'host' in tmp[5]:
                    worksheet5.write(row, col + 2, tmp[6])
                    if 'any' in tmp[7]:
                        worksheet5.write(row, col + 3, "Any")
                    elif ('object' or 'object-group') in tmp[7]:
                        worksheet5.write(row, col + 3, tmp[8])
                    elif 'host' in tmp[7]:
                        worksheet5.write(row, col + 3, tmp[8])
                    else:
                        worksheet5.write(row, col + 3, (tmp[7]+" "+tmp[8]))
                else:
                    worksheet5.write(row, col + 2, (tmp[5]+" "+tmp[6]))
                    if 'any' in tmp[7]:
                        worksheet5.write(row, col + 3, "Any")
                    elif ('object' or 'object-group') in tmp[7]:
                        worksheet5.write(row, col + 3, tmp[8])
                    elif 'host' in tmp[7]:
                        worksheet5.write(row, col + 3, tmp[8])
                    else:
                        worksheet5.write(row, col + 3, (tmp[7]+" "+tmp[8]))
                worksheet5.write(row, col + 4, "Any")

            elif "icmp" in tmp[4]:
                if 'any' in tmp[5]:
                    worksheet5.write(row, col + 2, "Any")
                    if 'any' in tmp[6]:
                        worksheet5.write(row, col + 3, "Any")
                    elif ('object' or 'object-group') in tmp[6]:
                        worksheet5.write(row, col + 3, tmp[7])
                    elif 'host' in tmp[6]:
                        worksheet5.write(row, col + 3, tmp[7])
                    else:
                        worksheet5.write(row, col + 3, (tmp[6]+" "+tmp[7]))
                elif ('object' or 'object-group') in tmp[5]:
                    worksheet5.write(row, col + 2, tmp[6])
                    if 'any' in tmp[7]:
                        worksheet5.write(row, col + 3, "Any")
                    elif ('object' or 'object-group') in tmp[7]:
                        worksheet5.write(row, col + 3, tmp[8])
                    elif 'host' in tmp[7]:
                        worksheet5.write(row, col + 3, tmp[8])
                    else:
                        worksheet5.write(row, col + 3, (tmp[7]+" "+tmp[8]))
                elif 'host' in tmp[5]:
                    worksheet5.write(row, col + 2, tmp[6])
                    if 'any' in tmp[7]:
                        worksheet5.write(row, col + 3, "Any")
                    elif ('object' or 'object-group') in tmp[7]:
                        worksheet5.write(row, col + 3, tmp[8])
                    elif 'host' in tmp[7]:
                        worksheet5.write(row, col + 3, tmp[8])
                    else:
                        worksheet5.write(row, col + 3, (tmp[7]+" "+tmp[8]))
                else:
                    worksheet5.write(row, col + 2, (tmp[5]+" "+tmp[6]))
                    if 'any' in tmp[7]:
                        worksheet5.write(row, col + 3, "Any")
                    elif ('object' or 'object-group') in tmp[7]:
                        worksheet5.write(row, col + 3, tmp[8])
                    elif 'host' in tmp[7]:
                        worksheet5.write(row, col + 3, tmp[8])
                    else:
                        worksheet5.write(row, col + 3, (tmp[7]+" "+tmp[8]))
                worksheet5.write(row, col + 4, "ICMP")

            elif "tcp" in tmp[4]:
                if 'any' in tmp[5]:
                    worksheet5.write(row, col + 2, "Any")
                    if 'any' in tmp[6]:
                        worksheet5.write(row, col + 3, "Any")
                    elif ('object' or 'object-group') in tmp[6]:
                        worksheet5.write(row, col + 3, tmp[7])
                    elif 'host' in tmp[6]:
                        worksheet5.write(row, col + 3, tmp[7])
                    else:
                        worksheet5.write(row, col + 3, (tmp[6]+" "+tmp[7]))
                elif ('object' or 'object-group') in tmp[5]:
                    worksheet5.write(row, col + 2, tmp[6])
                    if 'any' in tmp[7]:
                        worksheet5.write(row, col + 4, tmp[9])
                        worksheet5.write(row, col + 3, "Any")
                    elif ('object' or 'object-group') in tmp[7]:
                        worksheet5.write(row, col + 3, tmp[8])
                        worksheet5.write(row, col + 4, tmp[10])
                    elif 'host' in tmp[7]:
                        worksheet5.write(row, col + 3, tmp[8])
                        worksheet5.write(row, col + 4, tmp[10])
                    else:
                        worksheet5.write(row, col + 3, (tmp[7]+" "+tmp[8]))
                        worksheet5.write(row, col + 4, tmp[10])
                elif 'host' in tmp[5]:
                    worksheet5.write(row, col + 2, tmp[6])
                    if 'any' in tmp[7]:
                        worksheet5.write(row, col + 4, tmp[9])
                        worksheet5.write(row, col + 3, "Any")
                    elif ('object' or 'object-group') in tmp[7]:
                        worksheet5.write(row, col + 3, tmp[8])
                        worksheet5.write(row, col + 4, tmp[10])
                    elif 'host' in tmp[7]:
                        worksheet5.write(row, col + 3, tmp[8])
                        worksheet5.write(row, col + 4, tmp[10])
                    else:
                        worksheet5.write(row, col + 3, (tmp[7]+" "+tmp[8]))
                        worksheet5.write(row, col + 4, tmp[10])
                else:
                    worksheet5.write(row, col + 2, (tmp[5]+" "+tmp[6]))
                    if 'any' in tmp[7]:
                        worksheet5.write(row, col + 3, "Any")
                    elif ('object' or 'object-group') in tmp[7]:
                        worksheet5.write(row, col + 3, tmp[8])
                        worksheet5.write(row, col + 4, tmp[10])
                    elif 'host' in tmp[7]:
                        worksheet5.write(row, col + 3, tmp[8])
                        worksheet5.write(row, col + 4, tmp[10])
                    else:
                        worksheet5.write(row, col + 3, (tmp[7]+" "+tmp[8]))
                        worksheet5.write(row, col + 4, tmp[10])

            elif "udp" in tmp[4]:
                if 'any' in tmp[5]:
                    worksheet5.write(row, col + 2, "Any")
                    if 'any' in tmp[6]:
                        worksheet5.write(row, col + 3, "Any")
                    elif ('object' or 'object-group') in tmp[6]:
                        worksheet5.write(row, col + 3, tmp[7])
                    elif 'host' in tmp[6]:
                        worksheet5.write(row, col + 3, tmp[7])
                    else:
                        worksheet5.write(row, col + 3, (tmp[6]+" "+tmp[7]))
                elif ('object' or 'object-group') in tmp[5]:
                    worksheet5.write(row, col + 2, tmp[6])
                    if 'any' in tmp[7]:
                        worksheet5.write(row, col + 4, tmp[9])
                        worksheet5.write(row, col + 3, "Any")
                    elif ('object' or 'object-group') in tmp[7]:
                        worksheet5.write(row, col + 3, tmp[8])
                        worksheet5.write(row, col + 4, tmp[10])
                        #print tmp[10]
                    elif 'host' in tmp[7]:
                        worksheet5.write(row, col + 3, tmp[8])
                        worksheet5.write(row, col + 4, tmp[10])
                        #print tmp[10]
                    else:
                        worksheet5.write(row, col + 3, (tmp[7]+" "+tmp[8]))
                        worksheet5.write(row, col + 4, tmp[10])
                        #print tmp[10]
                elif 'host' in tmp[5]:
                    worksheet5.write(row, col + 2, tmp[6])
                    if 'any' in tmp[7]:
                        worksheet5.write(row, col + 4, tmp[9])
                        worksheet5.write(row, col + 3, "Any")
                    elif ('object' or 'object-group') in tmp[7]:
                        worksheet5.write(row, col + 3, tmp[8])
                        worksheet5.write(row, col + 4, tmp[10])
                    elif 'host' in tmp[7]:
                        worksheet5.write(row, col + 3, tmp[8])
                        worksheet5.write(row, col + 4, tmp[10])
                        #print tmp[10]
                    else:
                        worksheet5.write(row, col + 3, (tmp[7]+" "+tmp[8]))
                        worksheet5.write(row, col + 4, tmp[10])
                        #print tmp[10]
                else:
                    worksheet5.write(row, col + 2, (tmp[5]+" "+tmp[6]))
                    if 'any' in tmp[7]:
                        worksheet5.write(row, col + 3, "Any")
                    elif ('object' or 'object-group') in tmp[7]:
                        worksheet5.write(row, col + 3, tmp[8])
                        worksheet5.write(row, col + 4, tmp[10])
                    elif 'host' in tmp[7]:
                        worksheet5.write(row, col + 3, tmp[8])
                        worksheet5.write(row, col + 4, tmp[10])
                    else:
                        worksheet5.write(row, col + 3, (tmp[7]+" "+tmp[8]))
                        worksheet5.write(row, col + 4, tmp[10])

                if 'icmp' in tmp[4]:
                    worksheet5.write(row, col + 4, "ICMP")
                else:
                    worksheet5.write(row, col + 4, "Any")
                
            elif ('object' or 'object-group') in tmp[4]:
                worksheet5.write(row, col + 4, tmp[5])
                if 'any' in tmp[6]:
                    worksheet5.write(row, col + 2, "Any")
                    if 'any' in tmp[7]:
                        worksheet5.write(row, col + 3, "Any")
                    elif ('object' or 'object-group') in tmp[7]:
                        worksheet5.write(row, col + 3, tmp[8])
                    elif 'host' in tmp[7]:
                        worksheet5.write(row, col + 3, tmp[8])
                    else:
                        worksheet5.write(row, col + 3, (tmp[7]+" "+tmp[8]))

                elif ('object' or 'object-group') in tmp[6]:
                    worksheet5.write(row, col + 2, tmp[7])
                    if 'any' in tmp[8]:
                        worksheet5.write(row, col + 3, "Any")
                    elif ('object' or 'object-group') in tmp[8]:
                        worksheet5.write(row, col + 3, tmp[9])
                    elif 'host' in tmp[8]:
                        worksheet5.write(row, col + 3, tmp[9])
                    else:
                        worksheet5.write(row, col + 3, (tmp[8]+" "+tmp[9]))

                elif 'host' in tmp[6]:
                    worksheet5.write(row, col + 2, tmp[7])
                    if 'any' in tmp[8]:
                        worksheet5.write(row, col + 4, tmp[9])
                        worksheet5.write(row, col + 3, "Any")
                    elif ('object' or 'object-group') in tmp[8]:
                        worksheet5.write(row, col + 3, tmp[9])
                    elif 'host' in tmp[8]:
                        worksheet5.write(row, col + 3, tmp[9])
                        #print tmp[10]
                    else:
                        worksheet5.write(row, col + 3, (tmp[9]+" "+tmp[10]))
                        #print tmp[10]

                else:
                    worksheet5.write(row, col + 2, (tmp[6]+" "+tmp[7]))
                    if 'any' in tmp[8]:
                        worksheet5.write(row, col + 3, "Any")
                    elif ('object' or 'object-group') in tmp[8]:
                        worksheet5.write(row, col + 3, tmp[9])
                    elif 'host' in tmp[7]:
                        worksheet5.write(row, col + 3, tmp[9])
                    else:
                        worksheet5.write(row, col + 3, (tmp[8]+" "+tmp[9]))

            worksheet5.write(row, col + 1, tmp[1])
            row += 1
            
    print ("Extracting Access List Configuration : Completed")

##########################################
#                                        #
#   Extracting Interface Configuration   #
#                                        #
##########################################

    worksheet6 = workbook.add_worksheet('Interface Configuration')
    row = 0
    col = 0

    worksheet6.write(row, col + 1, "Interface Name")
    worksheet6.write(row, col + 2, "Interface Type")
    worksheet6.write(row, col + 3, "IP Address")
    worksheet6.write(row, col + 4, "Subnet Mask")
    worksheet6.write(row, col + 5, "Standby IP")
    worksheet6.write(row, col + 6, "Security Level")
    worksheet6.write(row, col + 7, "Remark")
    
    row += 1

    for parents in p.find_objects(r"^interface"):
        tmp = parents.text
        #print (tmp)
        worksheet6.write(row, col + 1, tmp.replace("interface", "").lstrip())

        if parents.re_search_children("no ip address"):
            #print (" ip address none")
            worksheet6.write(row, col + 3, "None")
        elif parents.re_search_children("ip address"):           
            for child in p.find_children_w_parents("^%s$" % tmp, 'ip address', ignore_ws=True):
                addrs = child.split()
                ipaddress = addrs[2]
                netmask = addrs[3]
#                VirtualIP = addrs[5]
                
                worksheet6.write(row, col + 3, ipaddress)
                worksheet6.write(row, col + 4, netmask)
#                worksheet6.write(row, col + 5, VirtualIP)

        if parents.re_search_children("nameif"):
             for child in p.find_children_w_parents(tmp, "nameif"):
                nameif = child.replace("nameif", "").lstrip()
                worksheet6.write(row, col + 2, nameif)

        if parents.re_search_children("no security-level"):
            worksheet6.write(row, col + 6, "None")
        elif parents.re_search_children("security-level"):
             for child in p.find_children_w_parents(tmp, "security-level"):
                worksheet6.write(row, col + 6, child.replace("security-level", "").lstrip())
                                
        row += 1
    print ("Extracting Interface Configuration : Completed")
 
    workbook.close()
    print ("All Task Completed")
    print ("Please find your xlsx file in folder")

def main():
	#Main Program
	parser = argparse.ArgumentParser(description='Process filename.')
	parser.add_argument('file', help="Put file name here, include the file extension")
	args = parser.parse_args()
	fileName=args.file
	ConfigParser(fileName) 

if __name__ == "__main__":
	#If this Python file runs by itself, run below command. If imported, this section is not run
	main()      
    
