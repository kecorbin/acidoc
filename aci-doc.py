#!/usr/bin/env python

import xlsxwriter
from acitoolkit import Session, Credentials

import yaml
import httplib

#httplib.HTTPConnection.debuglevel = 1
tabs = []

def class_query(session, kls, parentDn=None):
    """
    query ACI based on class
    """

    if parentDn:
        url = '/api/mo/{}.json?query-target=children&target-subtree-class={}'.format(parentDn, kls)
    else:
        url = '/api/class/{}.json'.format(kls)
    resp = session.get(url)
    return resp.json()['imdata']

def new_worksheet(workbook,name):
    '''
    Verify the sheet name is compatible with xlsx writer, tracks sheet names for case based duplicates
    '''
    global tabs
    name = name[0:31]
    while True:
        if name.lower() in tabs:
            name += '-'
        else:
            tabs.append(name.lower())
            sheet = workbook.add_worksheet(name)
            break
    return sheet

def createWorkSheet(session, workbook, sheetname, kls, properties, headers=None):
    outline = workbook.add_format()
    outline.set_border()
    sheet = new_worksheet(workbook, str(sheetname))
    print("Collecting information for {}".format(sheetname))
    row,col = 0,0

    if headers is None:
        headers = properties

    bold = workbook.add_format({'bold': 1})
    for h in headers:
        sheet.write(row,col,h,bold)
        col +=1
    row,col = 1,0

    for f in properties:
        sheet.write(row,col,f,outline)
        col += 1
    row,col = 1,0
    mos = class_query(session, kls)
    for i in mos:
        vals = [str(i[kls]['attributes'][s]) for s in properties ]
        for v in vals:
            sheet.write(row,col,v,outline)
            col += 1
        row += 1
        col = 0

def createTenantSheet(session, workbook,t):
    """
    creates a tab in workbook for tenant t using acitoolkit session
    """
    bold = workbook.add_format({'bold': 1})
    outline = workbook.add_format()
    outline.set_border()
    sheet = new_worksheet(workbook,str(t['fvTenant']['attributes']['name']))
    row ,col = 0,0
    headers = ['Application','EPGs','Bridge-Domain','Endpoints']
    for h in headers:
        sheet.write(row,col,h,bold)
        col +=1
    row,col = 1,0

    for ap in class_query(session, 'fvAp', parentDn=t['fvTenant']['attributes']['dn']):
        app_name = ap['fvAp']['attributes']['name']
        app_dn = ap['fvAp']['attributes']['dn']
        sheet.write(row,col, app_name, outline)
        col += 1
        for epg in class_query(session, 'fvAEPg', parentDn=app_dn):
            epg_dn = epg['fvAEPg']['attributes']['dn']
            epg_name = epg['fvAEPg']['attributes']['name']
            sheet.write(row, col, epg_name, outline)
            sheet.write(0, col, app_name, outline)
            bd = class_query(session, 'fvRsBd', parentDn=epg_dn)
            bd_name = bd[0]['fvRsBd']['attributes']['tRn']
            sheet.write(row, 0, app_name, outline)
            sheet.write(row, col+1, bd_name, outline)
            row += 1
        col -= 1

def CreateWorkBook(session, xls, tabs):
    workbook = xlsxwriter.Workbook(xls)
    for k in tabs:
        createWorkSheet(session,
                        workbook,
                        k,
                        tabs[k]['class'],
                        tabs[k]['properties'],
                        headers=tabs.get(k).get('headers'))

    tenants = class_query(session, 'fvTenant')
    for t in tenants:
        createTenantSheet(session,workbook,t)
    workbook.close()


if __name__ == "__main__":

    description = 'aci-doc'

    # Gather credentials for ACI
    creds = Credentials('apic', description)
    args = creds.get()

    # Establish an API session to the APIC
    apic = Session(args.url, args.login, args.password)

    if apic.login().ok:
        print("Connected to ACI")

    print("depending on your configuration, this could take a little while...")

    with open('config.yaml', 'r') as config:
        config = yaml.safe_load(config)

    CreateWorkBook(apic, config['filename'], config['tabs'])
