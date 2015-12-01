#!/usr/bin/env python

import xlsxwriter
import cobra.mit.access
import cobra.mit.session
import cobra.mit.request
from cobra.mit.request import DnQuery, ClassQuery
import yaml
import httplib

#httplib.HTTPConnection.debuglevel = 1
tabs = []

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


def fabricPathEpResolver(md,mac):
    c = ClassQuery('fvCEp')
    c.propFilter='and(eq(fvCEp.mac, "%s"))' % mac
    dn = str(md.query(c)[0].dn)
    q=DnQuery(dn)
    q.queryTarget = 'children'
    q.subtree = 'full'
    q.classFilter='fvRsCEpToPathEp'
    children = md.query(q)
    plist = []

    for o in children:
        plist.append(o.tDn)

    return plist




def createWorkSheet(md, workbook, sheetname, cls, columns):
    outline = workbook.add_format()
    outline.set_border()
    sheet = new_worksheet(workbook, str(sheetname))
    row,col = 0,0
    for f in columns:
        sheet.write(row,col,f,outline)
        col += 1
    row,col = 1,0
    mos = md.lookupByClass(cls)
    for i in mos:
        l = [str(getattr(i, s)) for s in columns]
        for i in l:
            sheet.write(row,col,i,outline)
            col += 1
        row += 1
        col = 0

def createTenantSheet(md, workbook,t):
    # creates a worksheet for a Tenant
    bold = workbook.add_format({'bold': 1})
    outline = workbook.add_format()
    outline.set_border()
    sheet = new_worksheet(workbook,str(t.name))
    row ,col = 0,0
    headers = ['Application','EPGs','Bridge-Domain','Endpoints']
    for h in headers:
        sheet.write(row,col,h,bold)
        col +=1
    row,col = 1,0
    for ap in md.lookupByClass('fvAp',parentDn=t.dn):
        sheet.write(row,col,ap.name,outline)
        col += 1
        for epg in md.lookupByClass('fvAEPg',parentDn=ap.dn):
            sheet.write(row,col,epg.name,outline)
            sheet.write(0,col,ap.name,outline)
            bd = md.lookupByClass('fvRsBd',parentDn=epg.dn)
            sheet.write(row, 0, ap.name,outline)
            sheet.write(row, col+1, bd[0].tRn,outline)
            row += 1

            '''
            epc = col + 2
            for cep in md.lookupByClass('fvCEp', parentDn=epg.dn):
                sheet.write(row, 0, ap.name,outline)
                sheet.write(row, 1, epg.name,outline)
                sheet.write(row, 2, bd[0].tRn,outline)
                sheet.write(row, epc, cep.ip,outline)
                sheet.write(row, epc+1, cep.mac,outline)
                epc += 1
                for p in fabricPathEpResolver(md,cep.mac):
                    sheet.write(row, epc+2, p, outline)
                    epc += 1
                row += 1
            '''
        col -= 1

def CreateWorkBook(md,xls,tabs):
    workbook = xlsxwriter.Workbook(xls)
    for k in tabs:
        createWorkSheet(md,workbook, k,tabs[k]['class'], tabs[k]['columns'])
    for t in md.lookupByClass('fvTenant', parentDn='uni'):
        createTenantSheet(md,workbook,t)
    workbook.close()


with open('config.yaml', 'r') as config:
    config = yaml.safe_load(config)
hostname, username, password = config['host'], config['name'], config['passwd']
ls = cobra.mit.session.LoginSession(hostname, username, password)
md = cobra.mit.access.MoDirectory(ls)
md.login()
CreateWorkBook(md, config['filename'], config['tabs'])
