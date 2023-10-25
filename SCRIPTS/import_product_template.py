#!/usr/bin/env python
# coding: utf-8
from xmlrpc import client
import datetime,xlrd
url = 'http://localhost:8115'
dbname = 'odoo_v15'
username = 'admin'
pwd = 'admin'

sock_common = client.ServerProxy (url+'/xmlrpc/common')
sock = client.ServerProxy(url+'/xmlrpc/object')
uid = sock_common.login(dbname, username, pwd)

workbook = xlrd.open_workbook('a.xlsx')
worksheet = workbook.sheet_by_name('Table 1')
num_rows = worksheet.nrows - 1
num_cells = worksheet.ncols - 1
curr_row = 1

count = 0
while curr_row < num_rows:
    curr_row += 1

    row = worksheet.row(curr_row)
    address = row[2].value.split(',')
    print(address)
    
    add_vals = {
                'house': address[0],
                'stree': address[1] ,
                'stree2': address[2],
                'city': address[3],
                'state': address[4],
                ''
    }
    # partner_id = sock.execute(dbname, uid, pwd, 'res.partner', 'search', [('name','=',row[0].value +" " + row[1].value)])
    # if not partner_id:
    #     vals = {
    #             'name' : row[0].value +" " + row[1].value,
    #             'street': address[0]', 'address[1],
    #             'street2': address[2],
    #             'city': address[3],
    #             # 'state_id':,
    #             'zip': address,
    #             # 'country_id':,
    #             }
    #     print(vals)
    #     product_id = sock.execute(dbname, uid, pwd,'res.partner', 'create', vals)
    #     print(partner_id)
    # else:
    #     sock.execute(dbname, uid, pwd, 'res.partner', 'write', partner_id,{'name': row[0].value +" " + row[1].value})  
# print (row)


# product_ids = sock.execute(dbname, uid, pwd, 'product.template', 'search', [('barcode','!=',False)])
# count = len(product_ids)
# print(count)
# for product_id in product_ids:
#     count -= 1
#     print(count)
#     barcode = sock.execute(dbname, uid, pwd, 'product.template', 'read', [product_id], ['barcode'])
#     update_barcode = barcode[0].get('barcode').split('.')[0]
#     search_barcode = sock.execute(dbname, uid, pwd, 'product.template', 'search', [('barcode','=',update_barcode)])
#     if search_barcode:
#         print("skip product to update due to duplicate barcode",search_barcode, update_barcode)
#         continue
#     sock.execute(dbname, uid, pwd, 'product.template', 'write', product_id,{'barcode': update_barcode})
print ("FINISH")

