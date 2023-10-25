

import xlrd
import xmlrpc.client as xmlrpclib

url = 'http://38.242.129.6:8073'
dbname = 'ELL_construction'
username = 'admin'
pwd = 'admin'

sock_common = xmlrpclib.ServerProxy (url+'/xmlrpc/common')
sock = xmlrpclib.ServerProxy(url+'/xmlrpc/object')
uid = sock_common.login(dbname, username, pwd)

workbook = xlrd.open_workbook('/home/cravit6/Downloads/to import in ell construction database.xlsx')
worksheet = workbook.sheet_by_name('export_contacts_1674560181')
num_rows = worksheet.nrows - 1
num_cells = worksheet.ncols - 1
curr_row = 0
count = 0


while curr_row < num_rows:
    curr_row += 1
    row = worksheet.row(curr_row)
    country_id = sock.execute(dbname, uid, pwd,'res.country', 'search', [('name','ilike',row[7].value)])
    tite_id = sock.execute(dbname, uid, pwd,'res.partner.title', 'search', [('name','ilike',row[3].value)])
    # lang_id = sock.execute(dbname, uid, pwd,'res.lang', 'search', [('name','like',row[3].value)])

    if row[1].value:
        company_type="company"
        name = row[1].value
    if row[2].value:
        company_type="person"
        name = row[2].value
    vals={'company_type':company_type,'name':name, 'title': tite_id[0] or False, 'street': row[4].value or "", 'zip': str(row[5].value) or "", 
           'city': row[6].value, 'country_id': country_id[0] or False, 'email': row[8].value or '', 'phone': str(row[9].value) or '', 'mobile': str(row[10].value) or '', 
           'website': row[11].value or '', 'lang': 'fr_FR' or ''}
    sock.execute(dbname, uid, pwd, 'res.partner', 'create', vals)    
print('~~~~~~~~~~~~Finesh~~~~~~~~~~~~~~~~~~~~~~')
