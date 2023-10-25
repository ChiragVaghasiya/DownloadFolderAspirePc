import xmlrpc.client as xmlrpclib
import datetime,xlrd
################################
url = 'http://localhost:8116'
dbname = 'test'
username = 'info@s-business.ch'
pwd = 'admin'

sock_common = xmlrpclib.ServerProxy (url+'/xmlrpc/common')
sock = xmlrpclib.ServerProxy(url+'/xmlrpc/object', allow_none=True)
uid = sock_common.login(dbname, username, pwd)

workbook = xlrd.open_workbook('sale_order.xlsx')
worksheet = workbook.sheet_by_name('Sheet1')
num_rows = worksheet.nrows - 1
num_cells = worksheet.ncols - 1
curr_row = 0

count = 0
while curr_row < num_rows:
    curr_row += 1
    print("current row", curr_row)
    row = worksheet.row(curr_row)

    order_number = row[1].value
    project_name = row[2].value.strip()
    print(project_name, project_name)
    
    #Lausanne  -> Belmont-sur
    if project_name:
        project_id = sock.execute(dbname, uid, pwd, 'project.project', 'search', [('test_name','like', project_name)])
        if not project_id:
            # project_id = todo
            print(project_id)
        else:
            project_id = project_id[0]
        print ("******************************************",project_id)
        order_id = sock.execute(dbname, uid, pwd, 'sale.order', 'search', [('name','=', order_number)])
        print(order_id)
        task_vals = {'name': "test 111", 'stage_id': 1, 'project_id': project_id}
        task_id = sock.execute(dbname, uid, pwd, 'project.task', 'create', task_vals)
        sock.execute(dbname, uid, pwd, 'sale.order', 'write', order_id, {'project_id': project_id, 'task_id': task_id})
        break
        # if product_id:
        #     external_id = sock.execute(dbname, uid, pwd, 'ir.model.data', 'search', [('model','=', 'product.template'),('res_id','=', product_id[0])])
        #     if external_id:
        #         print ("update.............", external_id)
print ('FINISH')
