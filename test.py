import gspread

service_account = gspread.service_account(filename='gs_account.json')
sheet = service_account.open('Ai workflow')
orders_task_list = sheet.worksheet('OrdersTaskList')
orders = sheet.worksheet('Orders')

orders_default = {'Quotation': '', 'Customer': '', 'Order': '', 'Order price': '', 'Delivery': '', 'Delivery Date': '', 'Invoice': '', 'Invoice Date': '', 'Invoice price': '', 'Received': '', 'Notes': ''}
print(orders.get_all_records())