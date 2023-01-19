import setWorkspace
import re
import gspread
import pyperclip as pc
from time import sleep
from pyrobogui import robo, pag
from itertools import zip_longest


service_account = gspread.service_account(filename='gs_account.json')
sheet = service_account.open('Ai workflow')
orders_task_list = sheet.worksheet('OrdersTaskList')
orders = sheet.worksheet('Orders')

orders_task_list_default = {'Created by': '', 'Create DateTime': '', 'Customer Name': '', 'Customer': '', 'Quotation': '', 'Need Approval': '', 'Brand Managers': '', 'Financial': '', 'Approved': '', 'Creditlimit': '', 'Branch Manager': '', 'CL Financial': '', 'CL Approved': '', 'Finished Date': ''}
orders_default = {'Quotation': '', 'Customer': '', 'Order': '', 'Order price': '', 'Delivery': '', 'Delivery Date': '', 'Invoice': '', 'Invoice Date': '', 'Invoice price': '', 'Received': '', 'Notes': ''}

def sap_response_time():
    robo.waitImageToAppear(image='./images/getRT_position.png')
    ms_position = pag.locateOnScreen(image='./images/getRT_position.png')
    while True:
        if not pag.locateOnScreen(image='./images/lastRT.png'):
            if not pag.locateOnScreen(image='./images/zeroRT.png'):
                pag.screenshot('./images/lastRT.png', region=(ms_position.left-150, ms_position.top, 185, 25))
                break

def check_items():
    while True:
        if pag.locateOnScreen(image='./images/information.png'):
            robo.click(image='./images/greenEnter.png')
            sap_response_time()
        elif pag.locateOnScreen(image='./images/delivery_proposal.png'):
            robo.click(image='./images/delivery_proposal.png')
            sap_response_time()
        elif pag.locateOnScreen(image='./images/continue.png'):
            robo.click(image='./images/continue.png')
            sap_response_time()
        else:
            break

def get_quote_status():
    with open('temp.txt', 'r', encoding='utf-8') as file:
        lines = file.readlines()
        customer_name = " ".join(lines[1].split()[2:])
        customer = lines[1].split()[1]
        q_state = lines[5].split()[7]
        return customer_name, customer, q_state

def update_orders_task_list():
    all_orders_task = orders_task_list.get_all_records()
    quotations = [d["Quotation"] for d in all_orders_task]
    for q in quotations:
        print(q)
    pass

print(orders_task_list.find('52118928').row)
update_orders_task_list()


# quotations = orders_task_list.get_values('E2:F')
# status = orders_task_list.get_values('I2:I')
# new_list = [[x[0], x[1], y[0] if y else ''] for x, y in zip_longest(quotations, status, fillvalue=[''])]
# print(new_list)




# try:
#     quotations = orders_task_list.get_values('E2:E')
#     for quotation in quotations:
#         storage_location = None
#         sales_office = None
#         robo.click(image='./images/command_box.png')
#         pag.typewrite('/N VA22\n')
#         sap_response_time()
#         pag.typewrite(quotation[0]+'\n')
#         sap_response_time()
#         while True:
#             if pag.locateOnScreen(image='./images/makkah.png', confidence=0.7) is not None:
#                 sales_office = 'S104'
#                 print(sales_office)
#                 break
#         while True:
#             if pag.locateOnScreen(image='./images/mmwm.png', confidence=0.7) is not None:
#                 storage_location = 'BN1'
#                 print(storage_location)
#                 break
    
#         robo.click(image='./images/sales_document.png')
#         robo.click(image='./images/create_subsequent_order.png')
#         sap_response_time()
#         robo.click(image='./images/sales_office.png')
#         pag.hotkey('tab')
#         pag.typewrite(sales_office)
#         pag.hotkey('tab')
#         pag.typewrite(storage_location+'\n')
#         sap_response_time()
#         robo.click(image='./images/copy.png')
#         sap_response_time()
#         check_items()
#         robo.click(image='./images/sales_document.png')
#         robo.click(image='./images/deliver.png')
#         sap_response_time()
#         if pag.locateOnScreen(image='./images/create_delivery_order.png'):
#             while True:
#                 pag.hotkey('enter')
#                 sap_response_time
#                 if not pag.locateOnScreen(image='.images/create_delivery_order.png'):
#                     break
#                 else:
#                     sleep(1)
#         robo.click(image='./images/hat.png')
#         sap_response_time()
#         if pag.locateOnScreen(image='./images/administration.png'):
#             robo.click(image='./images/administration.png')
#             sap_response_time()
#         pc.copy('تحضير وارسال فوري')
#         pag.hotkey('ctrl', 'v')
#         robo.click(image='./images/save.png')
#         sap_response_time()
#         pag.hotkey('ctrl', 'c')
#         order = pc.paste()
#         robo.doubleClick(image='./images/has_been_saved.png')
#         sap_response_time()
#         robo.click(image='./images/bold_delivery.png')
#         pag.tripleClick()
#         pag.hotkey('ctrl', 'c')
#         delivery = pc.paste()
#         delivery = re.findall(r'\d+', delivery)
#         delivery = delivery[0]
#         print(quotation[0])

# except gspread.exceptions.APIError:
#     print('No quotation')
    



        


