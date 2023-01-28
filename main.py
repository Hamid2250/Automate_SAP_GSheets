import setWorkspace
import re
import os
import datetime
import PyPDF2
import gspread
import pyperclip as pc
from time import sleep, time
from pyrobogui import robo, pag
from itertools import zip_longest


service_account = gspread.service_account(filename='gs_account.json')
sheet = service_account.open('Ai workflow')
orders_task_list = sheet.worksheet('OrdersTaskList')
orders = sheet.worksheet('Orders')

orders_task_list_default = {'Created by': '', 'Create DateTime': '', 'Customer Name': '', 'Customer': '', 'Quotation': '', 'Need Approval': '', 'Brand Managers': '', 'Financial': '', 'Approved': '', 'Creditlimit': '', 'Branch Manager': '', 'CL Financial': '', 'CL Approved': '', 'Finished Date': ''}
orders_default = {'Quotation': '', 'Customer': '', 'Order': '', 'Order price': '', 'Delivery': '', 'Delivery Date': '', 'Invoice': '', 'Invoice Date': '', 'Invoice price': '', 'Received': '', 'Notes': ''}

def sap_response_time():
    ms_position = pag.locateOnScreen(image='./images/getRT_position.png')
    start_time = time()
    while True:
        pix = pag.pixel(x=39, y=37)
        if not pag.locateOnScreen(image='./images/lastRT.png') and not pag.locateOnScreen(image='./images/zeroRT.png'):
            pag.screenshot('./images/lastRT.png', region=(ms_position.left-150, ms_position.top, 185, 25))
            break
        elif pag.locateOnScreen(image='./images/zeroRT.png'):
            start_time = time()
        elif pix != (242, 242, 242):
            start_time = time()
            break
        elif time() - start_time > 5:
            print('timeout')
            break
        else:
            sleep(1)

def check_items():
    while True:
        if pag.locateOnScreen(image='./images/information.png'):
            robo.doubleClick(image='./images/greenEnter.png')
            sap_response_time()
        elif pag.locateOnScreen(image='./images/delivery_proposal.png'):
            robo.doubleClick(image='./images/delivery_proposal.png')
            sap_response_time()
        elif pag.locateOnScreen(image='./images/continue.png'):
            robo.doubleClick(image='./images/continue.png')
            sap_response_time()
        elif pag.locateOnScreen(image='./images/sales_document.png'):
            break

def get_quote_status():
    with open('temp.txt', 'r', encoding='utf-8') as file:
        lines = file.readlines()
        customer_name = " ".join(lines[1].split()[2:])
        customer = lines[1].split()[1]
        q_state = lines[5].split()[7]
        return customer_name, customer, q_state

def update_orders_task_list():
    # all_orders_task = orders_task_list.get_all_records()
    quotations = [d["Quotation"] for d in all_orders_task]
    for q in quotations:
        if all_orders_task[quotations.index(q)]["Finished Date"] == "":
            if all_orders_task[quotations.index(q)]["Need Approval"] == "" or all_orders_task[quotations.index(q)]["Approved"] == "":
                robo.click(image='./images/command_box.png')
                pag.typewrite('/N VA22\n')
                sap_response_time()
                pag.typewrite(str(q))
                sap_response_time()
                robo.click(image='./images/status_overview.png')
                sap_response_time()
                if pag.locateOnScreen(image='./images/information.png'):
                    robo.click(image='./images/greenEnter.png')
                    sap_response_time()
                pag.hotkey('shift', 'f8')
                sap_response_time()
                robo.click(image='./images/text_with_tabs.png')
                robo.click(image='./images/greenEnter.png')
                sap_response_time()
                pag.hotkey('shift', 'tab')
                pag.typewrite(os.getcwd())
                pag.hotkey('tab')
                pag.typewrite('temp.txt')
                pag.hotkey('tab')
                pag.typewrite('4110')
                pag.hotkey('ctrl', 's')
                sap_response_time()
                robo.click(image='./images/command_box.png')
                pag.typewrite('/N\n')
                sap_response_time()
                current_row = orders_task_list.find(str(q)).row
                info = get_quote_status()
                current_task = all_orders_task[quotations.index(q)]
                if info[2] == 'Q000':
                    update_task = {'Customer Name': info[0], 'Customer': info[1], 'Need Approval': 'NO',}
                    update_task = dict(current_task, **update_task)
                    update_task = list(update_task.values())
                    orders_task_list.batch_update([{'range': f'A{current_row}:N{current_row}', 'values': [update_task]}])
                    all_orders_task[quotations.index(q)]['Customer Name'], all_orders_task[quotations.index(q)]['Customer'], all_orders_task[quotations.index(q)]['Need Approval'] = info[0], info[1], 'NO'
                elif info[2] == 'Q004':
                    update_task = {'Customer Name': info[0], 'Customer': info[1], 'Need Approval': 'YES', 'Approved': 'YES'}
                    update_task = dict(current_task, **update_task)
                    update_task = list(update_task.values())
                    orders_task_list.batch_update([{'range': f'A{current_row}:N{current_row}', 'values': [update_task]}])
                    all_orders_task[quotations.index(q)]['Customer Name'], all_orders_task[quotations.index(q)]['Customer'], all_orders_task[quotations.index(q)]['Need Approval'], all_orders_task[quotations.index(q)]['Approved'] = info[0], info[1], 'YES', 'YES'
                elif info[2] == 'Q002' or info[2] == 'Q003':
                    update_task = {'Customer Name': info[0], 'Customer': info[1], 'Need Approval': 'YES',}
                    update_task = dict(current_task, **update_task)
                    update_task = list(update_task.values())
                    orders_task_list.batch_update([{'range': f'A{current_row}:N{current_row}', 'values': [update_task]}])
                    all_orders_task[quotations.index(q)]['Customer Name'], all_orders_task[quotations.index(q)]['Customer'], all_orders_task[quotations.index(q)]['Need Approval'] = info[0], info[1], 'YES'

def transfer_quotations():
    # all_orders_task = orders_task_list.get_all_records()
    quotations = [d["Quotation"] for d in all_orders_task]
    for q in quotations:
        if all_orders_task[quotations.index(q)]["Finished Date"] == "":
            if all_orders_task[quotations.index(q)]["Need Approval"] == "NO" or all_orders_task[quotations.index(q)]["Approved"] == "YES":
                storage_location = None
                sales_office = None
                robo.click(image='./images/command_box.png')
                pag.typewrite('/N VA22\n')
                sap_response_time()
                pag.typewrite(str(q)+'\n')
                sap_response_time()
                while True:
                    if pag.locateOnScreen(image='./images/makkah.png') is not None:
                        sales_office = 'S104'
                        break
                while True:
                    if pag.locateOnScreen(image='./images/mmwm.png') is not None:
                        storage_location = 'BN1'
                        break
                    elif pag.locateOnScreen(image='./images/sm01.png') is not None:
                        storage_location = 'SM1'
                        break
                
                # Transfer To Order
                robo.click(image='./images/sales_document.png')
                robo.click(image='./images/create_subsequent_order.png')
                sap_response_time()
                robo.click(image='./images/sales_office.png')
                pag.hotkey('tab')
                pag.typewrite(sales_office)
                pag.hotkey('tab')
                pag.typewrite(storage_location+'\n')
                sap_response_time()
                robo.click(image='./images/copy.png')
                sap_response_time()
                check_items()

                # Transfer To Delivery
                robo.click(image='./images/sales_document.png')
                robo.click(image='./images/deliver.png')
                sap_response_time()
                if pag.locateOnScreen(image='./images/create_delivery_with_order.png'):
                    while True:
                        pag.hotkey('enter')
                        sap_response_time
                        if not pag.locateOnScreen(image='./images/create_delivery_with_order.png'):
                            break
                        else:
                            sleep(1)
                robo.waitImageToAppear(image='./images/hat.png', full_match=True)
                robo.click(image='./images/hat.png', full_match=True)
                sap_response_time()
                if pag.locateOnScreen(image='./images/administration.png'):
                    robo.click(image='./images/administration.png')
                    sap_response_time()
                pc.copy('تحضير وارسال فوري')
                pag.hotkey('ctrl', 'v')
                robo.click(image='./images/save.png')
                sap_response_time()
                robo.click(image='./images/command_box.png')
                pag.typewrite('/N\n')
                sap_response_time()

                # Transfer To Invoice
                if storage_location == 'SM1':
                    robo.click(image='./images/command_box.png')
                    pag.typewrite('/N VL03N\n')
                    sap_response_time()
                    robo.click(image='./images/outbound_delivery.png')
                    pag.hotkey('tab')
                    pag.hotkey('ctrl', 'c')
                    delivery = pc.paste()
                    robo.click(image='./images/command_box.png')
                    pag.typewrite('/N VL06G\n')
                    sap_response_time()
                    robo.click(image='./images/execute.png')
                    sap_response_time()
                    robo.click(image='./images/delivery_blue_bg.png')
                    robo.click(image='./images/filter.png')
                    sap_response_time()
                    pag.typewrite(str(delivery))
                    robo.click(image='./images/greenEnter.png')
                    sap_response_time()
                    robo.click(image='./images/delivery_check_box.png')
                    robo.click(image='./images/post_goods_issue.png')
                    robo.click(image='./images/greenEnter.png')
                    sap_response_time()
                    robo.click(image='./images/command_box.png')
                    pag.typewrite('/N VF01\n')
                    sap_response_time()
                    pag.hotkey('enter')
                    sap_response_time()
                    robo.click(image='./images/save.png')
                    sap_response_time()
                
                # Update in Google Sheet
                current_row = orders_task_list.find(str(q)).row
                now = datetime.datetime.now()
                current_task = all_orders_task[quotations.index(q)]
                finished_date = now.strftime("%d/%m/%Y")
                update_task = {"Finished Date" : finished_date}
                update_task = dict(current_task, **update_task)
                update_task = list(update_task.values())
                orders_task_list.batch_update([{'range': f'A{current_row}:N{current_row}', 'values': [update_task]}])
                all_orders_task[quotations.index(q)]['Finished Date'] = finished_date

def pdf_to_txt(f, pdf_file):
    pdf = PyPDF2.PdfReader(f)
    with open(pdf_file.replace('.pdf', '.txt'), 'w',encoding='utf-8') as f:
        for page in range(len(pdf.pages)):
            text = pdf.pages[page].extract_text()
            f.write(text)

def update_orders_from_orders_tasks_list():
    # all_orders_task = orders_task_list.get_all_records()
    all_orders_task = [{'Created by': '', 'Create DateTime': '', 'Customer Name': '', 'Customer': '', 'Quotation': 52118966, 'Need Approval': '', 'Brand Managers': '', 'Financial': '', 'Approved': '', 'Creditlimit': '', 'Branch Manager': '', 'CL Financial': '', 'CL Approved': '', 'Finished Date': '25/01/2023'}]
    quotations = [d["Quotation"] for d in all_orders_task]
    for q in quotations:
        if all_orders_task[quotations.index(q)]["Finished Date"] != "":
            robo.click(image='./images/command_box.png')
            pag.typewrite('/N VA22\n')
            sap_response_time()
            pag.typewrite(str(q))
            sap_response_time()
            robo.click(image='./images/document_flow.png')
            sap_response_time()
            pag.click(pag.locateCenterOnScreen(image='./images/print_view.png'), interval=0.2, clicks=2)
            sap_response_time()
            robo.waitImageToAppear(image='./images/greenEnter.png')
            pag.typewrite('LOCAL')
            if not pag.locateOnScreen(image='./images/immediately.png'):
                robo.click(image='./images/greenEnter.png', offsetUp=60, offsetLeft=10)
                robo.waitImageToAppear(image='./images/immediately.png')
                sleep(0.1)
                pag.hotkey('down')
                pag.hotkey('enter')
                # robo.click(image='./images/immediately.png')
                sleep(0.1)
            robo.click(image='./images/greenEnter.png')
            sap_response_time()
            if not pag.locateOnScreen(image='./images/print_to_pdf_selected.png'):
                robo.click(image='./images/printer_drop_list.png')
                robo.click(image='./images/print_to_pdf.png')
            pag.hotkey('enter')
            robo.waitImageToAppear(image='./images/file_name.png')
            pag.typewrite(os.getcwd()+'\\temp.pdf')
            pag.hotkey('enter')
            sleep(2)
            with open('./temp.pdf', 'rb') as f:
                pdf_to_txt(f, './temp.pdf')

# update_orders_from_orders_tasks_list()
# delivery = ''
# invoice = ''
# with open('temp.txt', 'r', encoding='utf-8') as file:
#     lines = file.readlines()
#     for line in lines:
#         if line.find('Business') != -1:
#             customer_name = " ".join(line.split()[3:])
#         if line.find('Quot') != -1:
#             quotation = line.split()[3][2:]
#         if line.find('SO') != -1:
#             order = line.split()[3][2:]
#         if line.find('Delivery') != -1:
#             delivery = line.split()[3][2:]
#             delivery_date = line.split()[4].replace('.', '/')
#         if line.find('Billing') != -1:
#             invoice = line.split()[3][2:]
#             invoice_date = line.split()[4].replace('.', '/')



while True:
    try:
        all_orders_task = orders_task_list.get_all_records()
        starting_tasks_time = time()
        update_orders_task_list()
        transfer_quotations()
        if time() - starting_tasks_time < 60:
            sleep(60)

    except gspread.exceptions.APIError:
        sleep(120)



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
    



