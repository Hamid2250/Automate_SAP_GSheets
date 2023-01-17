import setWorkspace
import gspread
import pyperclip as pc
from time import sleep
from pyrobogui import robo, pag
from itertools import zip_longest


service_account = gspread.service_account(filename='gs_account.json')
sheet = service_account.open('Ai workflow')
orders_task_list = sheet.worksheet('OrdersTaskList')
orders = sheet.worksheet('Orders')

def sap_response_time():
    robo.waitImageToAppear(image='./images/getRT_position.png')
    ms_position = pag.locateOnScreen(image='./images/getRT_position.png')
    while True:
        if not pag.locateOnScreen(image='./images/lastRT.png'):
            if not pag.locateOnScreen(image='./images/zeroRT.png'):
                pag.screenshot('./images/lastRT.png', region=(ms_position.left-150, ms_position.top, 185, 25))
                break

# quotations = orders_task_list.get_values('E2:F')
# status = orders_task_list.get_values('I2:I')
# new_list = [[x[0], x[1], y[0] if y else ''] for x, y in zip_longest(quotations, status, fillvalue=[''])]
# print(new_list)
try:
    quotations = orders_task_list.get_values('E2:E')
    for quotation in quotations:
        storage_location = None
        sales_office = None
        robo.click(image='./images/command_box.png')
        pag.typewrite('/N VA22\n')
        sap_response_time()
        pag.typewrite(quotation[0]+'\n')
        sap_response_time()
        sleep(1)
        if pag.locateOnScreen(image='./images/makkah.png', confidence=0.7) is not None:
            sales_office = 'S104'
            print(sales_office)
        if pag.locateOnScreen(image='./images/mmwm.png', confidence=0.7) is not None:
            storage_location = 'BN1'
            print(storage_location)
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
        robo.click(image='./images/sales_document.png')
        robo.click(image='./images/deliver.png')
        sap_response_time()
        robo.click(image='./images/hat.png')
        sap_response_time()
        if pag.locateOnScreen(image='./images/administration.png'):
            robo.click(image='./images/administration.png')
            sap_response_time()
        pc.copy('تحضير وارسال فوري')
        pag.hotkey('ctrl', 'v')
        robo.click(image='./images/save.png')
        sap_response_time()
        print(quotation[0])

except gspread.exceptions.APIError:
    print('No quotation')
    