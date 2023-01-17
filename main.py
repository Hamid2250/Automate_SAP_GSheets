import setWorkspace
import gspread
from pyrobogui import robo, pag
from itertools import zip_longest


service_account = gspread.service_account(filename='gs_account.json')
sheet = service_account.open('Ai workflow')
orders_task_list = sheet.worksheet('OrdersTaskList')
orders = sheet.worksheet('Orders')



# quotations = orders_task_list.get_values('E2:F')
# status = orders_task_list.get_values('I2:I')
# new_list = [[x[0], x[1], y[0] if y else ''] for x, y in zip_longest(quotations, status, fillvalue=[''])]
# print(new_list)
try:
    quotations = orders_task_list.get_values('E2:E')
    for quotation in quotations:
        robo.click(image='./images/command_box.png')
        print(quotation[0])

except gspread.exceptions.APIError:
    print('No quotation')
    