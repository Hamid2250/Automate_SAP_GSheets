import os
from datetime import datetime, timedelta
from time import sleep, time
import PyPDF2
import gspread
import pyperclip as pc
from pyrobogui import robo, pag
from itertools import zip_longest


class SAPAutomation:
    """Methods for automating SAP by Images"""
    def __init__(self, gs_account_file_path='gs_account.json'):

        # Connect to Google Sheets API
        self.service_account = gspread.service_account(
            filename=gs_account_file_path)

        # Define Google Sheets names and worksheets
        self.sheet = self.service_account.open('Ai workflow')
        self.orders_task_list = self.sheet.worksheet('OrdersTaskList')
        self.orders = self.sheet.worksheet('Orders')

        # Get all column names
        self.orders_task_list_column_names = self.orders_task_list.row_values(1)
        self.orders_column_names = self.orders.row_values(1)

        # Create a dictionary with empty values
        self.orders_task_list_default = dict.fromkeys(
            self.orders_task_list_column_names, '')
        self.orders_default = dict.fromkeys(self.orders_column_names, '')

    def sap_response_time(self, image=None, timeout=5):
        """Wait for SAP to finish processing task"""
        ms_position = pag.locateOnScreen(image='./images/getRT_position.png')
        if timeout is None:
            timeout=5
        else:
            try:
                timeout = int(float(timeout))
            except Exception as exc:
                raise ValueError("Number of seconds must be inserted not text!") from exc
        wait_until = datetime.now() + timedelta(seconds=timeout)
        while True:
            pix = pag.pixel(x=200, y=37)
            if pix == (242, 242, 242):
                if not pag.locateOnScreen(image='./images/lastRT.png') and not pag.locateOnScreen(
                    image='./images/zeroRT.png'):
                    if image is not None:
                        robo.waitImageToAppear(image=image)
                    pag.screenshot(
                        './images/lastRT.png',
                        region=(ms_position.left-150, ms_position.top, 185, 25))
                    break
                if wait_until < datetime.now():
                    if image is not None:
                        print(f"Image not found > {str(image)}")
                    break
            else:
                sleep(0.5)
