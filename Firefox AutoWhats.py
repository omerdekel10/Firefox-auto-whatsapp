"""Automate file and/or text messaging from Firefox Whatsapp Web. File formats are JPG, JPEG, PNG, BMP, or PNG,
contacts or numbers are to be provided in an .xlsx file, """

from copy import Error
from datetime import datetime as d
from logging import raiseExceptions
import os
from pandas import read_excel as xl
import pyautogui as pg
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from time import sleep
from urllib.parse import quote
import webbrowser
from win32clipboard import UNICODE

class Firefox_Whatsapp:

    def __init__(self, message="", image_path="", contacts_path="", numbers_path="", prefix="972", log=False) -> None:
        self.message = message
        self.image_path = image_path
        self.contacts_path = contacts_path
        self.numbers_path = numbers_path
        self.prefix = prefix.strip("+")
        self.is_log = log
        self.is_message = 0
        self.is_image = 0
        self.names = []

    def check_vars(self) -> None:
        if not self.message == "":
            self.is_message = 1
        if os.path.exists(self.image_path) and len(self.image_path) > 0:
            self.is_image = 1
        elif len(self.image_path) == 0:
            self.is_image = 0
        else:
            raise Error("Image path either invalid or does not exist. Please check.")
        if self.is_image == 0 and self.is_message == 0:
            raise Error(
                "I am a message robot, my purpose is to send messages. No message or file path provided :(")
        if self.contacts_path == "" and self.numbers_path == "":
            raise FileNotFoundError(
                "You must enter a valid excel path with contacts or celular numbers.")
        if not os.path.exists(self.contacts_path) and len(self.contacts_path) > 0:
            raise FileNotFoundError(
                r"Please provide a valid excel path for contacts. Example: C:\\Users\\Desktop\\contacts.xlsx")
        elif os.path.exists(self.contacts_path):
            self.contacts = xl(io=self.contacts_path, usecols=[0], header=None)
            self.names = [name for name in self.contacts[0] if type(name) == str]
            if len(self.names) == 0:
                raise Error(
                    "No contacts found, please check file. All contacts should be on column A, no headres.")
        if not os.path.exists(self.numbers_path) and len(self.numbers_path) > 0:
            raise FileNotFoundError(
                r"Please provide a valid excel path for numbers. Example: C:\Users\Desktop\numbers.xlsx")
        elif os.path.exists(self.numbers_path):
            self.numbers = xl(io=self.numbers_path, usecols=[0], header=None)
            self.numbers = [str(number) for number in self.numbers[0] if str(number).isdigit()]
            for i, number in enumerate(self.numbers):
                if number.startswith('0'):
                    number = '+'+self.prefix + number[1:12]
                    self.numbers[i] = number
                if number.startswith('5'):
                    number = '+'+self.prefix + number
                    self.numbers[i] = number
            if len(self.numbers) == 0:
                raise Error(
                    "No numbers fonund, please check file. All numbers should be on column A, no headres.")
        if not self.image_path.upper().endswith(("JPG","JPEG","PNG", "BMP", "PNG")):
            raise TypeError(
                f"{self.image_path} is not a valid image file type, (valid file types are .jpg .jpeg .png .bmp)" )
        return None

    def log_file(self, text) -> None:
        time = d.today()
        now = time.strftime('%d-%m-%y %H-%M')
        log_path = os.path.join(os.getcwd(), f"Messages log - {now}.txt")
        with open (log_path, "a") as log_file:
            if self.is_image == 1:
                log_file.write(f"File: {self.image_path}\n")
                log_file.write(f"Message: {self.message}\n")
                log_file.write(text)
            else:
                log_file.write(f"Message: {self.message}\n ")
                log_file.write(text)
        return None

    def contacts_message(self) -> None:

        self.check_vars()
        log_str = ""
        driver = webdriver.Firefox()
        driver.get("https://web.whatsapp.com/")
        driver.maximize_window()
        sleep(12)        

        for name in self.names:
            try:
                input_path = '//div[@class="_13NKt copyable-text selectable-text"][@data-tab="3"][@dir="rtl"]'
                search_box = WebDriverWait(driver,4).until(EC.presence_of_element_located((By.XPATH,input_path)))
                search_box.send_keys(name)
                sleep(1)
                search_box.send_keys(Keys.ENTER)
                sleep(1)
            except:
                if self.is_log == True:
                    time = d.today()
                    now = time.strftime('%H:%M:%S')
                    log_str += f"Failed to send to {name} at {now}. Contact not found.\n"
                continue

            input_path = '//div[@class="_13NKt copyable-text selectable-text"][@dir="rtl"][@data-tab="6"]'
            msg_box = WebDriverWait(driver,4).until(EC.presence_of_element_located((By.XPATH,input_path)))
            msg_box.send_keys(self.message)
            sleep(1)

            if self.is_image == 1:
                attech_button = driver.find_element_by_xpath("//div[@title = 'הוספת קובץ']")
                attech_button.click()
                sleep(2)

                media_button = driver.find_element_by_xpath("//input[@accept = 'image/*,video/mp4,video/3gpp,video/quicktime']")
                media_button.send_keys(self.image_path)
                sleep(2)

            send = driver.find_element_by_xpath("//span[@data-icon = 'send']")
            send.click()
            sleep(2)

            if self.is_log == True:
                time = d.today()
                now = time.strftime('%H:%M:%S')
                log_str += f"Message sent to {name} at {now}\n"

        driver.quit()

        if self.is_log == True:
            self.log_file(log_str)
        return None

    def numbers_message(self):
        self.check_vars()
        width, height = pg.size()
        pg.PAUSE = 1.2

        for number in self.numbers:

            log_str=""
            parsed_msg = quote(self.message)
            sleep(2)
            webbrowser.open(f'https://web.whatsapp.com/send?phone={number}&text={parsed_msg}')
            sleep(16)

            if self.is_image == 1:
                import win32clipboard
                from io import BytesIO
                from PIL import Image
                image = Image.open(self.image_path)
                output = BytesIO()
                image.convert('RGB').save(output, "BMP")
                data = output.getvalue()[14:]
                output.close()

                pg.leftClick(width*0.43125,height*0.89259) # smaller decimal - higher precision
                win32clipboard.OpenClipboard()
                win32clipboard.EmptyClipboard()
                win32clipboard.SetClipboardData(win32clipboard.CF_DIBV5,data )
                win32clipboard.CloseClipboard()
                sleep(2)
                pg.hotkey('ctrl', 'v')

                pg.press("enter")

            sleep(1)
            pg.hotkey('ctrl', 'w')
            pg.press('enter')

            if self.is_log == True:
                time = d.today()
                now = time.strftime('%H:%M:%S')
                log_str += f"Message sent to {number} at {now}\n"

        if self.is_log == True:
            self.log_file(log_str)
        return None

# image = r'C:\Users\user\Desktop\pic.jpg'
# con_path = "C:\\Users\\user\\Desktop\\contacts.xlsx" 
# num_path = "C:\\Users\\user\\Desktop\\numbers.xlsx"
# text="Hi! This is an automated message"

# fox_num = Firefox_Whatsapp(message=text, numbers_path=num_path, image_path=image, prefix = '972', log=True)
# fox_num.numbers_message()
# fox_con = Firefox_Whatsapp(message=text, contacts_path=con_path, image_path=image, prefix = '972', log=True)
# fox_con.contacts_message()