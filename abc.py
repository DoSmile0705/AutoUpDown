import os
import time
import pyautogui
import subprocess
import customtkinter
import tkinter as tk
import undetected_chromedriver as uc
from PIL import Image, ImageTk
from pyautogui import *
from pynput import *
from pynput.keyboard import *
from tkinter.ttk import *
from time import strftime
from selenium import *
from datetime import datetime
from selenium import webdriver
from selenium.common import NoSuchElementException, ElementNotInteractableException
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver import ActionChains
from selenium.webdriver import Keys
from selenium.webdriver.chrome.service import Service as ChromeService
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.select import Select
from zenrows import ZenRowsClient
from selenium.webdriver.chrome.service import Service
from googleapiclient.discovery import build
from google.oauth2 import service_account
from googleapiclient.http import MediaFileUpload
import datetime as dt
from tkinter import *
from tkcalendar import *
from tkinter import messagebox
import keyboard
import shutil

now = datetime.now()
Num = str(now.strftime("%b") + "_" + now.strftime("%d"))
mmm = 0
o = uc.ChromeOptions()

SERVICE_ACCOUNT_FILE = 'argos_cred.json'
FOLDER_ID = '1Ph6YRWYl_KJ9tCMAj19FgFj1aJH499bF'
credentials = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE
)
drive_service = build('drive', 'v3', credentials=credentials)
CurrentDate = datetime.now()
old_name = 'C:/Users/Zach/Documents/VSLeads.csv'
new_name = 'C:/Users/Zach/Documents/FullArgos_' + Num + '.csv'
old_lead = 'C:/Users/Zach/Downloads/leads.csv'
f = ('Times', 45)
start_flag = False


class App(customtkinter.CTk):
    WIDTH = 250
    HEIGHT = 350
    customtkinter.set_appearance_mode("dark")
    customtkinter.set_default_color_theme("blue")

    def __init__(self):
        super().__init__()
        self.title("AutoUpDown")
        self.geometry(f"{self.WIDTH}x{self.HEIGHT}")
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self.frame = customtkinter.CTkFrame(master=self)
        self.frame.grid(row=0, column=1, sticky="nsew", padx=10, pady=10)
        self.frame.grid_columnconfigure(0, weight=1)
        self.frame.grid_columnconfigure(1, weight=1)
        self.frame.grid_columnconfigure(5, weight=1)

        frame_time = customtkinter.CTkFrame(master=self,
                                            width=150,
                                            height=150)
        frame_time.place(relx=0.5, rely=0.46, anchor=CENTER)

        label = Label(
            frame_time,
            text="🎁",
            font=("Times", 100),
            bg="#2b2b2b",
            fg="white"
        )
        label.place(relx=0.5, rely=0.5, anchor=CENTER)

        def my_time():
            global time_string
            time_string = strftime('%H:%M:%S %p-%a, %m/%d/%Y')  # time format
            l.config(text=time_string, background="#2b2b2b", foreground="white")
            l.after(1000, my_time)  # time delay of 1000 milliseconds

        my_font = ('times', 12, 'bold')  # display size and style

        l = Label(self, font=my_font)
        l.place(x=25, y=50)

        my_time()

        def start_action():
            CurrentDate = datetime.now()
            year_ = CurrentDate.year
            month_ = CurrentDate.month
            day_ = CurrentDate.day
            h_ = 1
            h = 23
            m_ = 59
            s_ = 10
            ExpectedDate1 = f'{day_}/{month_}/{year_} {h_}:{m_}:{s_}'
            ExpectedDate1 = datetime.strptime(
                ExpectedDate1, "%d/%m/%Y %H:%M:%S")
            ExpectedDate2 = f'{day_}/{month_}/{year_} {h}:{m_}:{s_}'
            ExpectedDate2 = datetime.strptime(
                ExpectedDate2, "%d/%m/%Y %H:%M:%S")
            if CurrentDate > ExpectedDate2:
                messagebox.showinfo(
                    "AutoUpDown", f"Plz upload next day.")
                self.destroy()
            else:
                messagebox.showinfo(
                    "AutoUpDown", f"Automation will run at {h_}:{m_}:{s_}, {month_}/{day_}/{year_}.")
                pyautogui.hotkey('win', 'down')
                self.start_auto_button.configure(state="disabled")
                start_flag = False
                while ExpectedDate1 > CurrentDate or CurrentDate >= ExpectedDate2:
                    CurrentDate = datetime.now()
                    start_flag = False
                    self.update()
                start_flag = True
                while start_flag == True:
                    driver = uc.Chrome(options=o)
                    time.sleep(0.5)
                    driver.maximize_window()
                    time.sleep(2)
                    driver.get(
                        "http://ushtoolkit.com/login?next=%2Fhome")
                    time.sleep(1)
                    mailaddress = driver.find_element(By.ID, "email")
                    loginpwd = driver.find_element(By.ID, "password")
                    time.sleep(4)
                    mailaddress.send_keys("admin@admin")
                    loginpwd.send_keys("123@dmin123")
                    loginpwd.send_keys(Keys.ENTER)
                    time.sleep(2)
                    driver.get("http://ushtoolkit.com/home")
                    
                    def action():
                        if os.path.isfile(old_name):
                            global mmm
                            mmm = mmm + 1
                            print(mmm)
                            time.sleep(10)
                            if mmm == 9:
                                self.destroy()
                            nnn = str(mmm)
                            new_lead = 'C:/Users/Zach/Downloads/ArgosLeads(' + nnn + ').csv'
                            if os.path.isfile(new_lead):
                                os.remove(new_lead)
                            if os.path.isfile(old_lead):
                                os.remove(old_lead)
                            if os.path.isfile(new_name):
                                os.remove(new_name)
                            dest = shutil.copyfile(old_name, new_name)
                            print(dest)
                            # os.rename(old_name, new_name)
                            s = driver.find_element(
                                By.XPATH, "//input[@type='file']")
                            s.send_keys(
                                r"C:\Users\Zach\Documents\FullArgos_" + Num + r".csv")
                            time.sleep(2)
                            pullnumber = driver.find_element(
                                By.XPATH, "/html/body/div/div/div[2]/div[2]/form/input")
                            pullnumber.send_keys(3000)
                            pullnumber.send_keys(Keys.ENTER)
                            pullnumber.send_keys(Keys.CONTROL + 'a', Keys.BACKSPACE)
                            os.remove(new_name)
                            time.sleep(25)
                            os.rename(old_lead, new_lead)
                            # results = drive_service.files().list(
                            #     q=f"name='{Num}' and mimeType='application/vnd.google-apps.folder' and trashed=false",
                            #     fields="files(id)").execute()
                            # items = results.get('files', [])
                            # Delete the folder(s) with the specified name
                            # for item in items:
                            #     folder_id = item['id']
                            #     drive_service.files().delete(fileId=folder_id).execute()
                            #     print(f"Deleted folder: {folder_id}")
                            # drive_service.files().delete(filename = Num).exvute
                            folder_metadata = {
                                'name': Num,
                                'parents': [FOLDER_ID],
                                'mimeType': 'application/vnd.google-apps.folder'
                            }
                            # Create the folder in Google Drive
                            created_folder = drive_service.files().create(
                                body=folder_metadata,
                                fields='id'
                            ).execute()
                            CREATED_FOLDER_ID = created_folder.get("id")
                            # Change this to the file you want to upload
                            # a =0
                            # a = a+1

                            # Num = Num + 1
                            file_name = 'ArgosLeads(' + nnn + ').csv'
                            # # Change this to the actual file path
                            file_path = 'C:/Users/Zach/Downloads/' + file_name
                            # Create a file metadata with the desired folder ID
                            file_metadata = {
                                'name': file_name,
                                'parents': [CREATED_FOLDER_ID]
                            }
                            # Upload the file to Google Drive
                            media_body = MediaFileUpload(
                                file_path, resumable=True)
                            file = drive_service.files().create(
                                body=file_metadata,
                                media_body=media_body,
                                fields='id'
                            ).execute()
                            time.sleep(180)
                            action()
                        else:
                            time.sleep(10)
                            action()
                    action()
                return

        def on_close():
            self.destroy()

        # title section
        self.title = customtkinter.CTkLabel(
            master=self.frame, text="Current Time", font=("Times", -22),
        )
        self.title.place(x=60, y=10)
        # button section
        self.start_auto_button = customtkinter.CTkButton(
            master=self.frame,
            text="Start",
            fg_color=("black"),
            font=("Roboto Medium", -16),
            command=start_action,
        )
        self.start_auto_button.place(x=47, y=250)

        self.exit_auto_button = customtkinter.CTkButton(
            master=self.frame,
            text="Exit",
            fg_color=("red"),
            font=("Roboto Medium", -16),
            command=on_close,
        )
        self.exit_auto_button.place(x=47, y=285)

    def start(self):
        self.mainloop()


if __name__ == "__main__":
    app = App()
    app.attributes("-topmost", True)
    app.resizable(False, False)
    app.update()
    app.start()
