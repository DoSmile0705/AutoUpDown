import os
import time
import pyautogui
import subprocess
import customtkinter
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
import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.select import Select
from zenrows import ZenRowsClient
from selenium.webdriver.chrome.service import Service
from googleapiclient.discovery import build
from google.oauth2 import service_account
from googleapiclient.http import MediaFileUpload

now = datetime.now()
Num = str(now.strftime("%b") + "_" + now.strftime("%d"))
o = uc.ChromeOptions()
SERVICE_ACCOUNT_FILE = 'argos_cred.json'
FOLDER_ID = '1Ph6YRWYl_KJ9tCMAj19FgFj1aJH499bF'
credentials = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE
)
drive_service = build('drive', 'v3', credentials=credentials)
old_name = 'C:/Users/Zach/Documents/VSLeads.csv'
new_name = 'C:/Users/Zach/Documents/FullArgos_' + Num + '.csv'
old_lead = 'C:/Users/Zach/Downloads/leads.csv'
new_lead = 'C:/Users/Zach/Downloads/ArgosLeads_' + Num + '.csv'
uc.install(executable_path=PATH,)
drivers_dict = {}

class App(customtkinter.CTk):
    WIDTH = 350
    HEIGHT = 440

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

        def my_time():
            global time_string
            time_string = strftime('%H:%M:%S %p-%a, %m/%d/%Y')  # time format
            l.config(text=time_string, background="#2b2b2b", foreground="white")
            l.after(1000, my_time)  # time delay of 1000 milliseconds

        my_font = ('times', 16, 'bold')  # display size and style

        l = Label(self, font=my_font)
        l.place(x=35, y=50)

        my_time()

        # title section
        self.title = customtkinter.CTkLabel(
            master=self.frame, text="Current Time", font=("Times", -25),
        )
        self.title.grid(row=1, column=1, pady=10)

        # button section
        self.start_auto_button = customtkinter.CTkButton(
            master=self.frame,
            text="Start",
            fg_color=("black"),
            font=("Roboto Medium", -16),
            command=self.start_action,
        )
        self.start_auto_button.place(x=95, y=280)

        self.stop_auto_button = customtkinter.CTkButton(
            master=self.frame,
            text="Stop",
            fg_color=("black"),
            font=("Roboto Medium", -16),
            state="disabled",
            command=self.stop_action,
        )
        self.stop_auto_button.place(x=95, y=315)

        self.exit_auto_button = customtkinter.CTkButton(
            master=self.frame,
            text="Exit",
            fg_color=("red"),
            font=("Roboto Medium", -16),
            command=self.on_close,
        )
        self.exit_auto_button.place(x=95, y=370)

    def start_action(self):
        # global new_name, old_name
        self.start_auto_button.configure(state="disabled")
        self.stop_auto_button.configure(state="enable")
        if os.path.isfile(new_name):
            os.remove(new_name)
        if os.path.isfile(old_name):
            os.remove(old_name)
        if os.path.isfile(new_lead):
            os.remove(new_lead)
        if os.path.isfile(old_lead):
            os.remove(old_lead)
        print("This is test!")

        def section1():
            os.startfile("C:\VS Automation Suite\VSAutomationSuite.exe")
            time.sleep(20)
            keyDown("Alt")
            keyUp("Alt")
            keyDown("Tab")
            keyUp("Tab")
            keyDown("Tab")
            keyUp("Tab")
            keyDown("Tab")
            keyUp("Tab")
            keyDown("Tab")
            keyUp("Tab")
            keyDown("Enter")
            keyUp("Enter")
            time.sleep(15)
            keyDown("Tab")
            keyUp("Tab")
            keyDown("Tab")
            keyUp("Tab")
            keyDown("Enter")
            keyUp("Enter")
            os.rename(old_name, new_name)
            time.sleep(1)
            os.system("taskkill /f /im VSAutomationSuite.exe")
        section1()
        driver = uc.Chrome(options=o)
        def section2():
            if os.path.isfile(new_name):
                time.sleep(0.5)
                driver.maximize_window()
                driver.get("http://ushtoolkit.com/login?next=%2Fhome")
                time.sleep(0.5)
                mailaddress = driver.find_element(By.ID, "email")
                loginpwd = driver.find_element(By.ID, "password")
                time.sleep(4)
                mailaddress.send_keys("admin@admin")
                loginpwd.send_keys("123@dmin123")
                loginpwd.send_keys(Keys.ENTER)
                time.sleep(1)
                driver.get("http://ushtoolkit.com/home")
                time.sleep(0.5)
                s = driver.find_element(By.XPATH, "//input[@type='file']")
                s.send_keys(
                    r"C:\Users\Zack\Documents\FullArgos_" + Num + r".csv")
                time.sleep(1)
                pullnumber = driver.find_element(
                    By.XPATH, "/html/body/div/div/div[2]/div[2]/form/input")
                pullnumber.send_keys(3000)
                pullnumber.send_keys(Keys.ENTER)
            else:
                section1()
                section2()
            time.sleep(1)
            os.rename(old_lead, new_lead)
            driver.close()
        section2()

        results = drive_service.files().list(
            q=f"name='{Num}' and mimeType='application/vnd.google-apps.folder' and trashed=false",
            fields="files(id)").execute()
        items = results.get('files', [])

        # Delete the folder(s) with the specified name
        for item in items:
            folder_id = item['id']
            drive_service.files().delete(fileId=folder_id).execute()
            print(f"Deleted folder: {folder_id}")
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
        file_name = 'ArgosLeads_' + Num + '.csv'
        # # Change this to the actual file path
        file_path = 'C:/Users/Zach/Downloads/' + file_name

        # Create a file metadata with the desired folder ID
        file_metadata = {
            'name': file_name,
            'parents': [CREATED_FOLDER_ID]
        }

        # Upload the file to Google Drive
        media_body = MediaFileUpload(file_path, resumable=True)
        file = drive_service.files().create(
            body=file_metadata,
            media_body=media_body,
            fields='id'
        ).execute()

        self.start_auto_button.configure(state="enable")
        self.stop_auto_button.configure(state="disable")

    def stop_action(self):
        self.start_auto_button.configure(state="enable")
        self.stop_auto_button.configure(state="disabled")

    def on_close(self):
        self.destroy()

    def start(self):
        self.mainloop()


if __name__ == "__main__":
    app = App()
    app.attributes("-topmost", True)
    app.resizable(False, False)
    app.update()
    app.start()
