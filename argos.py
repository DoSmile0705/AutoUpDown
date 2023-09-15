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
import datetime as dt
from tkinter import *
from tkcalendar import *
from tkinter import messagebox
import keyboard


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
f = ('Times', 20)
start_flag = False


class App(customtkinter.CTk):
    WIDTH = 350
    HEIGHT = 500
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
                                            width=280,
                                            height=300,
                                            corner_radius=10)
        frame_time.place(relx=0.5, rely=0.46, anchor=CENTER)
        global cal, min_sb, sec_hour, sec
        cal = Calendar(
            frame_time,
            selectmode="day",
        )
        cal.place(relx=0.5, rely=0.32, anchor=CENTER)
        min_sb = Spinbox(
            frame_time,
            from_=0,
            to=23,
            wrap=True,
            width=2,
            state="readonly",
            font=f,
            justify=CENTER
        )
        sec_hour = Spinbox(
            frame_time,
            from_=0,
            to=59,
            wrap=True,
            state="readonly",
            font=f,
            width=2,
            justify=CENTER
        )
        sec = Spinbox(
            frame_time,
            from_=0,
            to=59,
            wrap=True,
            state="readonly",
            width=2,
            font=f,
            justify=CENTER
        )
        min_sb.place(relx=0.26, rely=0.65)
        sec_hour.place(relx=0.44, rely=0.65)
        sec.place(relx=0.62, rely=0.65)
        msg = Label(
            frame_time,
            text="H : M : S",
            font=("Times", 12),
            background="#2b2b2b",
            foreground="white"
        )
        msg.place(relx=0.5, rely=0.82, anchor=CENTER)

        def my_time():
            global time_string
            time_string = strftime('%H:%M:%S %p-%a, %m/%d/%Y')  # time format
            l.config(text=time_string, background="#2b2b2b", foreground="white")
            l.after(1000, my_time)  # time delay of 1000 milliseconds

        my_font = ('times', 16, 'bold')  # display size and style

        l = Label(self, font=my_font)
        l.place(x=35, y=50)

        my_time()

        def stop_action():
            global stop_flag
            stop_flag = False
            self.start_auto_button.configure(state="enable")
            self.stop_auto_button.configure(state="disabled")

        def start_action():
            if os.path.isfile(new_name):
                os.remove(new_name)
            if os.path.isfile(old_name):
                os.remove(old_name)
            if os.path.isfile(new_lead):
                os.remove(new_lead)
            if os.path.isfile(old_lead):
                os.remove(old_lead)
            self.stop_auto_button.configure(state="enable")
            global stop_flag
            stop_flag = True
            while stop_flag == True:
                date = cal.get_date()
                m = min_sb.get()
                h = sec_hour.get()
                s = sec.get()
                x = date.split("/")
                month = x[0]
                day = x[1]
                year = x[2]
                year_ = 2000 + int(year)
                month_ = int(month)
                day_ = int(day)
                h_ = int(m)
                m_ = int(h)
                s_ = int(s)
                CurrentDate = datetime.now()
                ExpectedDate = f'{day_}/{month_}/{year_} {h_}:{m_}:{s_}'
                ExpectedDate = datetime.strptime(
                    ExpectedDate, "%d/%m/%Y %H:%M:%S")

                if CurrentDate < ExpectedDate:
                    self.start_auto_button.configure(state="disabled")
                    self.stop_auto_button.configure(state="enable")
                    messagebox.showinfo(
                        "AutoUpDown", f"Automation will start at {m}:{h}:{s}, {month_}/{day_}/{year_}.")
                    # pyautogui.hotkey('win', 'down')
                    while CurrentDate < ExpectedDate:
                        start_flag = False  # Wait for 1 second
                        CurrentDate = datetime.now()
                        self.update()
                    start_flag = True
                else:
                    messagebox.showinfo("AutoUpDown", "Plz set correct time.")
                    start_flag = False
                    break
                

                while start_flag == True:
                    if os.path.isfile(new_name):
                        os.remove(new_name)
                    if os.path.isfile(old_name):
                        os.remove(old_name)
                    if os.path.isfile(new_lead):
                        os.remove(new_lead)
                    if os.path.isfile(old_lead):
                        os.remove(old_lead)

                    time.sleep(3)

                    def section1():
                        os.startfile(
                            "C:\VS Automation Suite\VSAutomationSuite.exe")
                        time.sleep(35)
                        pyautogui.click(x=800, y=500)
                        time.sleep(2)
                        print("ok!")
                        keyDown("Alt")
                        keyUp("Alt")
                        time.sleep(0.5)
                        keyDown("Tab")
                        keyUp("Tab")
                        time.sleep(0.5)
                        keyDown("Tab")
                        keyUp("Tab")
                        time.sleep(0.5)
                        keyDown("Tab")
                        keyUp("Tab")
                        time.sleep(0.5)
                        keyDown("Tab")
                        keyUp("Tab")
                        time.sleep(0.5)
                        keyDown("Enter")
                        keyUp("Enter")
                        time.sleep(15)
                        os.system("taskkill /f /im VSAutomationSuite.exe")
                    section1()

                    def section2():
                        if os.path.isfile(old_name):
                            os.rename(old_name, new_name)
                        else:
                            section1()
                            section2()
                        time.sleep(1)
                        driver = uc.Chrome(options=o)
                        if os.path.isfile(new_name):
                            time.sleep(0.5)
                            driver.maximize_window()
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
                            s = driver.find_element(
                                By.XPATH, "//input[@type='file']")
                            s.send_keys(
                                r"C:\Users\Zack\Documents\FullArgos_" + Num + r".csv")
                            time.sleep(2)
                            pullnumber = driver.find_element(
                                By.XPATH, "/html/body/div/div/div[2]/div[2]/form/input")
                            pullnumber.send_keys(3000)
                            pullnumber.send_keys(Keys.ENTER)
                            os.remove(new_name)
                        else:
                            section1()
                            section2()
                        time.sleep(2)
                        os.rename(old_lead, new_lead)
                        # driver.close()

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
                        

                    section2()
                    self.update()
                    break
                time.sleep(3)
                # self.destroy()
                break

        def on_close():
            self.destroy()

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
            command=start_action,
        )
        self.start_auto_button.place(x=95, y=370)

        self.stop_auto_button = customtkinter.CTkButton(
            master=self.frame,
            text="Stop",
            fg_color=("black"),
            font=("Roboto Medium", -16),
            state="disabled",
            command=stop_action,
        )
        self.stop_auto_button.place(x=95, y=405)

        self.exit_auto_button = customtkinter.CTkButton(
            master=self.frame,
            text="Exit",
            fg_color=("red"),
            font=("Roboto Medium", -16),
            command=on_close,
        )
        self.exit_auto_button.place(x=95, y=440)

        # keyboard.add_hotkey('space', start_action)
        # keyboard.add_hotkey('ctrl+space', stop_action)
        # keyboard.add_hotkey('ctrl+alt+space', on_close)

    def start(self):
        self.mainloop()


if __name__ == "__main__":
    app = App()
    app.attributes("-topmost", True)
    app.resizable(False, False)
    app.update()
    app.start()
