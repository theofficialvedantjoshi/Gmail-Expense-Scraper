from __future__ import print_function
import customtkinter
import os.path
import re
import pandas as pd
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from PIL import Image
import tkinter as tk
from tkinter import ttk
import threading


SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']

customtkinter.set_appearance_mode("dark")
customtkinter.set_default_color_theme("green")

root = customtkinter.CTk()
root.title("Expense Tracker")
root.geometry("400x650")
entry1 = None
entry2 = None
entry3 = None
pg = None
def action():
    def process_emails():
        global entry1
        global entry2
        global entry3
        global pg
        global root
        df = pd.DataFrame(columns=['Amount','Date'])
        creds = None
        # The file token.json stores the user's access and refresh tokens, and is
        # created automatically when the authorization flow completes for the first
        # time.
        if os.path.exists('token.json'):
            creds = Credentials.from_authorized_user_file('token.json', SCOPES)
        # If there are no (valid) credentials available, let the user log in.
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    'credentials.json', SCOPES)
                creds = flow.run_local_server(port=0)
            # Save the credentials for the next run
            with open('token.json', 'w') as token:
                token.write(creds.to_json())

        try:
            # Call the Gmail API
            amounts = []
            dates = []
            bnk = []
            bank = entry1.get()
            service = build('gmail', 'v1', credentials=creds)
            l_list = service.users().messages().list(userId='me',q=("from:{} is:unread".format(bank))).execute()
            ids=[]  
            for i in l_list.get('messages'):
                ids.append(i.get('id'))
            n = int(entry2.get())
            stepval = 100/n
            pg['maximum'] = 100
            #getting message ids and email bodies and adding that to our lists.
            for i in range(n):
                request = service.users().messages().get(userId='me',id=ids[i],)
                request.execute()
                transaction = str(request.execute().get('snippet'))
                #print("Computing transaction ",i+1," of ",n,'...\n')
                try:
                    amounts.append(float(re.findall(r'Rs.\d+',transaction)[0].strip('Rs.')))
                    dates.append(re.findall(r'\d{2}[-]\d{2}[-]\d{2}',transaction)[0])
                except:
                    continue
                pg['value']+=stepval
            path = entry3.get() + "\RECORDS.xlsx"
            #creating dataframe.
            df['Amount'] = amounts
            df['Date'] = dates
            df.set_index('Date', inplace=True)
            inf = df.describe()
            daywise = df.groupby('Date').sum()
            #writing to excel file.
            with pd.ExcelWriter(path) as writer:
                df.to_excel(writer, sheet_name="BANK RECORDS", index=True)
                daywise.to_excel(writer, sheet_name="DAYWISE", index=True)
                inf.to_excel(writer, sheet_name="STATISTICS", index=True)
            #print("Records computed.")
            
        except HttpError as error:
            # TODO(developer) - Handle errors from gmail API.
            print(f'An error occurred: {error}')
    email_processing_thread = threading.Thread(target=process_emails)
    email_processing_thread.start()


def main():
    global frame1
    global entry1
    global entry2
    global entry3
    global pg
    def combobox_callback(choice):
        if choice == "Last 20":
            entry2.delete(0, "end")
            entry2.insert(0, "20")
        elif choice == "Last 50":
            entry2.delete(0, "end")
            entry2.insert(0, "50")
        elif choice == "Last 100":
            entry2.delete(0, "end")
            entry2.insert(0, "100")
    frame1.destroy()
    frame = customtkinter.CTkFrame(root)
    frame.pack(pady=20, padx=20, fill="both", expand=False)

    label = customtkinter.CTkLabel(frame, text="Expense Tracker", font=("Roboto", 20))
    label.pack(pady=12, padx=10)

    label = customtkinter.CTkLabel(frame, text="Enter your bank alerts email address", font=("Roboto", 12))
    label.pack(pady=12, padx=10)

    entry1 = customtkinter.CTkEntry(frame,placeholder_text="Email", font=("Roboto", 12))
    entry1.pack(pady=12, padx=10)

    label = customtkinter.CTkLabel(frame, text="Enter n.o of bank alerts you want to see ", font=("Roboto", 12))
    label.pack(pady=12, padx=10)

    

    entry2 = customtkinter.CTkEntry(frame, font=("Roboto", 12))
    entry2.pack(pady=12, padx=10)

    label = customtkinter.CTkLabel(frame, text="Or, ", font=("Roboto", 12))
    label.pack(pady=5, padx=5)
    combobox = customtkinter.CTkComboBox(master=frame,
                                     values=["Null","Last 20", "Last 50", "Last 100"],
                                     command=combobox_callback)
    combobox.pack(padx=20, pady=10)
    combobox.set("Null")
    label = customtkinter.CTkLabel(frame, text="Enter path for Excel Sheet", font=("Roboto", 12))
    label.pack(pady=12, padx=10)

    entry3 = customtkinter.CTkEntry(frame,placeholder_text="PATH", font=("Roboto", 12))
    entry3.pack(pady=12, padx=10)

    button = customtkinter.CTkButton(frame, text="Submit", font=("Roboto", 12), command=action)
    button.pack(pady=12, padx=10)

    style = ttk.Style()
    style.configure("TProgressbar", background="black")
    
    pg = ttk.Progressbar(frame,orient="horizontal", mode="determinate", length=100)
    pg.pack(pady=12, padx=10)
    
    label = customtkinter.CTkLabel(frame, text="YOU WILL BE REDIRECTED TO SIGN-IN", font=("Roboto", 12))
    label.pack(pady=12, padx=10)

frame1 = customtkinter.CTkFrame(root)
frame1.pack(pady=2, padx=2, fill="both", expand=True)
my_image = customtkinter.CTkImage(dark_image=Image.open("image.png"),
                                    size=(400, 600))
image_label = customtkinter.CTkLabel(frame1, image=my_image, text="")
image_label.pack()

button1 = customtkinter.CTkButton(frame1, text="PROCEED", font=("Roboto", 20), command=main)
button1.pack(pady=12, padx=10)

root.mainloop()
