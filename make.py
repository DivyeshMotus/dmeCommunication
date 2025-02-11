from google.oauth2 import service_account
from googleapiclient.discovery import build
import pandas as pd
import os
from dotenv import load_dotenv
from PyPDFForm import FormWrapper, PdfWrapper
import datetime
import shutil
import time
from twilio.rest import Client
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

load_dotenv()

SHEET_ID = os.getenv('SHEET_ID')
RANGE_NAME = os.getenv('RANGE_NAME')
USER_EMAIL = os.getenv("USER_EMAIL")
USER_PASSWORD = os.getenv("USER_PASSWORD")
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

def authenticate_services():
    credentials = service_account.Credentials.from_service_account_file('credentials.json', scopes=SCOPES)
    sheets_service = build('sheets', 'v4', credentials=credentials)
    return sheets_service

def read_all_data_from_sheet(sheets_service, sheet_id, range_name):
    sheet = sheets_service.spreadsheets()
    result = sheet.values().get(spreadsheetId=sheet_id, range=range_name).execute()
    values = result.get('values', [])
    return values

def add_to_dataframe(data):
    column_names = data[0]
    data = data[1:]
    return pd.DataFrame(data, columns=column_names)

def format_date(date_string):
    try:
        formatted_date = datetime.datetime.strptime(date_string, "%m/%d/%Y").date()    
        return formatted_date
    except Exception as e:
        print(f"Error: {e}")
        return None

def choose_dme(dme_number):
    dme_number = int(dme_number)
    if dme_number == 1:
        dme_name = 'ROM (GA)'
    elif dme_number == 2:
        dme_name = 'MedEx (TX)'
    elif dme_number == 3:
        dme_name = 'Bestcare (NY)'
    elif dme_number == 4:
        dme_name = 'PrimeMedical (NY)'
    elif dme_number == 5:
        dme_name = 'SSMedicalSupply (FL)'
    elif dme_number == 6:
        dme_name = 'SOSortho (VA)'
    elif dme_number == 7:
        dme_name = 'Medequip'
    elif dme_number == 8:
        dme_name = 'ATI (OH/KY)'
    elif dme_number == 9:
        dme_name = 'RMS (GA)'
    elif dme_number == 10:
        dme_name = 'FlexOrtho'
    elif dme_number == 11:
        dme_name = 'Innova'
    elif dme_number == 12:
        dme_name = 'Eclipse'
    else:
        dme_name = 'Invalid DME number'
    return dme_name

def check_date(record):
    formatted_date = format_date(record[2])
    today_date = datetime.datetime.today().date()
    diff = today_date - formatted_date
    days_dif = -1
    if diff.days < 0:
        days_dif = 365 + diff.days
    else:
        days_dif = diff.days
    if days_dif > 7:
        return True
    else:
        return False

def get_dme_claims(dme_name, df):
    new_claims = []
    for record in df.itertuples():
        if record[5] == dme_name:
            send_email = check_date(record)
            if send_email:
                new_claims.append(f'{record[1]}: {record[3]} {record[4]}')
    return new_claims

def generate_email_message(new_claims):
    claims_list = "\n".join(new_claims)
    message = f"""
Hi,

Could you please have a look at the new claims in the sheet? Are you able to bill any of these?

Here is a list of the new claims:
{claims_list}

Best,
Divyesh Ved."""
    return message

def send_email(to_address, cc_address, subject, body, smtp_server, smtp_port, smtp_user, smtp_password):
    msg = MIMEMultipart()
    msg['From'] = smtp_user
    msg['To'] = ', '.join(to_address)  # Join addresses with commas
    msg['Cc'] = ', '.join(cc_address)  # Join CC addresses with commas
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))
    recipients = to_address + cc_address
    server = None
    try:
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(smtp_user, smtp_password)
        server.send_message(msg, to_addrs=recipients)
        print(f"Email sent successfully to {', '.join(recipients)}")
    except Exception as e:
        print(f"Error sending email: {e}")
    finally:
        if server:
            server.quit()

def main():
    dme_number = input("Enter the DME supplier's name that you want send the email to: ")
    dme_name = choose_dme(dme_number) 
    sheets_service = authenticate_services()
    data = read_all_data_from_sheet(sheets_service, SHEET_ID, RANGE_NAME)
    df = add_to_dataframe(data)
    new_claims = get_dme_claims(dme_name, df)
    email_body = generate_email_message(new_claims)
    to_address = ["duveds1@gmail.com"]
    cc_address =["duveds@gmail.com"]
    subject = f"New claims to bill for Motus Nova!"
    smtp_server = "smtp.gmail.com"
    smtp_port = 587
    smtp_user = USER_EMAIL
    smtp_password = USER_PASSWORD
    send_email(to_address, cc_address, subject, email_body, smtp_server, smtp_port, smtp_user, smtp_password)

if __name__ == '__main__':
    main()