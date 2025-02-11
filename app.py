from flask import Flask, render_template, request, redirect, url_for, flash
from google.oauth2 import service_account
from googleapiclient.discovery import build
import pandas as pd
import datetime
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import os
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

app = Flask(__name__)
app.secret_key = os.urandom(24) # Required for flashing messages

SHEET_ID = os.getenv('SHEET_ID')
RANGE_NAME = os.getenv('RANGE_NAME')
USER_EMAIL = os.getenv("USER_EMAIL")
USER_PASSWORD = os.getenv("USER_PASSWORD")
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# Helper functions
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
        return datetime.datetime.strptime(date_string, "%m/%d/%Y").date()
    except Exception as e:
        print(f"Error: {e}")
        return None

def choose_dme(dme_number):
    dme_names = {
        1: 'ROM (GA)',
        2: 'MedEx (TX)',
        3: 'Bestcare (NY)',
        4: 'PrimeMedical (NY)',
        5: 'SSMedicalSupply (FL)',
        6: 'SOSortho (VA)',
        7: 'Medequip',
        8: 'ATI (OH/KY)',
        9: 'RMS (GA)',
        10: 'FlexOrtho',
        11: 'Innova',
        12: 'Eclipse'
    }
    return dme_names.get(int(dme_number), 'Invalid DME number')

def check_date(record_date):
    today_date = datetime.datetime.today().date()
    diff = (today_date - record_date).days
    return diff > 7

def get_dme_claims(dme_name, df):
    new_claims = []
    for _, row in df.iterrows():
        record_date = format_date(row['Date'])  # Replace with the actual column name
        if record_date and row['DME'] == dme_name and check_date(record_date):  # Replace 'DME Name' with your column
            new_claims.append(f"{row['PID']}: {row['FirstName']} {row['LastName']}")  # Replace with actual column names
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
    msg['To'] = ', '.join(to_address)
    msg['Cc'] = ', '.join(cc_address)
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))
    recipients = to_address + cc_address
    server = smtplib.SMTP(smtp_server, smtp_port)
    server.starttls()
    server.login(smtp_user, smtp_password)
    server.send_message(msg, to_addrs=recipients)
    server.quit()

# Flask routes
@app.route('/')
def index():
    dme_names = {
        1: 'ROM (GA)',
        2: 'MedEx (TX)',
        3: 'Bestcare (NY)',
        4: 'PrimeMedical (NY)',
        5: 'SSMedicalSupply (FL)',
        6: 'SOSortho (VA)',
        7: 'Medequip',
        8: 'ATI (OH/KY)',
        9: 'RMS (GA)',
        10: 'FlexOrtho',
        11: 'Innova',
        12: 'Eclipse'
    }
    return render_template('index.html', dme_names=dme_names)

@app.route('/send_email', methods=['POST'])
def email_for_new_claims():
    dme_number = request.form['dme_number']
    dme_name = choose_dme(dme_number)
    sheets_service = authenticate_services()
    data = read_all_data_from_sheet(sheets_service, SHEET_ID, RANGE_NAME)
    df = add_to_dataframe(data)
    new_claims = get_dme_claims(dme_name, df)
    if not new_claims:
        flash("No new claims found for the selected DME.")
        return redirect(url_for('index'))
    email_body = generate_email_message(new_claims)
    to_address = ["duveds15@gmail.com"]
    cc_address = ["duveds@gmail.com"]
    subject = f"New claims to bill for Motus Nova!"
    smtp_server = "smtp.gmail.com"
    smtp_port = 587
    smtp_user = USER_EMAIL
    smtp_password = USER_PASSWORD
    send_email(to_address, cc_address, subject, email_body, smtp_server, smtp_port, smtp_user, smtp_password)
    flash("Email sent successfully!")
    return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(debug=True, port=8002)