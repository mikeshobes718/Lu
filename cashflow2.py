import warnings
import imaplib
import email
from email.header import decode_header
import datetime
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build

# Suppress DeprecationWarnings
warnings.filterwarnings("ignore", category=DeprecationWarning)

# Email credentials
GMAIL_USER = 'lucrecia.pedrozo@gmail.com'
GMAIL_PASS = 'gwcr mgjw tzqu lfgt'  # Use your password

# Google Sheets setup
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SERVICE_ACCOUNT_FILE = 'mythic-rain-404717-176959db5d8e.json'  # Path to your service account file

# The ID and range of your spreadsheet
SPREADSHEET_ID = '1ZI9dbE98p88K_VrPuDcE17447XUU2TLrzgJlVRFgBXU'  # Replace with your actual spreadsheet ID
SHEET_NAME = 'Hoja 1'

creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
service = build('sheets', 'v4', credentials=creds)

# Connect to Gmail's IMAP server
def connect_to_gmail():
    print("Connecting to Gmail...")
    mail = imaplib.IMAP4_SSL('imap.gmail.com')
    mail.login(GMAIL_USER, GMAIL_PASS)
    print("Connected to Gmail.")
    return mail

# Function to safely decode a string
def safe_decode(s):
    try:
        return s.decode('utf-8')
    except UnicodeDecodeError:
        return s.decode('iso-8859-1')

# Function to clear the sheet except for the header
def clear_sheet():
    print(f"Clearing all data from {SHEET_NAME} except headers...")
    range_all = f"{SHEET_NAME}!A2:Z"
    service.spreadsheets().values().clear(spreadsheetId=SPREADSHEET_ID, range=range_all).execute()

# Function to insert data into the Google Sheet
def insert_into_sheet(date, subject, body):
    print("Inserting data into the sheet...")
    sheet = service.spreadsheets()
    values = [[date, subject, body]]
    body = {'values': values}
    range_to_append = f"{SHEET_NAME}!A:C"
    result = sheet.values().append(
        spreadsheetId=SPREADSHEET_ID,
        range=range_to_append,
        valueInputOption="USER_ENTERED",
        body=body).execute()
    print(f"{len(values)} row(s) appended.")

# Function to check if subject contains all required keywords
def subject_contains_keywords(subject, keywords):
    print("Checking if email subject contains all required keywords...")
    return all(keyword.lower() in subject.lower() for keyword in keywords)

# Function to get the email body
def get_email_body(part):
    print("Extracting email body...")
    try:
        body = part.get_payload(decode=True).decode('utf-8')
    except UnicodeDecodeError:
        body = part.get_payload(decode=True).decode('iso-8859-1')
    return body

# Function to get emails
def get_emails(mail):
    print("Fetching emails...")
    keywords = ["enviamos", "un", "pago"]

    mail.select('inbox')

    # Calculate the date range for the last month
    date_today = datetime.date.today()
    first_day_last_month = date_today.replace(day=1) - datetime.timedelta(days=1)
    first_day_month_before_last = first_day_last_month.replace(day=1)

    # Format dates for IMAP search
    date_format = "%d-%b-%Y"
    str_first_day_month_before_last = first_day_month_before_last.strftime(date_format)
    str_first_day_last_month = first_day_last_month.strftime(date_format)

    # Simplified version for testing: Search within a single month
    search_query = f'(SINCE "{str_first_day_month_before_last}" BEFORE "{str_first_day_last_month}")'

    status, messages = mail.search(None, search_query)
    if status != 'OK':
        print("No emails found in the specified date range.")
        return

    for num in messages[0].split():
        print(f"Processing email number {num.decode('utf-8')}...")
        status, data = mail.fetch(num, '(RFC822)')
        if status != 'OK':
            print(f"Failed to fetch email {num.decode('utf-8')}.")
            continue

        msg = email.message_from_bytes(data[0][1])
        from_ = decode_header(msg.get("From"))[0][0]
        if isinstance(from_, bytes):
            from_ = safe_decode(from_)
        subject = msg.get("Subject")
        if subject is not None:
            subject = decode_header(subject)[0][0]
            if isinstance(subject, bytes):
                subject = safe_decode(subject)
        else:
            subject = "No Subject"

        if not subject_contains_keywords(subject, keywords):
            print("Email does not contain the required keywords. Skipping...")
            continue

        # Extract the date in a human-readable format
        date_tuple = email.utils.parsedate_tz(msg.get('Date'))
        if date_tuple:
            local_date = datetime.datetime.fromtimestamp(email.utils.mktime_tz(date_tuple))
            str_date = local_date.strftime('%Y-%m-%d %H:%M:%S %z')  # Format the date as you like

        # Extracting the email body
        body_text = []
        for part in msg.walk():
            if part.get_content_type() == 'text/plain':
                body_text.append(get_email_body(part))
        body = ' '.join(body_text)

        # Insert data into Google Sheet
        insert_into_sheet(str_date, subject, body)

def main():
    mail = connect_to_gmail()
    clear_sheet()  # Clear the sheet before inserting new data
    get_emails(mail)
    mail.close()
    mail.logout()
    print("Disconnected from Gmail and finished processing emails.")

if __name__ == "__main__":
    main()
