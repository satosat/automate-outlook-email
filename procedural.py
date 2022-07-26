"""
Email outomation using Outlook
Credit: https://medium.com/@neonforge/how-to-send-emails-with-attachments-with-python-by-using-microsoft-outlook-or-office365-smtp-b20405c9e63a
"""

import smtplib
import csv
from email.mime.text import MIMEText    
from email.mime.multipart import MIMEMultipart

# Sender email and password
sender: dict = {
    'email': 'your email address',
    'password': 'your email password'
}

# Email subject and body
subject = "Subject"
body = """Email Body"""

# List to store failed recipients
failed_recipients = []

# File path to recipients 
# Use absolute path if unsure
recipients_file_path = "recipients.csv"

with open(recipients_file_path, 'r') as recipients_file:

    # Read the recipients
    recipients = csv.DictReader(recipients_file)

    # Get server connection
    try:
        print('Connecting to server')
        server = smtplib.SMTP('smtp.office365.com', 587)
        server.ehlo()
        server.starttls()
        server.login(sender['email'], sender['password'])
        print('Connected to server')
    except Exception as e:
        print(f'{e}')
        print("Failed connecting to server")
        exit(1)

    for recipient in recipients:
        email = MIMEMultipart()
        email['From'] = sender['email']
        email['To'] = recipient['email']
        email['Subject'] = subject
        email.attach(MIMEText(body, 'plain'))
        text = email.as_string()

        # Send email for every recipient
        try:
            server.sendmail(sender['email'], recipient['email'], text)
        except smtplib.SMTPException as e:
            print(f'{e}')
            print(f'Failed sending to {recipient["email"]}')

            # Append every failed recipient
            failed_recipients.append(recipient)
    
    # Quit server
    try:
        print("Quitting server")
        server.quit()
        print("Quit successful")
    except Exception as e:
        print(f'{e}')
    finally:
        # Print any failed recipients
        if len(failed_recipients):
            for recipient in failed_recipients:
                print(f'{recipient["name"]}: {recipient["email"]}')
