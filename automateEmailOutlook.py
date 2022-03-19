import smtplib
import csv
import os
from email.mime.text import MIMEText    
from email.mime.multipart import MIMEMultipart

class Email:
    def __init__(self, recipient: str, subject: str, body: str):
        self.message = MIMEMultipart()
        self.message['From'] = os.environ["SENDER_EMAIL"]
        self.message['To'] = recipient
        self.message['Subject'] = subject
        self.message.attach(MIMEText(body, 'plain'))
        self.message.as_string()


def main():
    # Check if env variables are set
    if 'SENDER_EMAIL' not in os.environ or 'SENDER_PASSWORD' not in os.environ:
        print('OS Envrionment variables not set')
        exit(1)

    # Create email subject and body
    subject = """Subject goes here"""
    body = """Body goes here"""

    # List variable to store failed recipients' emails
    failed_recipients = []

    # Read through recipients csv file
    with open('recipients.csv', 'r') as recipients_file:
        recipients = csv.DictReader(recipients_file)

        # Get server connection
        server = get_server()

        for recipient in recipients:
            # Create new Email instance
            email = Email(recipient['email'], subject, body)

            # Send email
            print(f'Sending to {email.message["To"]}')
            send_email(server, email, failed_recipients)
        
        server.quit()
    
    if len(failed_recipients):
        print('List of failed recipients:')

        for recipient in failed_recipients:
            print(recipient)
            
# Create new SMTP connection, will stop program if failure occured
def get_server() -> smtplib.SMTP:
    try:
        server = smtplib.SMTP('smtp.office365.com', 587)
        server.ehlo()
        server.starttls()
        server.login(os.environ['SENDER_EMAIL'], os.environ['SENDER_PASSWORD'])
        return server
    except Exception as err:
        print(format(err))
        print("Server connection failed")
        exit(1)

def send_email(server: smtplib.SMTP, email: Email, failed: list):
    try:
        server.sendmail(email.message['From'], email.message['To'], email.message)
        print('email sent')
    except Exception as err:
        print(format(err))
        print(f"Failed sending to {email.message['To']}")

        # Failed recipients will be stored
        failed.append(email.message['To'])


if __name__ == "__main__":
    main()