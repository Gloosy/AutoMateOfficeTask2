import os
import json
from pandas import read_excel
from datetime import date

from smtplib import SMTP
from msal import ConfidentialClientApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from base64 import b64encode




def send_email(recipient_email, subject, cc_emails='', file_attached=''):
    """

    """

    # Get email credentials from environment
    email = os.getenv('EMAIL_ADDRESS')
    password = os.getenv('EMAIL_PASSWORD')
    print('email    :',email)

    ## Create a message object
    msg = MIMEMultipart()
    msg['From'] = email
    msg['To'] = recipient_email.replace(';',',')
    msg['Cc'] = cc_emails.replace(';',',')
    msg['Subject'] = subject

    # Email body content
    body = open('template.html',mode='r',encoding='utf8').read()
    msg.attach(MIMEText(body, 'html'))

    ## Attachment File
    if os.path.isfile(file_attached):
        filename = os.path.basename(file_attached)
        # Open the file in binary mode and attach it
        # Create MIMEBase object for attachment
        try:
            with open(file_attached, mode="rb") as fr:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(fr.read())
                encoders.encode_base64(part) # Encode to base64
                part.add_header('Content-Disposition', f"attachment; filename= {filename}")
                msg.attach(part) # Attach the file to the message
        except Exception as e:
            print(f"Failed to attach file. Error: {str(e)}")

    ## Connect to the Outlook SMTP server
    try:
        server = SMTP('smtp.office365.com', 587)
        print(f'server res{server}')
        status, response = server.ehlo()        # Send EHLO command
        print('Ehlo     :',status, response)
        status, response = server.starttls()    # Start TLS encryption
        print('Starttls :',status, response)
        status, response = server.ehlo()        # Send EHLO again after TLS starts (required by some servers)
        print('Ehlo     :',status, response)

        status, response = server.login(email, password)
        print('Login    :',status, response)

        if status < 500:
            server.sendmail(msg['From'], msg['To'], msg.as_string())
            print("Success!")
    except Exception as e:
        print(f"Failed to send email. Error: {str(e)}")
    finally:
        server.quit()




def distribute_emails(filepath):
    df = read_excel(filepath, sheet_name='Email_List', header=0)
    print('Shape:',df.shape)
    # print(df.info())
    # print(df)

    folder = os.path.dirname(filepath)
    for i,row in df.iterrows():
        if row['Active']:
            name = row['Recipient Name']
            recipient_email = row['Recipient Email']
            cc_emails = row['CC Email']
            filename = row['Attachment File']
            subject = f'Financial Data {date.today()}'
            file_attached = os.path.join(folder, 'Attachments', filename)
            print(f"\n{name}: {filename}")
            send_email(recipient_email, subject, cc_emails, file_attached)
            break




if __name__ == '__main__':
    os.system('cls||clear')

    # Test Code
    this_path = os.path.dirname(os.path.realpath(__file__))
    filepath = os.path.join(this_path, "Financial_Data.xlsx")
    distribute_emails(filepath)
