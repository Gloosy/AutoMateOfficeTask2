import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from .config import EMAIL_ADDRESS, EMAIL_PASSWORD, EMAIL_TEMPLATE_PATH

def send_email(recipient_email, subject, cc_emails='', attachment_path=''):
    """
    Sends an email to the specified recipient with optional CC and attachment.
    """
    # Create a message object
    msg = MIMEMultipart()
    msg['From'] = EMAIL_ADDRESS
    msg['To'] = recipient_email.replace(';', ',')
    msg['Cc'] = cc_emails.replace(';', ',')
    msg['Subject'] = subject

    # Email body content from template
    with open(EMAIL_TEMPLATE_PATH, 'r', encoding='utf-8') as file:
        body = file.read()
    msg.attach(MIMEText(body, 'html'))

    # Attachment
    if os.path.isfile(attachment_path):
        filename = os.path.basename(attachment_path)
        with open(attachment_path, "rb") as attachment:
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header(
            'Content-Disposition',
            f"attachment; filename= {filename}",
        )
        msg.attach(part)

    # Connect to SMTP server and send email
    try:
        server = smtplib.SMTP('smtp.office365.com', 587)
        server.starttls()
        server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
        server.send_message(msg)
        server.quit()
        print(f"Email sent to {recipient_email}")
    except Exception as e:
        print(f"Failed to send email to {recipient_email}: {e}")
