import smtplib
import ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from dotenv import load_dotenv
import os
load_dotenv()
# Create empty e-mail
def build_email(subject, body):
    msg = MIMEMultipart()
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))
    return msg

# Add attachments to e-mail
def add_attachments(msg, attachment_list):
    for file in attachment_list:
        if os.path.exists(file):
            try:
                with open(file, 'rb') as attach:
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(attach.read())
                    encoders.encode_base64(part)
                    part.add_header(
                        'Content-Disposition',
                        f'attachment; filename="{os.path.basename(file)}"'
                    )
                    msg.attach(part)
            except FileNotFoundError:
                print(f'File not found for {file}')
            except Exception as e:
                print (f'Error {e}')
        else:
            print(f'File does not exist: {file}')
    return msg

# Send e-mail
def send_email(msg, receiver):
    email_user = os.getenv('SMTP_USER')
    email_pass = os.getenv('SMTP_PASS')
    SMTP_SERVER = os.getenv('SMTP_SERVER')
    SMTP_PORT = int(os.getenv('SMTP_PORT'))

    msg['From'] = email_user
    msg['To'] = receiver

    context = ssl.create_default_context()

    try:
        with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT, context=context) as server:
            server.login(email_user, email_pass)
            server.sendmail(email_user, receiver, msg.as_string())
            return True
    except Exception as e:
        print(f'Error: {e}')
        return False
    
    
def send_weather_report(result, pdf_path=None, xlsx_path=None):
    subject = f'Weather report from: {result["city"]}'
    body =  f'''
        City name: {result['city']}
        Temperature: {result['temperature']}°C
        '''.strip()
    msg = build_email(subject=subject, body=body)
    attachments_list = []
    if pdf_path:
        attachments_list.append(pdf_path)
    if xlsx_path:
        attachments_list.append(xlsx_path)

    msg = add_attachments(msg=msg, attachment_list=attachments_list)

    receiver = os.getenv('MAIL_TO')
    send = send_email(msg=msg, receiver=receiver)

    return send
