import smtplib
import ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import os

def send_email(result):
    SMTP_SERVER = os.getenv('SMTP_SERVER')
    SMTP_PORT = int(os.getenv('SMTP_PORT'))
    SMTP_USER = os.getenv('SMTP_USER')
    SMTP_PASS = os.getenv('SMTP_PASS')
    MAIL_TO = os.getenv('MAIL_TO')

    msg = MIMEMultipart()
    msg['From'] = SMTP_USER
    msg['To'] = MAIL_TO
    msg['Subject'] = f"Relatório Climático - {result['city']}"

    body = (
        f"📍 Cidade: {result['city']}\n"
        f"🌡️ Temperatura: {result['temperature']}°C\n"
        f"💧 Umidade: {result['humidity']}%\n"
        f"🌤️ Clima: {result['description']}\n"
        f"Relatado em: {result['datetime']}\n"
    )

    msg.attach(MIMEText(body, 'plain'))

    filename = 'weather_log.csv'
    with open(filename, 'rb') as attachment:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f'attachment; filename= {filename}')
    msg.attach(part)

    context = ssl.create_default_context()

    with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT, context=context) as server:
        server.login(SMTP_USER, SMTP_PASS)
        server.sendmail(SMTP_USER, MAIL_TO, msg.as_string())

    print('Email enviado com sucesso!')