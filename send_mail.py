import smtplib
import os
from dotenv import load_dotenv
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

load_dotenv()


def send_mail(to: list[str], from_email: list[str], password: str, cc: list[str], subject: str, body: str, file_path=None):
    msg = MIMEMultipart()
    msg['From'] = from_email
    msg['To'] = ', '.join(to)
    msg['Cc'] = ', '.join(cc)
    msg['Subject'] = subject

    text = MIMEText(body, 'html')
    msg.attach(text)

    if file_path:
        with open(file_path, "rb") as f:
            attach = MIMEApplication(f.read(), _subtype="xlsx")
            attach.add_header('Content-Disposition', 'attachment', filename=os.path.basename(file_path))
            msg.attach(attach)
        os.remove(file_path)

    server = smtplib.SMTP('smtp-mail.outlook.com', 587)
    server.starttls()
    server.login(from_email, password)
    status = server.sendmail(from_email, msg['To'].split(', ') + msg['Cc'].split(', '), msg.as_string())
    server.quit()
    if not status:
        print(f'Mail has been successfully sent to {to[0]}')
