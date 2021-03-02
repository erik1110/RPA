import smtplib
import ssl
import os
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

def sendmail3(email_user, email_pwd, subject, context, recipents, folder_path, filename):
    # https://realpython.com/python-send-email/#adding-attachments-using-the-email-package
    gmailUser = email_user
    gmailPasswd = email_pwd
    message = MIMEMultipart()
    message['From'] = email_user
    message['To'] = email_user    
    message['Subject'] = subject
    #message['Cc'] = cc
    #message['Bcc'] = cc
    # 增加內文
    message.attach(MIMEText(context, "html"))
    
    # 增加附件
    with open(f'./{folder_path}/{filename}', "rb") as attachment:
      # Add file as application/octet-stream
      # Email client can usually download this automatically as attachment
      part = MIMEBase("application", "octet-stream")
      part.set_payload(attachment.read())
    
    encoders.encode_base64(part)

    part.add_header(
        "Content-Disposition",
        f"attachment; filename= {filename}",
    )
    # Add attachment to message and convert message to string
    message.attach(part)
    text = message.as_string()
    # 寄送作業
    context = ssl.create_default_context()
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
        server.login(gmailUser, gmailPasswd)
        server.sendmail(message['From'], message['To'], text)
        print("寄信成功")
  
if __name__ == '__main__':
    email_user = os.environ["MAIL_USERNAME"]
    email_pwd = os.environ["MAIL_PASSWORD"]
    subject = 'Github Actions job result'
    context =  'github actction 測試'
    recipents =  os.environ["MAIL_ADDRESS"]
    print("recipents:", recipents)
    folder_path = 'data'
    filename = 'result.xlsx'
    sendmail3(email_user, email_pwd, subject, context, recipents, folder_path, filename)
