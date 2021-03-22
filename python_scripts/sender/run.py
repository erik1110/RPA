#!/usr/bin/python
# coding:utf-8
'''
sending email
'''
import os
import ssl
import smtplib
from datetime import datetime
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

def send_mail(email_user, email_pwd, subject, context, recipents, carbon_copy=None, filepath=None):
    '''
    send email
    Args:
        email_user: email account
        email_pwd: email password
        subject: email title
        context: email body
        recipents: email recipient
        carbon_copy: email cc, default None
        fiilpath: the attachment path, default None
    Return:
        True: Sending successfully
        False: Sending unsuccessfully
    '''
    gmailUser = email_user
    gmailPasswd = email_pwd
    message = MIMEMultipart()
    message['From'] = email_user
    message['Subject'] = subject
    try:
        # Add Body
        message.attach(MIMEText(context, "html"))
        # Add attechment
        if filepath:
            with open(filepath, "rb") as attachment:
            # Add file as application/octet-stream
            # Email client can usually download this automatically as attachment
                part = MIMEBase("application", "octet-stream")
                part.set_payload(attachment.read())
            encoders.encode_base64(part)
            part.add_header("Content-Disposition", 
                            "attachment",
                            filename=("gbk", "", os.path.basename(filepath)))
            # Add attachment to message and convert message to string
            message.attach(part)
        # meassage to string
        text = message.as_string()
        # Add cc
        if carbon_copy:
            message['Cc'] = carbon_copy.split(';')
        # Add recepients
        message['To'] = recipents.split(';')
        # Sending
        context = ssl.create_default_context()
        with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
            server.login(gmailUser, gmailPasswd)
            server.sendmail(message['From'], message['To'], text)
            print("Sending email successfully")
            return True
    except Exception as error:
        print("Error:", error)
        return False
  
if __name__ == '__main__':
    email_user = os.environ["MAIL_USERNAME"]
    email_pwd = os.environ["MAIL_PASSWORD"]
    subject = 'Github Actions job result'
    context =  'github actction 測試'
    recipents =  os.environ["MAIL_ADDRESS"]
    carbon_copy = os.environ["MAIL_ADDRESS"]
    filepath = './data/result.xlsx'
    send_mail(email_user, email_pwd, subject, context, recipents, carbon_copy, filepath)
