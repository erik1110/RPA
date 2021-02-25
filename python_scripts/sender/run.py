import smtplib
import ssl
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText



def sendmail3(email_user,email_pwd,subject,context,recipents,folder_path,filename):
    # https://realpython.com/python-send-email/#adding-attachments-using-the-email-package
    gmailUser = email_user
    gmailPasswd = email_pwd
    message = MIMEMultipart()
    message['From'] = email_user
    message['To'] = ",".join(recipents)    
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

if __name__ == '__main__':
  email_user = '${{secrets.MAIL_USERNAME}}'
  email_pwd = '${{secrets.MAIL_PASSWORD}}'
  subject = 'Github Actions job result'
  context = '${{ github.job }} job in worflow ${{ github.workflow }} of ${{ github.repository }} has ${{ job.status }}'
  recipents = '${{secrets.MAIL_ADDRESS}}'
  folder_path = 'data'
  filename = 'result.xlsx'
  sendmail3(email_user, email_pwd, subject, context, recipents, folder_path, filename)