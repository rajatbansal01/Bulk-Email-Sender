import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import pandas as pd

email_user = 'sender_email_address'
email_password = 'password'
e=pd.read_excel('emails.xlsx')
email_to_send=e['students'].values

subject =input("enter the subject to send\n")
body = input('what msg you want to send\n')
filename=input('enter file name first put it on desktop\n')
path='\\'+filename
attachment  =open(r'C:\Users\HP\Desktop'+path,'rb')


msg = MIMEMultipart()
msg['Subject'] = subject


msg.attach(MIMEText(body,'plain'))


part = MIMEBase('application','octet-stream')
part.set_payload((attachment).read())
encoders.encode_base64(part)
part.add_header('Content-Disposition',"attachment; filename= "+filename)

msg.attach(part)
text = msg.as_string()
server = smtplib.SMTP('smtp.gmail.com',587)
server.starttls()
server.login(email_user,email_password)
for i in email_to_send:
    server.sendmail(email_user,i,text)
server.quit()
print('email sent')
