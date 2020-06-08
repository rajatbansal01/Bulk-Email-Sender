import pandas as pd
import smtplib
e=pd.read_excel('emails.xlsx') #here provide the excel file which is contaning the email addresses.
emails_to_send=e['students'].values #students is the header for the emails.
print(emails_to_send)
server = smtplib.SMTP('smtp.gmail.com' , 587) #starting the SMTP server.
server.starttls()
server.login('sender_email_address','password_for_login')
msg='the meassage you want to send'
sub="SUBJECT FOR THE EMAIL"
body="Subject:{}\n\n{}".format(sub,msg)
for i in emails_to_send:
    server.sendmail('sender_email_address',i,body)
print('email sent')
server.quit() #quit the server here.
