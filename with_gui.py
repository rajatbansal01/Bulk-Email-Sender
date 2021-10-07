import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import pandas as pd
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog

filename = None

def browseFiles():
    global filename
    temp_filename = filedialog.askopenfilename(initialdir = "/",
        title = "Select a File"
    )
    if temp_filename:
        filename = temp_filename
        browseFileLabel.configure(text="File Opened: "+filename)


def send():
    progressBar.grid(column=0, row=0)
    root.update()
    subject = subjectEntry.get(1.0, 'end')
    body = bodyEntry.get(1.0, 'end')

    msg = MIMEMultipart()
    msg['Subject'] = subject

    msg.attach(MIMEText(body,'plain'))

    if filename:
        attachment  =open(filename,'rb')
        part = MIMEBase('application','octet-stream')
        part.set_payload((attachment).read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition',"attachment; filename= "+filename.split("/")[-1])
        msg.attach(part)
    text = msg.as_string()
    server = smtplib.SMTP('smtp.gmail.com',587)
    server.starttls()
    try:
        server.login(email_user,email_password)
    except:
        progressBar.grid_forget()
        progressBar['value'] = 0
        tk.messagebox.showerror(title="Error", message="Bad Credentials")
        return
    count = 0
    for i in email_to_send:
        root.update()
        server.sendmail(email_user,i,text)
        count+=1
        progressBar['value'] = (count / len(email_to_send)) * 100

    server.quit()
    progressBar.grid_forget()
    progressBar['value'] = 0
    tk.messagebox.showinfo(title='Success', message="Emails sent successfully.")

email_user = 'sender_email_address'
email_password = 'password'
e=pd.read_excel('emails.xlsx')
email_to_send=e['students'].values

root = tk.Tk()
root.title("Email Sender")

# Subject
subjectFrame = tk.Frame(root)
subjectFrame.pack(padx=10, pady=10)

tk.Label(subjectFrame, text="Enter subject:").grid(row=0, column=0)

subjectEntry = tk.Text(subjectFrame, height=2)
subjectEntry.grid(row="0", column="1")

# Body
bodyFrame = tk.Frame(root)
bodyFrame.pack(padx=10, pady=10)

tk.Label(bodyFrame, text="Enter body:").grid(row=0, column=0)

bodyEntry = tk.Text(bodyFrame)
bodyEntry.grid(row="0", column="1")

# Browse Files
browseFileFrame = tk.Frame(root)
browseFileFrame.pack(padx=10, pady=10)

browseFileLabel = tk.Label(browseFileFrame, text = "No File Selected")
browseFileLabel.grid(row=0, column=0)
  
      
button_explore = tk.Button(browseFileFrame, text = "Browse Files", command = browseFiles).grid(row=0, column=2)


# Submit Buttons
submitFrame = tk.Frame(root)
submitFrame.pack(padx=10, pady=10)

submitEntry = tk.Button(submitFrame, text="Go", command=send)
submitEntry.grid(row="0", column="2")

# progressbar
progressBarFrame = tk.Frame(root)
progressBarFrame.pack(padx=10, pady=10)
progressBar = ttk.Progressbar(
    progressBarFrame,
    orient='horizontal',
    length=280,
)

root.mainloop()

