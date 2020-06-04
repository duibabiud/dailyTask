import imaplib
import email
import os
from tkinter import *
import xlrd
#from tkinter import filedialog

def fetchAttachments():
    user = userEntry.get()
    password = passwordEntry.get()

    # Getting path From config excel file
    loc = ("./config.xlsx")
    wb = xlrd.open_workbook(loc)
    sheet = wb.sheet_by_index(0)
    attachment_dir = sheet.cell_value(1, 1)  # Getting the path from B2 cell of sheet 1.

    imap_url = 'imap.outlook.com'
    con = imaplib.IMAP4_SSL(imap_url)
    con.login(user, password)
    con.select('Inbox')

    type, msgs = con.search(None,'(SUBJECT "Resume")')
    msgs = msgs[0].split()
    count = 0

    for msg in msgs:
        resp, data = con.fetch(msg, "(RFC822)")
        email_body = data[0][1]
        m = email.message_from_bytes(email_body)

        if m.get_content_maintype() != 'multipart':
            continue

        for part in m.walk():
            if part.get_content_maintype() == 'multipart':
                continue
            if part.get('Content-Disposition') is None:
                continue

            filename = part.get_filename()
            if filename is not None:
                save_path = os.path.join(attachment_dir, filename)
                if not os.path.isfile(save_path):
                    #print(save_path)
                    count += 1
                    fp = open(save_path, 'wb')
                    fp.write(part.get_payload(decode=True))
                    fp.close()
    print("{} Files downloaded".format(count))

window = Tk()
window.configure(background="light green")
window.title("FETCHING EMAIL ATTACHMENTS")
window.geometry("300x200")

user_label = Label(window, text = "EMAIL  ", bg = "light green")
user_label.grid(row = 0, column = 0, padx = 0, pady = 10)
userEntry = Entry(window, width = 20)
userEntry.grid(row=0, column=1)

password_label = Label(window, text = "PASSWORD  ", bg = "light green")
password_label.grid(row = 1, column = 0, padx = 0, pady = 10)
passwordEntry = Entry(window, textvariable = 'password', show = "*",width = 20)
passwordEntry.grid(row = 1, column = 1)


submitButton = Button(window, text="Submit", command=fetchAttachments)
submitButton.grid(row = 3, column =1 , padx = 0, pady = 10,sticky='nesw')

window.mainloop()