import imaplib
import email
import os
import xlrd

user = input("Plesae enter Outlook email id : ")
password = input("Please enter your password : ")
imap_url = 'imap.outlook.com'

#Getting data From config excel file
loc = ("./config.xlsx")
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)
attachment_dir  = sheet.cell_value(1, 1) # Getting the path from B2 cell of shhet 1.

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
                count += 1
                fp = open(save_path, 'wb')
                fp.write(part.get_payload(decode=True))
                fp.close()
print("{} Files downloaded".format(count))
