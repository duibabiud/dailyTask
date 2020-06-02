import imaplib
import email
import os

user = input("Plesae enter Gmail id : ")
password = input("Please enter your password : ")
attachment_dir = input("Enter path of directory : ")
imap_url = 'imap.gmail.com'

con = imaplib.IMAP4_SSL(imap_url)
con.login(user, password)
con.select('Inbox')

type, msgs = con.search(None, 'UNSEEN', '(SUBJECT "Resume")')
msgs = msgs[0].split()

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
                print(save_path)
                fp = open(save_path, 'wb')
                fp.write(part.get_payload(decode=True))
                fp.close()
