import imaplib
import email
import os
from credentials import useName, passWord

imap_url = 'imap.gmail.com'
my_mail = imaplib.IMAP4_SSL(imap_url)
my_mail.login(useName, passWord)

my_mail.select('inbox')
# print(my_mail.select('inbox'))

data = my_mail.search(None, 'ALL')
mail_ids = data[1]
id_list = mail_ids[0].split()
# print(id_list)
first_email_id = int(id_list[0])  #default values # 0 
latest_email_id = int(id_list[-1])  #default values # -1
print("1st pass")
for i in range(latest_email_id, first_email_id, -1):
    print(f"Processing email with UID: {i}")
    data = my_mail.fetch(str(i), '(RFC822)')
    print("2nd pass")
    for response_part in data:
        arr = response_part[0]
        if isinstance(arr, tuple):
            msg = email.message_from_bytes(arr[1])
            # print(msg)
            email_subject = msg['subject']
            email_from = msg.get('from', '')
            email_date = msg['Date']
            
            print(f'From: {email_from}')
            print('Subject: ' + email_subject)
            print('Date: ' + email_date + '\n')
            print("3rd pass")

            if 'Google' in email_from:   #instead of google you can any you specific keywords
                print("4th pass")
                for part in msg.walk():
                    if part.get_content_maintype() == 'multipart':
                        continue
                    if part.get('Content-Disposition') is None:
                        continue
                    print("5th pass")

                    if part.get_content_type() == 'application/pdf':
                        fileName = part.get_filename()
                        filePath = os.path.join('E:/invoice2data/emailfiles', fileName)  #To download pdfs
                        print("6th pass")
                        if not os.path.isfile(filePath):
                            with open(filePath, 'wb') as fp:
                                fp.write(part.get_payload(decode=True))
                            print(f"Downloaded: {fileName}")
                            print("7th pass")

# Close the mailbox and logout
my_mail.close()
my_mail.logout()
