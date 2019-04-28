import smtplib
import imapclient
import pyzmail
import pprint
import bs4
import urllib

from openpyxl import Workbook

username = 'karthiek.umich@gmail.com'
#password = input("enter password : ")
password = 'amgmercedesM1596.2'

# try:
#     smtpobj = smtplib.SMTP('smtp.gmail.com', 587)
#     smtpobj.starttls()
#     smtpobj.login(username,password)
#     for i in range(0,10):
#         smtpobj.sendmail(username,'knaaga90@gmail.com','Hello Buddy')
#     smtpobj.quit()
#
# except:
#     print("Something went wrong...")

imapobj = imapclient.IMAPClient('imap.gmail.com', ssl=True)
imapobj.login(username, password)

imapobj.select_folder('INBOX', readonly=True)
UIDs = imapobj.search(['BODY', 'twitter', 'ON', '07-Sep-2018', 'SUBJECT', 'Oliver'])
print (UIDs)
rawMessages = imapobj.fetch(39452, ['BODY.PEEK[HEADER]'])
message = pyzmail.PyzMessage.factory(rawMessages[39452][b'BODY[HEADER]'])
for line in (str(message).splitlines()):
    if 'List-Unsubscribe: <' in line:
        print (line)
        break


wb = Workbook()
wb.save('Email_Analytics.xlsx')

imapobj.logout()