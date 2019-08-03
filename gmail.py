import imapclient
import pyzmail
import time
import datetime
import calendar

from openpyxl import Workbook
from openpyxl import load_workbook

username = 'karthiek.umich@gmail.com'
#password = input("enter password : ")
password = 'amgmercedesM1596.2'

imapobj = imapclient.IMAPClient('imap.gmail.com', ssl=True)
imapobj.login(username, password)

start_time = time.time()

imapobj.select_folder('INBOX', readonly=True)
UIDs = imapobj.search(['SINCE', '01-Jul-2019', 'BEFORE', '30-Jul-2019'])
print (str(len(UIDs)) + " emails recieved in this time frame")

from_addresses = []
subjects = []
dates = []
days = []
months = []
years = []
times = []
unsub_links = [None]*20000

for i in range(len(UIDs)):
    raw_message = imapobj.fetch(UIDs[i], ['BODY.PEEK[HEADER]'])
    message = pyzmail.PyzMessage.factory(raw_message[UIDs[i]][b'BODY[HEADER]'])
    from_addresses.append(message.get_address('from'))
    subjects.append(message.get_subject(''))
    full_date = message.get_decoded_header('date')

    if (',' in full_date):
        day = full_date.split(', ')[0]
        date = full_date.split(', ')[1].split()[0]
        if date[0] == '0':
            date = date.split()[0][1:]
        month = full_date.split(', ')[1].split()[1]
        year = full_date.split(', ')[1].split()[2]
        time2 = full_date.split(', ')[1].split()[3]
    else:
        date = full_date.split()[0]
        if date[0] == '0':
            date = date.split()[0][1:]
        month = full_date.split()[1]
        year = full_date.split()[2]
        time2 = full_date.split()[3]
        month_no = list(calendar.month_abbr).index(month)
        day = (calendar.day_name[datetime.date(int(year), int(month_no), int(date)).weekday()]).split()[0][0:3]

    days.append(day)
    dates.append(date)
    months.append(month)
    years.append(year)
    times.append(time2)

    unsub_links[i] = message.get_decoded_header('List-Unsubscribe')

end_time = time.time()
print ("execution time : " + str((end_time-start_time)))

imapobj.logout()

wb = Workbook()
ws = wb.active
ws.title = "Data"
ws.cell(1,1).value = "Date"
ws.cell(1,2).value = "Month"
ws.cell(1,3).value = "Year"
ws.cell(1,4).value = "Day"
ws.cell(1,5).value = "Time"
ws.cell(1,6).value = "From (Sender)"
ws.cell(1,7).value = "From (Email ID)"
ws.cell(1,8).value = "Subject"
ws.cell(1,9).value = "Unsubscribe Link"

print(len(unsub_links))
for i in range(len(UIDs)):

    ws.cell(row=i+2, column=1).value = dates[i]
    ws.cell(row=i+2, column=2).value = months[i]
    ws.cell(row=i+2, column=3).value = years[i]
    ws.cell(row=i+2, column=4).value = days[i]
    ws.cell(row=i+2, column=5).value = times[i]
    ws.cell(row = i+2, column = 6).value = from_addresses[i][0]
    ws.cell(row = i+2, column = 7).value = from_addresses[i][1]
    ws.cell(row = i+2, column = 8).value = str(subjects[i])
    ws.cell(row=i+2, column=9).value = unsub_links[i]
wb.save('Email_Analytics.xlsx')