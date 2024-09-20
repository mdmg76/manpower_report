#!python


import os
from pathlib import Path
from datetime import datetime
import openpyxl
from exchangelib import Credentials, Account, Message, Mailbox, HTMLBody, FileAttachment

today = datetime.today().strftime('%d-%m-%Y')
sheet_date = datetime.today().strftime('%d')
sheet_month = datetime.today().strftime('%b')
# '''
report_loc = Path('//network/path/to/Daily Manpower Updates/') #change to actual path
report_fil = [os.path.basename(name) for name in list(report_loc.glob(f'* {sheet_month} *'))][0]
month_wb = openpyxl.load_workbook(report_loc / report_fil)
for day in month_wb.sheetnames:
    if day != sheet_date:
        month_wb.remove(month_wb[day])
        month_wb.save(report_fil)

credentials = Credentials('your email', 'ypur email password') # better to udpate to use .env
account = Account('your email', credentials=credentials, autodiscover=True)

m = Message(
    account=account,
    subject='RE: pharmacy daily manpower updates',
    body=HTMLBody('''
    <html>
        <body style="font-family:Segoe UI; color:#228B99">
            <p>Good Morning,</p>
            <p>Kindly find <function name> manpower report for today {today}.</p>
            <p>Regards,</p>
        </body>
    </html>
    '''.format(today=today)),
    to_recipients=[Mailbox(email_address='recipient email')
    ],

    cc_recipients=['cc1 email', 'cc2 email'], 
)
my_file = FileAttachment(name=report_fil, content=open(report_fil, 'rb').read())
m.attach(my_file)
m.send()
os.remove(report_fil)
