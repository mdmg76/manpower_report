#!python
# auto_sotrovimab_reporter.py automatically sends the daily sotrovimab report to all assigned recipients.

import os
from pathlib import Path
from datetime import datetime
import openpyxl
from exchangelib import Credentials, Account, Message, Mailbox, HTMLBody, FileAttachment

today = datetime.today().strftime('%d-%m-%Y')
sheet_date = datetime.today().strftime('%d')
sheet_month = datetime.today().strftime('%b')
# '''
report_loc = Path('//rah-fileserver/Users/Rahba Shared Folder/Pharmacy/Pharmacy Reports/ARH Pharmacy Daily Manpower Updates/')
report_fil = [os.path.basename(name) for name in list(report_loc.glob(f'* {sheet_month} *'))][0]
month_wb = openpyxl.load_workbook(report_loc / report_fil)
for day in month_wb.sheetnames:
    if day != sheet_date:
        month_wb.remove(month_wb[day])
        month_wb.save(report_fil)

credentials = Credentials('mmughazi@thelifecorner.ae', '114622@PharmD')
account = Account('mmughazi@thelifecorner.ae', credentials=credentials, autodiscover=True)

m = Message(
    account=account,
    subject='RE: pharmacy daily manpower updates',
    body=HTMLBody('''
    <html>
        <body style="font-family:Segoe UI; color:#228B99">
            <p>Good Morning Sara,</p>
            <p>Kindly find Al Rahba manpower report for today {today}.</p>
            <p>Regards,</p>
        </body>
    </html>
    '''.format(today=today)),
    to_recipients=[Mailbox(email_address='sjjaberi@thelifecorner.ae')
    ],

    cc_recipients=['maltunaiji@thelifecorner.ae', 'msyed@thelifecorner.ae'], 
)
my_file = FileAttachment(name=report_fil, content=open(report_fil, 'rb').read())
m.attach(my_file)
m.send()
os.remove(report_fil)
