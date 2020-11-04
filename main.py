# Created by Marios Paraskevas on November 2020
# Import Modules
import os
import datetime
import pandas as pd
from pathlib import Path
from twilio.rest import Client
from dotenv import load_dotenv

# Loading Environment Variables i.e. Twilio_account_sid and Twilio_auth_token
load_dotenv()

# Opening the excel warehouse inspection excel file and specifically only the first sheet 
# This workbook is being referred as wb 
wb = pd.read_excel(r"C:\Users\marios\Desktop\inspector\WH_inspection.xlsx", "Sheet1")
# Getting a copy of the columns with the expiration dates and the one with the items' description but deleting the N/A rows
col_with_expD = wb.iloc[:, 4].dropna().copy()
col_with_description = wb.iloc[:,0].dropna().copy()
# Getting the indexes of the items with expiration dates and setting up the new dataframes
ind = col_with_expD.index
des = col_with_description.loc[ind]
dates = list(pd.to_datetime(col_with_expD, errors="coerce").dt.date)

# Creating 2 lists with the items with the corresponding expiration dates and then making a dictionary of those
descr = list(des)
expD = [date.strftime('%Y-%m-%d') for date in dates]
table = dict(zip(descr,expD))

# Setting the Twilio Whatsapp API request
account_sid = os.environ['TWILIO_ACCOUNT_SID']
auth_token = os.environ['TWILIO_AUTH_TOKEN']
client = Client(account_sid, auth_token)

# Getting today's date for date relating 
today = datetime.datetime.now().strftime("%Y-%m-%d")

# Initializing an expired items dictionary and the bodyText of the message
bodyText = ""
expired = []

# Checking what has already expired
for k,v in table.items():
    if datetime.datetime.strptime(v,"%Y-%m-%d") < datetime.datetime.strptime(today,"%Y-%m-%d"):
        bodyText += '\nΤο ' + k + ' έληξε στις ' + v
        expired.append(k)

# Deleting the expired items from the the dictionary
for j in expired:
    table.pop(j)

# After-30-days date                
days30 = datetime.datetime.today() + datetime.timedelta(30)

for k,v in table.items():
    if datetime.datetime.strptime(v,"%Y-%m-%d") < days30:
        bodyText += '\nΤο ' + k + ' θα λήξει σε λιγότερο από 30 ημέρες (' + v + ')'

# Sending the message to Whatsapp
message = client.messages.create(
                              body=bodyText,
                              from_='whatsapp:{}'.format(os.environ['TWILIO_NUMBER']),
                              to='whatsapp:{}'.format(os.environ['MY_NUMBER'])
                          )

