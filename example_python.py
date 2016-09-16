# -*- coding: utf-8 -*-
"""

@author: aclark

"""

# import libraries
# pyodbc for ODBC connection
import pyodbc
# pandas for 'excel' like sheets
import pandas as pd
# datetime for date computer date
import datetime

# create connection to AX 2009 database with Windows authentication
# if you do not use windows authentication, use the following instead:
# databaseConnection = 'DRIVER={SQL Server}; SERVER=ServerName; Database=DatbaseName; UID=UserId; PWD=password;'
cnxn = pyodbc.connect('DRIVER={SQL Server};SERVER=YourServerAddressHere;DATABASE=YourDatabaseHere;Trusted_Connection=yes;')
cursor = cnxn.cursor()

# SQL query for testing if a series of three
# 999 exist in a journal entry amount
JournalNines = '''
SELECT
	LEDGERTABLE.ACCOUNTNAME,
	LEDGERTRANS.ACCOUNTNUM,
	LEDGERTRANS.TRANSDATE,
	LEDGERTRANS.AMOUNTCUR,
	LEDGERTRANS.CREDITING,
	LEDGERTRANS.CURRENCYCODE,
	LEDGERTRANS.TRANSTYPE,
	LEDGERTRANS.DOCUMENTDATE,
	LEDGERTRANS.POSTING,
	LEDGERTRANS.DOCUMENTNUM,
	LEDGERTRANS.VOUCHER,
	LEDGERTRANS.PAYMREFERENCE,
	LEDGERTRANS.PERIODCODE,
	LEDGERTRANS.TXT,
	LEDGERTRANS.CREATEDBY,
	USERINFO.NAME,
	LEDGERTRANS.CREATEDDATETIME,
	LEDGERTRANS.MODIFIEDDATETIME
FROM	LEDGERTRANS
LEFT JOIN LEDGERTABLE
	ON LEDGERTABLE.ACCOUNTNUM = LEDGERTRANS.ACCOUNTNUM
LEFT JOIN USERINFO
	ON USERINFO.ID = LEDGERTRANS.CREATEDBY
WHERE
	CONVERT(varchar(30), LEDGERTRANS.AMOUNTCUR) LIKE '%999%' AND
	LEDGERTRANS.TRANSDATE > (GetDate()-31);
'''

# save the sql query to a pandas dataframe, or excel like sheet
Journal999 = pd.read_sql(JournalNines,cnxn)

# returns all of the journal entries that occur on the weekend
Weekends = '''
SELECT
	LEDGERTABLE.ACCOUNTNAME,
	LEDGERTRANS.ACCOUNTNUM,
	LEDGERTRANS.TRANSDATE,
	DATENAME(dw,LEDGERTRANS.CREATEDDATETIME) AS Day,
	LEDGERTRANS.AMOUNTCUR,
	LEDGERTRANS.CREDITING,
	LEDGERTRANS.CURRENCYCODE,
	LEDGERTRANS.TRANSTYPE,
	LEDGERTRANS.DOCUMENTDATE,
	LEDGERTRANS.POSTING,
	LEDGERTRANS.DOCUMENTNUM,
	LEDGERTRANS.TXT,
	LEDGERTRANS.CREATEDBY,
	USERINFO.NAME,
	LEDGERTRANS.CREATEDDATETIME,
	LEDGERTRANS.MODIFIEDDATETIME
FROM	LEDGERTRANS
LEFT JOIN LEDGERTABLE
	ON LEDGERTABLE.ACCOUNTNUM = LEDGERTRANS.ACCOUNTNUM
LEFT JOIN USERINFO
	ON USERINFO.ID = LEDGERTRANS.CREATEDBY
WHERE
	LEDGERTRANS.CREATEDDATETIME > (GetDate()-31) AND DATENAME(dw,LEDGERTRANS.CREATEDDATETIME) IN ('Saturday','Sunday')
ORDER BY
	TRANSDATE DESC;
    '''

# save the sql query to a pandas dataframe, or excel like sheet
WeekendEntries = pd.read_sql(Weekends, cnxn)

# test looks at the journal entry descriptions and returns results
# that have fraudulent keywords. This is a simple example, you could search for
# thousands of keywords with slightly different code
Keywords = '''

SELECT
	LEDGERTABLE.ACCOUNTNAME,
	LEDGERTRANS.ACCOUNTNUM,
	LEDGERTRANS.TRANSDATE,
	LEDGERTRANS.AMOUNTCUR,
	LEDGERTRANS.CREDITING,
	LEDGERTRANS.CURRENCYCODE,
	LEDGERTRANS.TRANSTYPE,
	LEDGERTRANS.DOCUMENTDATE,
	LEDGERTRANS.POSTING,
	LEDGERTRANS.DOCUMENTNUM,
	LEDGERTRANS.VOUCHER,
	LEDGERTRANS.PAYMREFERENCE,
	LEDGERTRANS.TXT,
	LEDGERTRANS.CREATEDBY,
	USERINFO.NAME,
	LEDGERTRANS.CREATEDDATETIME,
	LEDGERTRANS.MODIFIEDDATETIME
FROM	LEDGERTRANS
LEFT JOIN LEDGERTABLE
	ON LEDGERTABLE.ACCOUNTNUM = LEDGERTRANS.ACCOUNTNUM
LEFT JOIN USERINFO
	ON USERINFO.ID = LEDGERTRANS.CREATEDBY
WHERE
	LEDGERTRANS.CREATEDDATETIME > (GetDate()-31) AND (LEDGERTRANS.TXT LIKE '%fraud%' OR LEDGERTRANS.TXT LIKE '%bribe%'
	OR LEDGERTRANS.TXT LIKE '%corruption%' OR LEDGERTRANS.TXT LIKE '%plug%' OR LEDGERTRANS.TXT LIKE '%kickback%'
	OR LEDGERTRANS.TXT LIKE '%payola%' OR LEDGERTRANS.TXT LIKE '%sweetener%' OR LEDGERTRANS.TXT LIKE '%backhander%'
	OR LEDGERTRANS.TXT LIKE '%hush money%' OR LEDGERTRANS.TXT LIKE '%grease%' OR LEDGERTRANS.TXT LIKE '%wet my beak%')
'''


KeyWords = pd.read_sql(Keywords,cnxn)

# closes the database connections
cnxn.close()

# gets the current year
year = datetime.datetime.now().year
# gets the current month
month = datetime.datetime.now().month

# combines current year and month and the desired file name to
# 2016_12_MISTI_Analytics_Example.xlsc
fileName = str(year) + '_' + str(month) + '_MISTI_Analytics_Example.xlsx'

writer = pd.ExcelWriter(fileName)

def ExcelExport(inputName, outputName):
    ''' This function takes the dataframe as an input and the
desired name of the workbook sheet name as the output. If the dataframe
doesn't have any data in it, the sheet will not be added to the workbook'''
    if inputName.empty:
       pass
    else:
       inputName.to_excel(writer, outputName)



ExcelExport(Journal999,'Journal999')
ExcelExport(WeekendEntries,'WeekendEntries')
ExcelExport(KeyWords,'KeyWords')
writer.save()

# Add where you the python file (and this program export) is located.
# Be sure to include two backslashes, as this is a python idiosyncrasy
email_path = "C:\\Users\\aclark\\Desktop\\" + fileName

# email results as excel attachment, if you use Outlook.
import win32com.client
const=win32com.client.constants
olMailItem = 0x0
obj = win32com.client.Dispatch("Outlook.Application")
newMail = obj.CreateItem(olMailItem)
newMail.Subject = "Monthly Analytic Test Results"
newMail.Body = '''Attached are the results of the monthly analytics program.
\nSincerely, \n\nAuditMachine'''
newMail.To = "YourEmailHere" # put the email you would like to send to here
attachment1 = email_path

newMail.Attachments.Add(Source=attachment1)

newMail.Send()
