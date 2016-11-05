#Andrew Clark

#install.packages("pacman")

# install and load the three libraries needed, RODBC, xlsx and sendmailR
pacman::p_load(RODBC, xlsx,sendmailR) 

library(RODBC)
#establish odbc connection
connection <- odbcDriverConnect(connection="Driver={SQL Server}; server=YourServerNameHere;database=YourDataBaseHere;trusted_connection=yes;",readOnlyOptimize = TRUE)

#query database
JournalNines <- sqlQuery(connection, "
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
FROM	TEL_AX09_SP1_PROD.dbo.LEDGERTRANS
LEFT JOIN LEDGERTABLE
	ON LEDGERTABLE.ACCOUNTNUM = LEDGERTRANS.ACCOUNTNUM
LEFT JOIN USERINFO
	ON USERINFO.ID = LEDGERTRANS.CREATEDBY
WHERE
	CONVERT(varchar(30), LEDGERTRANS.AMOUNTCUR) LIKE '%999%' AND
	LEDGERTRANS.TRANSDATE > (GetDate()-31);")


Weekends <- sqlQuery(connection, "SELECT
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
	TRANSDATE DESC;")

KeyWords <- sqlQuery(connection, "SELECT
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
OR LEDGERTRANS.TXT LIKE '%corruption%' OR LEDGERTRANS.TXT LIKE '%plug%')")
#close database connection
close(connection)



#Convert from scientific notation:
options(scipen=999)


# gets the current month and year
monthYear<-format(Sys.Date(),"%Y-%m")

# combines current year and month and the desired file name to
fileName = paste(monthYear,"_MISTI_Analytics_Example.xlsx")

# type in the location you would like to output to be in
Path = "C:/Users/aclark/desktop/"

# combine, "concatenate" the fill name to the path
fileNameWithPath = paste(Path, fileName, collapse=)

# Check if the dataframe is empty. If it isn't, export to excel

if(dim(JournalNines)[1] != 0){
  write.xlsx(JournalNines, file=fileNameWithPath, sheetName="JournalNines")
}
if(dim(KeyWords)[1] != 0){
  write.xlsx(KeyWords, file=fileNameWithPath, sheetName="KeyWords", append=TRUE)
}
if(dim(Weekends)[1] != 0){
  write.xlsx(Weekends, file=fileNameWithPath, sheetName="Weekends", append=TRUE)
}


# https://github.com/rpremraj/mailR

#send.mail(from = "sender@gmail.com",
#          to = c("recipient1@gmail.com", "recipient2@gmail.com"),
#          subject = "Subject of the email",
#          body = "Body of the email",
#          smtp = list(host.name = "smtp.gmail.com", port = 465, user.name = "gmail_username", passwd = "password", ssl = TRUE),
#          authenticate = TRUE,
#          send = TRUE,
#          attach.files = email_path,
#          debug = TRUE)

