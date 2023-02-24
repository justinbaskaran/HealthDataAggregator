
import email
import getpass, imaplib
import os
import sys
import zipfile
from datetime import date
import gspread
import re
strengthID=''
swimmingID=''
emailUserName=''
emailPassword=''

def parseFile(file_path):
#    print("hi")
    filePathSwimming=""
    filePathStrengthTraining=""
    with zipfile.ZipFile(file_path, 'r') as zip_ref:
        listOfFileNames = zip_ref.namelist()
        # Iterate over the file names
        for fileName in listOfFileNames:
           # Check filename endswith csv
      
            if re.match(".*com.samsung.shealth.exercise.(\d+).csv",fileName):
               # Extract a single file from zip
               zip_ref.extract(fileName)
               filePathSwimming=fileName
            if re.match('.*com.samsung.health.weight.(\d+).csv',fileName):
               # Extract a single file from zip
               zip_ref.extract(fileName)
               filePathStrengthTraining=fileName

    with open(filePathSwimming, 'r') as fin:
        data = fin.read().splitlines(True)
    with open(filePathSwimming, 'w') as fout:
        fout.writelines(data[1:])
    with open(filePathStrengthTraining, 'r') as fin:
        data = fin.read().splitlines(True)
    with open(filePathStrengthTraining, 'w') as fout:
        fout.writelines(data[1:])

    gc = gspread.service_account()
    swimmingContent = open(filePathSwimming, 'r').read()
    strengthContent = open(filePathStrengthTraining, 'r').read()
    gc.import_csv(swimmingID,swimmingContent)
    gc.import_csv(strengthID,strengthContent)
    

            
with open('passwords.txt', 'r') as fin:
    data = fin.read().splitlines(True)
for line in data:
    if line.split("=")[0] == "swimmingID":
        swimmingID=line.split("=")[1].strip()
    if line.split("=")[0] == "strengthID":
        strengthID=line.split("=")[1].strip()
    if line.split("=")[0] == "emailUsername":
        emailUserName=line.split("=")[1].strip()
    if line.split("=")[0] == "emailPassword":
        emailPassword=line.split("=")[1].strip()    


try:
    imapSession = imaplib.IMAP4_SSL('imap.gmail.com')
    typ, accountDetails = imapSession.login(emailUserName, emailPassword)
    if typ != 'OK':
        print("couldnt login")
    #    print 'Not able to sign in!'
        raise
    
    imapSession.select('"[Gmail]/Sent Mail"')

    dateTime = '(SINCE "{today}")'.format(today=date.today().strftime("%d-%b-%Y"))
   # dateTime = '(SINCE "22-JAN-2023")' 
    result, data = imapSession.uid('search', None, dateTime ) 
    #print ("hi")
    if typ != 'OK':
        print("not ok")
    #    print 'Error searching Inbox.'
        raise
 #   print("hi")    
    # Iterating over all emails
    if result == 'OK':
        for num in data[0].split():
            result, data = imapSession.uid('fetch', num, '(RFC822)')
            if result == 'OK':
                email_message = email.message_from_bytes(data[0][1])    # raw email text including headers
#                print('From:' + email_message['From'])
  #          print(email_message['To'])
            if email_message['To'] != None and 'justinbaskaran5+health@gmail.com' in email_message['To']:
#                print(email_message)
                for part in email_message.walk():
                    if part.get_content_maintype() == 'multipart':
                        continue
                    if part.get('Content-Disposition') is None:
                        continue

                    filename = part.get_filename()
                    att_path = os.path.join('./', filename)

                    if not os.path.isfile(att_path):
                        fp = open(att_path, 'wb')
                        fp.write(part.get_payload(decode=True))
                        fp.close()
#                    print("hi")
                    parseFile(att_path)
            

    imapSession.close()
    imapSession.logout()
except Exception as e:
    print(e)
    print("not able to downlaod all attachments")
 #   print 'Not able to download all attachments.'











