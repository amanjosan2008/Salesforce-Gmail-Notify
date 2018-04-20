#!/usr/bin/env python3
# Add Subject & Case# to Notification

# https://<yoursalesforcehostname>/_ui/system/security/ResetApiTokenEdit?retURL=%2Fui%2Fsetup%2FSetup%3Fsetupid%3DPersonalInfo&setupid=ResetApiToken
# https://avi.my.salesforce.com/_ui/system/security/ResetApiTokenEdit?retURL=%2Fui%2Fsetup%2FSetup%3Fsetupid%3DPersonalInfo&setupid=ResetApiToken
from simple_salesforce import Salesforce
import time, datetime
import imaplib
import email, sys
import re, os
import socket
import credential

# Function to print Date
def date():
    Time = time.strftime("%I:%M %p", time.localtime())
    return Time

def is_connected():
  try:
    host = socket.gethostbyname("salesforce.com")
    return True
  except:
    return False

# SalesForce Case list dump
def sf():
    try:
        sf = Salesforce(username=credential.username, password=credential.password, security_token=credential.security_token)
    except:
        print("Fatal: Invalid Credentials/Token; Exitting")
        sys.exit()
    try:
        prt = sf.query_all("SELECT CaseNumber,Subject,IsClosed FROM Case WHERE OwnerId = '0050G00000AyfBz'")
        print(date()+ ': Fetched results from SalesForce')
    except simple_salesforce.exceptions.SalesforceGeneralError:
        print('Salesforce Error, retrying in 60 seconds.')
        time.sleep(60)
        pass
    global CASES_LIST
    CASES_LIST = []
    for i in range(len(prt['records'])):
        CASES_LIST.append(str(prt['records'][i]['CaseNumber']))
    print('Total Cases: '+str(len(CASES_LIST))+' => ' +str(CASES_LIST).strip('[]'))

# Function to get & save latest MessID
def id():
    mail.select('inbox', readonly=True)
    type, data = mail.search(None, '(UNSEEN)')
    mail_ids = data[0]
    id_list = mail_ids.split()
    global b
    b = int(id_list[-1])
    f = open('curr.ini','w')
    f.write(str(b))
    f.close()

# Fetch new mails function
def mymail():
    for i in range(b, a, -1):
        typ, data = mail.fetch(str(i), '(RFC822)' )

        for response_part in data:
            if isinstance(response_part, tuple):
                msg = email.message_from_bytes(response_part[1])
                email_subject = msg['subject']
                email_from = msg['from']
                email_to = msg['to']
                #print(msg.get_payload(decode=True))
                date_tuple = email.utils.parsedate_tz(msg['Date'])
                if date_tuple:
                    localdate = datetime.datetime.fromtimestamp(email.utils.mktime_tz(date_tuple))
                    local_date = str(localdate.strftime("%d %b %H:%M"))
                RAW_SUB = re.search(r'\s\d{4}', email_subject)
                for part in msg.walk():
                    if part.get_content_type() == "text/html":
                        body = part.get_payload(decode=True)
                        try:
                            string = body.decode('utf-8')[0:200]
                        except UnicodeDecodeError as u:
                            string = body.decode('ISO-8859-1')[0:200]
                        except:
                            print("Exception: " +str(u))
                            print(part.get_content_type())
                            string = "0"
                try:
                    c = str(RAW_SUB[0].strip())
                    #try:
                    if c in CASES_LIST:
                        print(local_date+ ': [' + str(i)+ '] Case: ' + c)
                        print('From: ' + email_from)
                        print('Subject: ' + email_subject + '\n')
                        notify("Update", "Case Update", "blow")
                    #elif 'New Case:' in msg:
                    elif 'New Case:' in string:
                        print(local_date+ ': ['+ str(i)+'] New Case in QUEUE - Case: ' + c)
                        notify("New Case", "New Case in Queue", "purr")
                    #elif re.search('\[JIRA\]\s\(AV-\d{5}', email_subject) & re.search('(a|A)man@avinetworks.com', email_to):
                    #    print('Jira Update: '+re.search(r'AV-\d{5}', email_subject))
                    else:
                        print(local_date+ ': [' + str(i)+ '] Case: ' + c+'\n')
                except TypeError:
                    if re.search('\[JIRA\]\s\(AV-\d{5}', email_subject):
                          if re.search('(a|A)man@avinetworks.com', email_to):
                              d = re.compile(r'AV-\d{5}')
                              print(local_date+' ['+str(i)+']: Jira Update: '+d.findall(email_subject)[0]+'\n')
                              notify("Jira", "Jira Update", "Sosumi")
                    else:
                        print(local_date+ ': [' + str(i)+'] No Case ID\n')
                        pass
                #except IndexError:

# MacOS X Notify Function
def notify(title, text, sound):
    os.system("""
              osascript -e 'display notification "{}" with title "{}" sound name "{}"'
              """.format(text, title, sound))

# Connect to SF:
if is_connected():
    sf()
else:
    print("Error: Internet Connection down, Retrying after 60 seconds.")
    time.sleep(60)
    is_connected()

# Email Credentails:
ORG_EMAIL   = credential.ORG_EMAIL
FROM_EMAIL  = credential.FROM_EMAIL
FROM_PWD    = credential.FROM_PWD
SMTP_SERVER = "imap.gmail.com"

mail = imaplib.IMAP4_SSL(SMTP_SERVER)

try:
    mail.login(FROM_EMAIL,FROM_PWD)
    print(date()+ ': Connected to the Mail Server!')
except:
    print('Unable to connect! Check Credentials.')
    sys.exit()

# Fetch Last Seen MessID
f = open('curr.ini','r')
l = f.readlines()
a = int(l[0])
f.close()
print('Last processed MessID: '+str(a)+'\n')

# Main Loop
while True:
    try:
        id()
        if a != b:
            mymail()
            a = b
        time.sleep(30)
#    except TimeoutError:
#        print('Internet Down')
#        time.sleep(60)
#    except Exception as e:
#        print("Exception: "+str(e))
#        time.sleep(60)
#        continue
    except KeyboardInterrupt:
        mail.logout()
        print("Process Killed by Keyboard")
        sys.exit()
    except imaplib.IMAP4.abort:
        print("Socket Error: Internet might be down, Retrying in 60 seconds")
        time.sleep(60)
        continue
    except TimeoutError:
        print("TimeoutError: Internet might be down, Retrying in 60 seconds")
        time.sleep(60)
        continue
    except Exception as e:
        print("Exception: "+str(e))
        time.sleep(60)
        continue
