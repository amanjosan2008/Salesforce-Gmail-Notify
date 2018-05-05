#!/usr/bin/env python3
# Debug.log file more Details
# Save Exceptions to Debug.log
# Stuck in Loop error = > Monitor
# If internet connection comes up, print it
# Ignore My Comment added alert; not customer's

from simple_salesforce import Salesforce
import time, datetime
import imaplib
import email, sys
import re, os
import socket
import credential
import traceback

# Enable Debug Logging
def debug():
    return False

# Enable Flags assignment; This will change to IMAPLIB to Read/Write Mode
def flags_enabled():
    return True

# Function to print date
def date():
    Time = time.strftime("%I:%M %p", time.localtime())
    return Time

# Check Internet connectivity
def is_connected():
  try:
    host = socket.gethostbyname("www.google.com")
    s = socket.create_connection((host, 80), 2)
    return True
  except:
    print(date()+ ': Error: Internet Connection down, Retrying after 60 seconds')
    time.sleep(60)
    return False

# SalesForce Case list dump
def sf():
    try:
        sf = Salesforce(username=credential.username, password=credential.password, security_token=credential.security_token)
    except:
        print(date()+ ': Fatal: Invalid Credentials/Token; Exitting')
        sys.exit()
    try:
        prt = sf.query_all("SELECT CaseNumber,Subject,IsClosed FROM Case WHERE OwnerId = '0050G00000AyfBz'")
    except simple_salesforce.exceptions.SalesforceGeneralError:
        print(date()+ ': Salesforce Error, retrying in 60 seconds.')
        time.sleep(60)
        pass
    global CASES_LIST
    CASES_LIST = []
    for i in range(len(prt['records'])):
        CASES_LIST.append(str(prt['records'][i]['CaseNumber']))

# Function to get & save latest MessID
def id():
    #global e
    try:
        if flags_enabled():
            mail.select('inbox', readonly=False)
        else:
            mail.select('inbox', readonly=True)
        type, data = mail.search(None, '(ALL)')
        mail_ids = data[0]
        id_list = mail_ids.split()
        global b
        b = int(id_list[-1])
        #e = True
    except imaplib.IMAP4.abort:
        print(date()+ ': Imaplib.IMAP4.abort Error: Retrying in 60 seconds')
        time.sleep(60)
        mailbox()
        print(date()+ ': Re-connected to the Mail Server!')
        #e = False
        return
    except TimeoutError:
        print(date()+ ': TimeoutError: Retrying in 60 seconds')
        time.sleep(60)
        #e = False
        return
    except OSError:
        print(date()+ ': OSError: Retrying in 60 seconds')
        time.sleep(60)
        #e = False
        return
    except BrokenPipeError:
        print(date()+ ': BrokenPipeError: Retrying in 60 seconds')
        time.sleep(60)
        #e = False
        return
    except:
        traceback.print_exc()
        sys.exit()

# Current Mail Fetch Details
def write_b():
    if b > a:
        dir_path = os.path.dirname(os.path.realpath(__file__))
        curr = os.path.join(dir_path, "curr.ini")
        f = open(curr,'w')
        f.write(str(b))
        if debug():
            log('Write b='+str(b)+'\n')         # Debugging
        f.close()
    else:
        print(date()+ ': Error: Rollback event occured, ignored')

# Fetch Last Seen MessID
def read_a():
    dir_path = os.path.dirname(os.path.realpath(__file__))
    curr = os.path.join(dir_path, "curr.ini")
    f = open(curr,'r')
    l = f.readlines()
    global a
    a = int(l[0])
    f.close()
    print('Last processed MessID: '+str(a)+'\n')

# Fetch new mails function
def fetchmail():
    #for i in range(b, a, -1):
    for i in range(a+1,b+1):
        if debug():
            log('Mail Function: a='+str(a)+' b='+str(b)+' i='+str(i)+'\n')             # Debugging
        typ, data = mail.fetch(str(i), '(RFC822)' )

        for response_part in data:
            if isinstance(response_part, tuple):
                msg = email.message_from_bytes(response_part[1])
                email_subject = msg['subject']
                email_from = msg['from']
                email_to = msg['to']
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
                            print('Exception: '+str(u))
                            print('part.get_content_type = '+part.get_content_type()+'\n')
                            string = "0"
                    else:
                        string = "0"
                try:
                    c = str(RAW_SUB[0].strip())
                    #try:
                    if c in CASES_LIST:
                        print('\n'+local_date+'['+str(i)+']Case:'+c+'\n'+'From   : '+email_from+'\n'+'Subject: '+email_subject+'\n')
                        notify('Case Update', 'Case# '+c+': '+email_subject[0:30], 'Case Update')
                        if flags_enabled():
                            flags(i,"MyCase")
                    elif "New Case Assigned:" in email_subject:
                        if re.search(EMAIL_ID, email_to):
                            sf()
                            print(date()+ ': New Case in my Queue: '+c+' Refreshing SF data, Total Cases: '+str(len(CASES_LIST)))
                    elif 'New Case:' in string:
                        print(local_date+'['+str(i)+']New Case in QUEUE - Case: ' + c)
                        notify('New Case', 'New Case in Queue: '+ c, 'New Case')
                        if flags_enabled():
                            flags(i,"")
                    else:
                        print(local_date+ '['+str(i)+']Case:' + c)
                        if flags_enabled():
                            flags(i,"")
                except TypeError:
                    if re.search('\[JIRA\]\s\(AV-\d{5}', email_subject):
                          #if re.search('(a|A)man@avinetworks.com', email_to):
                          if re.search(EMAIL_ID, email_to):
                              d = re.compile(r'AV-\d{5}')
                              print('\n'+local_date+'['+str(i)+']Jira Update: '+d.findall(email_subject)[0]+'\n')
                              notify('Jira Update', 'Jira# '+d.findall(email_subject)[0]+':'+email_subject[0:30], 'Jira Update')
                              if flags_enabled():
                                  flags(i,"MyJIRA")
                    else:
                        print(local_date+'['+str(i)+']No Case ID')
                        if flags_enabled():
                            flags(i,"")
                        #pass

# Function to assign Flags
def flags(x,flag):
    if flag:
        mail.store(str(x), '+X-GM-LABELS', flag)
    else:
        pass
    try:
        mail.store(str(x), '-FLAGS', '(\Seen)')
        mail.expunge()
    except ConnectionResetError:
        print(date()+ ': ConnectionResetError: Retrying in 60 seconds')
        time.sleep(60)
        return

# MacOS X Notify Function
def notify(title, text, say):
    os.system("""osascript -e 'display notification "{}" with title "{}"'; osascript -e 'say "{}"' """.format(text, title, say))

# Debug log file
def log(text):
    f = open('debug.log','a')
    f.write(str(date())+': '+str(text))
    f.close()

# Email Credentails:
ORG_EMAIL   = credential.ORG_EMAIL
EMAIL_ID  = credential.FROM_EMAIL
FROM_PWD    = credential.FROM_PWD
SMTP_SERVER = "imap.gmail.com"

# Email Box connect function:
def mailbox():
    global mail
    mail = imaplib.IMAP4_SSL(SMTP_SERVER)
    try:
        mail.login(EMAIL_ID,FROM_PWD)
    except:
        print(date()+ ': Unable to connect! Check Credentials')
        sys.exit()

# Connect to SF:
while True:
    if is_connected():
        sf()
        print(date()+ ': Fetched results from SalesForce\n'+'Total Cases: '+str(len(CASES_LIST))+' => ' +str(CASES_LIST).strip('[]'))
        mailbox()
        print(date()+ ': Connected to the Mail Server!')
        break
    else:
        #print(date()+ ': Error: Internet Connection down, Retrying after 60 seconds')
        is_connected()

read_a()
d = 0

# Main Loop
while True:
    d += 1
    #print(d)
    if d == 60:
        sf()
        print(date()+ ': SF info retreived: Total Cases: '+str(len(CASES_LIST)))
        d = 0
    try:
        if is_connected():
            id()
            if a != b:
                if debug():
                    log('Loop Values: a='+str(a)+' b='+str(b)+'\n')     # Debugging
                fetchmail()
                write_b()
                a = b
            time.sleep(30)
        else:
            print(date()+ ': Error: Internet Connection down, Retrying after 60 seconds')
            continue
            time.sleep(60)
    except KeyboardInterrupt:
        mail.logout()
        print("Process Killed by Keyboard")
        sys.exit()
    except:
        traceback.print_exc()
        sys.exit()
