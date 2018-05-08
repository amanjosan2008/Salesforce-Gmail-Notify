#!/home/pi/berryconda3/bin/python3
# Debug.log file more Details
# Save Exceptions to Debug.log
# If internet connection comes up, print it
# Ignore My Comment added alert; not customer's

from simple_salesforce import Salesforce
from slackclient import SlackClient
import time, datetime
import imaplib, email, sys
import re, os, socket
import credential, traceback

# Enable Debug Logging to Debug.log file
def debug():
    return False

# Enable Flags assignment; This will change to IMAPLIB to Read/Write Mode
def flags_enabled():
    return True

# MAC Notifications
def mac():
    return False

# Logging Function
def en_log():
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
    if en_log():
        log('Error: Internet Connection down, Retrying after 60 seconds')
    time.sleep(60)
    return False

# SalesForce Case list dump
def sf():
    try:
        sf = Salesforce(username=credential.username, password=credential.password, security_token=credential.security_token)
    except:
        if en_log():
            log('Fatal: Invalid Credentials/Token; Exitting'+'\n')
        sys.exit()
    try:
        prt = sf.query_all("SELECT CaseNumber,Subject,IsClosed FROM Case WHERE OwnerId = '0050G00000AyfBz'")
    except simple_salesforce.exceptions.SalesforceGeneralError:
        if en_log():
            log('Salesforce Error, retrying in 60 seconds.'+'\n')
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
        if en_log():
            log('Imaplib.IMAP4.abort Error: Retrying in 60 seconds'+'\n')
        time.sleep(60)
        mailbox()
        if en_log():
            log('Re-connected to the Mail Server!'+'\n')
        #e = False
        return
    except TimeoutError:
        if en_log():
            log('TimeoutError: Retrying in 60 seconds'+'\n')
        time.sleep(60)
        #e = False
        return
    except OSError:
        if en_log():
            log('OSError: Retrying in 60 seconds'+'\n')
        time.sleep(60)
        #e = False
        return
    except BrokenPipeError:
        if en_log():
            log('BrokenPipeError: Retrying in 60 seconds'+'\n')
        time.sleep(60)
        #e = False
        return
    except:
        if en_log():
            log(traceback.print_exc())
        sys.exit()

# Current Mail Fetch Details
def write_b():
    if b > a:
        dir_path = os.path.dirname(os.path.realpath(__file__))
        curr = os.path.join(dir_path, "curr.ini")
        f = open(curr,'w')
        f.write(str(b))
        if debug():
            debug_log('Write b='+str(b)+'\n')         # Debugging
        f.close()
    else:
        if en_log():
            log('Error: Rollback event occured, ignored'+'\n')

# Fetch Last Seen MessID
def read_a():
    dir_path = os.path.dirname(os.path.realpath(__file__))
    curr = os.path.join(dir_path, "curr.ini")
    f = open(curr,'r')
    l = f.readlines()
    global a
    a = int(l[0])
    f.close()
    if en_log():
        log('Last processed MessID: '+str(a)+'\n')

# Fetch new mails function
def fetchmail():
    #for i in range(b, a, -1):
    for i in range(a+1,b+1):
        if debug():
            debug_log('Mail Function: a='+str(a)+' b='+str(b)+' i='+str(i)+'\n')             # Debugging
        typ, data = mail.fetch(str(i), '(RFC822)' )

        for response_part in data:
            if isinstance(response_part, tuple):
                msg = email.message_from_bytes(response_part[1])
                email_subject = msg['subject']
                email_from = msg['from']
                email_to = msg['to']
                if debug():
                    debug_log("Subject: "+email_subject+'\n'+"To: "+email_to+'\n')
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
                            if en_log():
                                log('Exception: '+str(u)+'\n')
                                log('part.get_content_type = '+part.get_content_type()+'\n')
                            string = "0"
                    else:
                        string = "0"
                if debug():
                    debug_log("Content type: "+str(part.get_content_type())+'\n')
                    debug_log("String: "+str(string.encode('utf-8').strip())+'\n')
                try:
                    c = str(RAW_SUB[0].strip())
                    #try:
                    if c in CASES_LIST:
                        if en_log():
                            log('\n'+local_date+'['+str(i)+']Case:'+c+'\n'+'From   : '+email_from+'\n'+'Subject: '+email_subject+'\n')
                        #slack(local_date+'['+str(i)+']Case:'+c+'\n'+'From   : '+email_from+'\n'+'Subject: '+email_subject, ':robot_face:')
                        slack(local_date+' Case:'+c+'\n'+'From   : '+email_from+'\n'+'Subject: '+email_subject, ':robot_face:')
                        if mac():
                            notify('Case Update', 'Case# '+c+': '+email_subject[0:30], 'Case Update')
                        if flags_enabled():
                            flags(i,"MyCase")
                    elif "New Case Assigned:" in email_subject:
                        if re.search(EMAIL_ID, email_to):
                            sf()
                            if en_log():
                                log('New Case in my Queue: '+c+' Refreshing SF data, Total Cases: '+str(len(CASES_LIST))+'\n')
                    elif 'New Case:' in string:
                        if en_log():
                            log(local_date+'['+str(i)+']New Case in QUEUE - Case: ' + c+'\n')
                        slack(local_date+' New Case in QUEUE - Case: ' + c, ':robot_face:')
                        if mac():
                            notify('New Case', 'New Case in Queue: '+ c, 'New Case')
                        if flags_enabled():
                            flags(i,"")
                    else:
                        if en_log():
                            log(local_date+ '['+str(i)+']Case:' + c+'\n')
                        if flags_enabled():
                            flags(i,"")
                except TypeError:
                #except:
                    #traceback.print_exc()
                    if re.search('\[JIRA\]\s\(AV-\d{5}', email_subject):
                          #if re.search('(a|A)man@avinetworks.com', email_to):
                          if re.search(EMAIL_ID, email_to):
                              d = re.compile(r'AV-\d{5}')
                              if en_log():
                                  log('\n'+local_date+'['+str(i)+']Jira Update: '+d.findall(email_subject)[0]+'\n')
                              slack(local_date+' Jira Update: '+d.findall(email_subject)[0], ':robot_face:')
                              if mac():
                                  notify('Jira Update', 'Jira# '+d.findall(email_subject)[0]+':'+email_subject[0:30], 'Jira Update')
                              if flags_enabled():
                                  flags(i,"MyJIRA")
                    else:
                        if en_log():
                            log(local_date+'['+str(i)+']No Case ID'+'\n')
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
        if en_log():
            log('ConnectionResetError: Retrying in 60 seconds'+'\n')
        time.sleep(60)
        return

# MacOS X Notify Function
def notify(title, text, say):
    os.system("""osascript -e 'display notification "{}" with title "{}"'; osascript -e 'say "{}"' """.format(text, title, say))

# Slack Alerts
def slack(message, icon):
    #token = credential.token
    sc = SlackClient(credential.token)
    sc.api_call('chat.postMessage', channel=credential.channel, text=message, username='My PyBot', icon_emoji=icon)

# Debug log file
def log(text):
    dir_path = os.path.dirname(os.path.realpath(__file__))
    log_file = os.path.join(dir_path, "log_file.log")
    f = open(log_file,'a')
    f.write(str(date())+': '+str(text))
    f.close()

# Debug log file
def debug_log(text):
    dir_path = os.path.dirname(os.path.realpath(__file__))
    debug_file = os.path.join(dir_path, "debug.log")
    f = open(debug_file,'a')
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
        if en_log():
            log('Unable to connect! Check Credentials'+'\n')
        sys.exit()

# Connect to SF:
while True:
    if is_connected():
        sf()
        if en_log():
            log('Fetched results from SalesForce\n'+'Total Cases: '+str(len(CASES_LIST))+' => ' +str(CASES_LIST).strip('[]')+'\n')
        mailbox()
        if en_log():
            log('Connected to the Mail Server!'+'\n')
        break
    else:
        #print('Error: Internet Connection down, Retrying after 60 seconds')
        is_connected()

read_a()
d = 0

# Main Loop
while True:
    d += 1
    #print(d)
    if d == 60:
        sf()
        if en_log():
            log('SF info retreived: Total Cases: '+str(len(CASES_LIST))+'\n')
        d = 0
    try:
        if is_connected():
            id()
            if a != b:
                if debug():
                    debug_log('Loop Values: a='+str(a)+' b='+str(b)+'\n')     # Debugging
                fetchmail()
                write_b()
                a = b
            time.sleep(30)
        else:
            if en_log():
                log('Error: Internet Connection down, Retrying after 60 seconds'+'\n')
            continue
            time.sleep(60)
    except KeyboardInterrupt:
        mail.logout()
        if en_log():
            log('Process Killed by Keyboard'+'\n')
        sys.exit()
    except:
        traceback.print_exc()
        sys.exit()
