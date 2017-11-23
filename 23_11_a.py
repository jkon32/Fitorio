import sys
import imaplib
import email
import datetime
import smtplib
import pyodbc

def find_in_mtrl_non_grafted(hybrid,tray): # Find the active item and return list with CODE,MTRL,NAME
    cnxn = pyodbc.connect('DRIVER={SQL Server};SERVER=10.0.0.7\SQLEXPRESS;DATABASE=softone;UID=sa;PWD=')
    cursor = cnxn.cursor()
    cursor.execute("""SELECT A.MTRL,A.CODE,A.NAME FROM MTRL A WHERE A.COMPANY=1 AND A.SODTYPE=51
    AND A.CODE>='70.00.00000' AND A.CODE<='70.10.99999' AND A.CODE LIKE '70%' AND A.ISACTIVE=1
    AND A.NAME LIKE '%' + ? + '%' + ? + '%' ORDER BY A.CODE,A.MTRL""",hybrid,tray)
    confirmed_item = list(cursor.fetchall())
    return confirmed_item
def find_in_mtrl_grafted(hybrid_name,grafted,tray): # Find the active item and return list with CODE,MTRL,NAME
    cnxn = pyodbc.connect('DRIVER={SQL Server};SERVER=10.0.0.7\SQLEXPRESS;DATABASE=softone;UID=sa;PWD=')
    cursor = cnxn.cursor()
    cursor.execute("""SELECT A.MTRL,A.CODE,A.NAME FROM MTRL A WHERE A.COMPANY=1 AND A.SODTYPE=51
    AND A.CODE>='70.10.99999' AND A.CODE<='70.99.99999' AND A.CODE LIKE '70%' AND A.ISACTIVE=1
    AND A.NAME LIKE '%' + ? + '%' + ? + '%' + ? + '%' ORDER BY A.CODE,A.MTRL""",hybrid_name,grafted,tray)
    confirmed_item = list(cursor.fetchall())
    return confirmed_item

error_count = 0 #Λάθη στην αναγνώριση

M = imaplib.IMAP4_SSL('imap.gmail.com') #M will refer to the mail server
gmail_user = 'fitorio.geoponiki@gmail.com'
gmail_password = 'Th7313qK'

try: #Login to gmail server
    M.login(gmail_user, gmail_password)
except imaplib.IMAP4.error:
    print("LOGIN FAILED!!! ")
    # ... exit or deal with failure...
rv, mailboxes = M.list() #Get mailbox status
if rv == 'OK':
    print("Mailboxes:",mailboxes)
    print (M.status('INBOX','(MESSAGES RECENT UIDNEXT UIDVALIDITY UNSEEN)'))

M.select("INBOX")#Inbox will only show from now on
result, data = M.search(None, "ALL")
ids = data[0]  # data is a list.
id_list = ids.split()  # ids is a space separated string
oldest_email_id = id_list[0]  # get the oldest email
result, data = M.fetch(oldest_email_id, "(RFC822)")  # fetch the email body (RFC822) for the given ID
raw_email = data[0][1]  # raw text of the whole email incl headers and alternate payloads
utf_email=raw_email.decode('utf-8')
email_message = email.message_from_string(utf_email)

sender_address = email.utils.parseaddr(email_message['From'])
print(sender_address)


b= email_message
confirmed_lines = ''
empty_lines = 0
rejected_lines = ''
if b.is_multipart():
    for payload in b.get_payload():
        if payload.get_content_type() == "text/plain" :
            body = payload.get_payload(decode=True).decode(payload.get_content_charset())
            for line in str(body).splitlines():
              if line.startswith(tuple('-1-2-3-4-5-6-7-8-90123456789')) and len(line.split(" ")) == 3 and line.endswith(('600','504','330','228','150','77','54','24')):#Check line structure
                 print ('Αναγνωρίστηκε:',line)
                 number_of_trays,hybrid,tray =line.split(" ")# Split the three elements of the line
                 line_output = ''
                 if len(find_in_mtrl_non_grafted(hybrid,tray)) == 1:# Check records found in MTRL
                     print("Confirmed: ", number_of_trays, "δίσκοι ", find_in_mtrl_non_grafted(hybrid, tray))  # Get hybrid data from MTRL
                     line_output = ("Confirmed: ", str(number_of_trays), "δίσκοι ", str(find_in_mtrl_non_grafted(hybrid, tray)))
                 elif len(find_in_mtrl_non_grafted(hybrid,tray)) == 0:
                     print("Δεν βρέθηκε το είδος:",line)
                     line_output = ("Δεν βρέθηκε το είδος:", line)
                     error_count += 1
                 elif len(find_in_mtrl_non_grafted(hybrid,tray)) > 1:
                     print("ΠΡΟΣΟΧΗ: ΠΟΛΛΑΠΛΕΣ ΕΓΓΡΑΦΕΣ:",find_in_mtrl_non_grafted(hybrid, tray))
                     line_output = ("ΠΡΟΣΟΧΗ: ΠΟΛΛΑΠΛΕΣ ΕΓΓΡΑΦΕΣ:",find_in_mtrl_non_grafted(hybrid, tray))
                     error_count += 1
                 confirmed_lines += str(line_output)
                 confirmed_lines += '\n'
              elif not line:
                  print('Κενή γραμμή')
                  empty_lines +=1
              else:
                print('Αγνοήθηκε:',line)
                rejected_lines += line
                rejected_lines += '\n'

else:
    print('Not multipart')
    print (b.get_payload())

print ('Καταχωρήθηκαν τα κάτωθι:','\n',confirmed_lines)
print ('Αγνοήθηκαν τα κάτωθι:','\n',rejected_lines)
print ('Κενές σειρές:',empty_lines)
print ("Email id:",oldest_email_id)
print("Errors:",error_count)

sent_from = gmail_user
to = [sender_address[1]]
subject = 'OMG Super Important Message'
body = (confirmed_lines, '\n', 'Αγνοήθηκαν τα κάτωθι:',rejected_lines, '\n','Κενές σειρές:',empty_lines)

email_text = """\
From: %s
To: %s
Subject: %s

%s
""" % (sent_from, ", ".join(to), subject, body)

try:
    server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
    server.ehlo()
    server.login(gmail_user, gmail_password)
    server.sendmail(sent_from, to, email_text.encode('utf-8'))
    server.close()

    print ('Email sent!')
    error_count = 0
except:
    print ('Something went wrong...')