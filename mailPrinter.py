import smtplib
import time
import imaplib
import email
import os

ORG_EMAIL   = "@gmail.com"
FROM_EMAIL  = "impresorarodrigodp05" + ORG_EMAIL
FROM_PWD    = ""
SMTP_SERVER = "imap.gmail.com"
SMTP_PORT   = 993

filterMail = 'rodrigodp05@gmail.com'
filterSubject = 'imprimir'
attachDir = 'printed/'
command = 'lpr -P HLL2310D'
optionFlag = '-o'
twoSided = 'sides=two-sided-long-edge'
oneSided = 'sides=one-sided'


def from_me(msg):
    # Mail has to be mine and subject has to be filterSubject
    if filterMail in msg['from'] and filterSubject in msg['subject']:
        return True
    return False

def print_attachments(msg):
    # Establish the printing options
    if len(msg['subject'].split())>=2:
        # Two ore more words in the subject => one sided
        options = optionFlag+' '+oneSided
    else:
        # Default => two sided
        options = optionFlag+' '+twoSided

    # Iterate on potential attachments
    for attach in msg.get_payload():
        # We filter body,... from attachments
        name = attach.get_filename()
        disp = attach.get_content_disposition()
        if disp and 'attachment' in disp and name:
            # Save attachment
            fp = open(attachDir+name, 'wb')
            fp.write(attach.get_payload(decode=True))
            fp.close()
            # Print file
            os.system(command+' '+options+' '+attachDir+name)
            
def check_email():
    try:
        # Connect
        mail = imaplib.IMAP4_SSL(SMTP_SERVER)
        mail.login(FROM_EMAIL,FROM_PWD)
        # Look for unread mails
        mail.select('inbox')
        ret, data = mail.search(None, '(UNSEEN)')
        if ret != 'OK':
            return
        # Iterate over unread mails
        id_list = data[0].split()
        for i in id_list:
            ret, data = mail.fetch(i, '(RFC822)')
            if ret=='OK':
                data = email.message_from_bytes(data[0][1])
                if from_me(data):
                    print_attachments(data)

    except Exception as e:
        print(str(e))
    
    mail.logout()
    return

def main():
    while(True):
        check_email()
        time.sleep(5)


main()