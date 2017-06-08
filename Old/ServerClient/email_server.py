# import yagmail
#
# yag = yagmail.SMTP('accounts1@cooperandturner-usa.com', 'Denver8o221!')
# yag.send('matthewfred@gmail.com', 'Test', 'Test')

# imap
# Server name: outlook.office365.com
# Port: 993
# Encryption method: TLS

import smtplib
from email.mime.multipart import MIMEMultipart
from email import encoders
from email.mime.base import MIMEBase
from email.mime.text import MIMEText

FROM = 'accounts1@cooperandturner-usa.com'
TO = 'comma3llc@gmail.com'
SUBJECT = 'Email Server'
BODY = 'This is the code for smtp sending of emails.'
#Do this better when we have proper file structure.
FILENAME = 'email_server.py'
PATH = './email_server.py'

msg = MIMEMultipart()
msg["From"] = FROM
msg["To"] = TO
msg["Subject"] = SUBJECT
msg.preamble = "help I cannot send an attachment to save my life"
msg.attach(MIMEText(BODY, 'plain'))

attachment = open(PATH, 'rb')

#Trying to parse any filetypes according to the python documentation kept giving me errors.
#Just send it as octet stream. Should work for any file type so we don't have to think about it.
part = MIMEBase('application', "octet-stream")
part.set_payload((attachment).read())
encoders.encode_base64(part)
part.add_header('Content-Disposition', 'attachment; filename= %s' % FILENAME)

msg.attach(part)

composed = msg.as_string()

server = smtplib.SMTP('smtp.office365.com', 587)
server.connect('smtp.office365.com',587) #465
server.ehlo()
server.starttls()
server.ehlo()

#Next, log in to the server
server.login('accounts1@cooperandturner-usa.com', 'Denver8o221!')

#Send the mail
server.sendmail('accounts1@cooperandturner-usa.com', 'comma3llc@gmail.com', composed)
server.quit()