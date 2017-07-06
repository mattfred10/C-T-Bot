import email
import imaplib
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email import encoders
from email.mime.base import MIMEBase
from email.mime.text import MIMEText


class FetchMail:

    def __init__(self, mail_server, username, password, leaveunread, download_folder=".\\UnprocessedPOs\\"):
        self.connection = imaplib.IMAP4_SSL(mail_server)
        self.connection.login(username, password)
        self.connection.select(readonly=leaveunread)  # True = will leave messages unread False - mark as read
        self.path = download_folder

        # double checking here because debug can change self.path - easier to generalize in the future.
        if not os.path.exists(self.path):
            os.makedirs(self.path)

    def close_connection(self):
        """
        Close the connection to the IMAP server
        """
        self.connection.close()

    def save_attachment(self, msg):
        """
        Given a message, save its attachments to the 
        download folder specified at init (default is working directory/year/month/day)
        iterator can be used to ensure unique names in a session if emails have same attachment names (e.g., doc1.pdf)
        """
        skipped = []

        att_path = "No attachment found."
        for part in msg.walk():
            if part.get_content_maintype() == 'multipart':
                continue
            if part.get('Content-Disposition') is None:
                continue

            filename = part.get_filename()
            if filename is None:  # appears due to "outlook item files" i.e., tables in emails - safe to ignore?
                continue

            for file in filename.split('\n'):
                fnparts = filename.split('.')

                if 'Terms for Goods  Services' in fnparts[0] or 'Packing List' in fnparts[0] or 'invoice' in fnparts[0].lower() or 'Payment Advice Note' in fnparts[0] or 'scan' in fnparts[0].lower() or 'estes_' in fnparts[0].lower(): #don't want terms documents or invoices or scanned documents
                    skipped.append(fnparts[0] + '.' + fnparts[1].lower())
                elif fnparts[1].lower() == 'pdf' or 'xls' in fnparts[1].lower():  # Take pdf, xls, and xlsx
                    att_path = os.path.join(self.path + fnparts[0] + '.' + fnparts[1].lower())  # adding i such that file is unique !'Doc1.pdf'
                    i = 0
                    while os.path.isfile(att_path):  # don't overwrite
                        i += 1
                        att_path = os.path.join(self.path + fnparts[0] + '-' + str(i) + '.' + fnparts[1].lower())
                    fp = open(att_path, 'wb')
                    fp.write(part.get_payload(decode=True))
                    fp.close()
                else:
                    skipped.append(fnparts[0] + '.' + fnparts[1].lower())

        return skipped

    def fetch_unread_messages(self):
        """
        Retrieve unread messages
        """
        emails = []
        (result, messages) = self.connection.search(None, 'UnSeen')
        if result == "OK":
            print(messages)
            for message in messages[0].split():

                print(message)
                try:
                    ret, data = self.connection.fetch(message,'(RFC822)')
                except:
                    print("No new emails to read.")
                    self.close_connection()
                    exit()
                msg = email.message_from_bytes(data[0][1])
                if isinstance(msg, str) == False:
                    emails.append(msg)
                response, data = self.connection.store(message, '+FLAGS','\\Seen')
            return emails

    def parse_email_address(self, email_address):
        """
        Helper function to parse out the email address from the message

        return: tuple (name, address). Eg. ('John Doe', 'jdoe@example.com')
        """
        return email.utils.parseaddr(email_address)


class SendMail:

    def __init__(self, mail_server, port, username, password):
        self.mail_server = mail_server  # 'smtp.office365.com', 587)
        self.port = port
        self.sender = username
        self.password = password
        self.mail_server = mail_server
        self.composed = ''

    def close_connection(self):
        """
        Close the connection to the IMAP server
        """
        self.connection.quit()

    # make sure the name is unique
    def open_connection(self):
        self.connection = smtplib.SMTP(host=self.mail_server, port=self.port)
        self.connection.connect(self.mail_server, port=self.port)  # 'smtp.office365.com',587) #465
        self.connection.ehlo()
        self.connection.starttls()
        self.connection.ehlo()
        self.connection.login(self.sender, self.password)

    def composemsg(self, to, subject, body, attachmentpath=None):
        self.to = to
        msg = MIMEMultipart()
        msg["From"] = self.sender
        msg["To"] = to
        msg["Subject"] = subject
        msg.preamble = "I'm a bot. BEEP. BOOP."
        msg.attach(MIMEText(body, 'plain'))

        # Trying to parse filetypes according to the python documentation kept giving me errors.
        # Just send it as octet stream. Should work for any file type so we don't have to think about it.
        if attachmentpath:
            for log in attachmentpath:
                attachment = open(log, 'rb')
                filename = log.replace('./', '').split('\\')[-1]
                part = MIMEBase('application', "octet-stream")
                part.set_payload((attachment).read())
                encoders.encode_base64(part)
                part.add_header('Content-Disposition', 'attachment; filename= %s' % filename)
                msg.attach(part)
        self.composed = msg.as_string()

    def send(self):
        self.connection.sendmail(self.sender, self.to, self.composed)