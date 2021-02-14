from constants import *
import imaplib
import getpass
import imaplib
import base64
import os
import email


class GmailParser:
    def __init__(self):
        f = open(PASSWORD_FILE)
        self.password = f.read()
        self.username = "juliotestemail00@gmail.com"
        self.write_route = TMP_FOLDER

    def parse_emails(self):
        try:
            imap_session = self.imap_connect()
        except:
            print("Error connecting")

        try:
            unseen_mails = self.get_mails(imap_session)
            print(
                f"New (unseen) emails discovered for {self.username}: [{len(unseen_mails)}] "
            )
            if (len(unseen_mails) > 0):
                os.makedirs(self.write_route)
                for msg_id in unseen_mails:
                    email_message = self.get_email_message(imap_session, msg_id)
                    for part in email_message.walk():
                        if not self.is_attachment(part):
                            continue
                        self.write_attachment(part)
            print("Done")
        except:
            print("Error parsing emails")
        finally:
            self.close_session(imap_session)

    def imap_connect(self):
        imap_session = imaplib.IMAP4_SSL("imap.gmail.com")
        typ, accountDetails = imap_session.login(self.username, self.password)
        if typ != "OK":
            print("Not able to sign in!")
            raise
        return imap_session

    def get_mails(self, imap_session):
        imap_session.select()
        typ, data = imap_session.search(None, "UNSEEN")
        if typ != "OK":
            print("Error searching Inbox.")
            raise
        return data[0].split()

    def get_email_message(self, imap_session, message_id):
        typ, data = imap_session.fetch(message_id, "(RFC822)")
        raw_email = data[0][1]
        raw_email_string = raw_email.decode("utf-8")
        return email.message_from_string(raw_email_string)

    def is_attachment(self, part):
        is_multipart = part.get_content_maintype() == "multipart"
        has_content_disposition = part.get("Content-Disposition") is not None
        return has_content_disposition and not is_multipart

    def write_attachment(self, attachment):
        filename = attachment.get_filename()
        if bool(filename):
            if filename.endswith(".docx"):
                filePath = os.path.join(self.write_route, filename)
            else:
                filePath = os.path.join(ATTACHMENTS_FOLDER, filename)
            if not os.path.isfile(filePath):
                fp = open(filePath, "wb")
                fp.write(attachment.get_payload(decode=True))
                fp.close()

    def close_session(self, imap_session):
        try:
            imap_session.close()
            print(f"Session clossed for {self.username}")
        except:
            print(f"Error closing session for {self.username}")
        print("Logging out...")
        imap_session.logout()

