from constants import *
import imaplib
import getpass
import imaplib
import base64
import os
import email

class GmailParser():
    def __init__(self):
        f = open(PASSWORD_FILE)
        self.password = f.read()
        self.username = "juliotestemail00@gmail.com"

    def parse_emails(self):
        try:
            imapSession = imaplib.IMAP4_SSL("imap.gmail.com")
            typ, accountDetails = imapSession.login(self.username, self.password)
            if typ != "OK":
                print("Not able to sign in!")
                raise

            imapSession.select()
            typ, data = imapSession.search(None, "UNSEEN")
            if typ != "OK":
                print("Error searching Inbox.")
                raise

            unseen_mails = data[0].split()
            print(f"New (unseen) emails discovered for {self.username}: [ {len(unseen_mails)} ] ")

            for msgId in data[0].split():
                typ, data = imapSession.fetch(msgId, "(RFC822)")
                raw_email = data[0][1]

                # converts byte literal to string removing b''
                raw_email_string = raw_email.decode("utf-8")
                email_message = email.message_from_string(raw_email_string)

                # downloading attachments
                for part in email_message.walk():
                    if part.get_content_maintype() == "multipart":
                        continue
                    if part.get("Content-Disposition") is None:
                        continue
                    fileName = part.get_filename()

                    if bool(fileName):
                        if (fileName.endswith('.docx')):
                            filePath = os.path.join(DOCS_FOLDER, fileName)
                        else:
                            filePath = os.path.join(ATTACHMENTS_FOLDER, fileName)

                        if not os.path.isfile(filePath):
                            fp = open(filePath, "wb")
                            fp.write(part.get_payload(decode=True))
                            fp.close()

            print('Done')

        except:
            print("Error")

        finally:
            try:
                imapSession.close()
                print(f"Session clossed for {self.username}")
            except:
                print(f"Error closing session for {self.username}")

            print("Logging out...")
            imapSession.logout()


if __name__ == "__main__":
    gm_parser = GmailParser()
    gm_parser.parse_emails()