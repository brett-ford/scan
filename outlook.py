import email
import imaplib
import os

from schedule import Schedule


class Outlook(Schedule):
    """Read Microsoft Outlook to download scanned files."""
    folder = "/Users/bford/Downloads"
    connection = None
    error = None

    def __init__(self):
        self.connection = imaplib.IMAP4_SSL(self.server, 993)
        self.connection.login(self.name, self.password)
        self.connection.select(mailbox='INBOX', readonly=False)
        print('Outlook Connection: {}'.format(self.connection))

    def read_inbox(self):
        """Retrieve unread messages from 'mbcopier@mosesbrown.org'."""
        emails = []
        status, results = self.connection.search(None, '(UNSEEN)', '(FROM "mbcopier@mosesbrown.org")')
        if status == "OK":
            print('Read Inbox: {}'.format(status))
            email_ids = results[0].split()
            if email_ids:
                for email_id in email_ids:
                    try:
                        envelope, data = self.connection.fetch(email_id, '(RFC822)')  # 'list' 2 'tuple', 'bytes'
                    except Exception as e:
                        print("Cannot fetch message: {}".format(email_id))
                        print("Error: {}".format(e))
                    else:
                        raw_email = data[0][1]  # 'bytes'
                        msg = email.message_from_bytes(raw_email)
                        self.save_attachment(msg)
                        self.archive(email_id)
            else:
                print("No new emails.")
                exit()
        else:
            self.error = "Failed to retrieve emails."
            print(self.error)
        return emails

    def save_attachment(self, msg):
        """Given a message, save its attachment to the downloads folder."""
        attachment_path = None
        for part in msg.walk():
            if part.get_content_maintype() == 'multipart':
                continue
            if part.get('Content-Disposition') is None:
                continue
            filename = part.get_filename()
            attachment_path = os.path.join(self.folder, filename)
            if not os.path.isfile(attachment_path):
                pdf = open(attachment_path, 'wb')
                pdf.write(part.get_payload(decode=True))
                pdf.close()
        print('Attachment: {}'.format(attachment_path))
        return attachment_path

    def archive(self, msg_uid):
        """Archive the message."""
        status, action = self.connection.copy(msg_uid, 'Archive')  # Copy to archive.
        print('Archive: {}, {}'.format(status, action))
        status, action = self.connection.store(msg_uid, '+FLAGS', '\\Deleted')  # Delete message in inbox.
        print('Delete: {}, {}'.format(status, action))

    def close_connection(self):
        """Close the connection to the IMAP server."""
        self.connection.close()


if __name__ == "__main__":
    mail = Outlook()
    mail.read_inbox()
    mail.close_connection()
