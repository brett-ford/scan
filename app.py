#!/usr/bin/env python3

from outlook import Outlook
from data import Data
from process import Process


def main():
    """Get scans from the copier and send them to Google Drive."""
    mail = Outlook()  # Open connection to the Outlook server.
    mail.read_inbox()  # Download scans from the copier.
    mail.close_connection()  # Close connection to the Outlook server.
    student_data = Data()  # Read information from the gradebooks.
    process = Process(student_data)  # Create process object.
    process.to_uploads()  # Process each pdf and store results in uploads folder.
    process.upload()  # Upload each new PDF to the appropriate Google Drive folder.
    process.empty_uploads()  # Empty the local uploads folder
    process.empty_downloads()  # Empty the local downloads folder.


if __name__ == '__main__':
    main()
