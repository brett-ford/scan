# Scan

Before returning papers to students, I scan them in our copier for storage. The copier creates one PDF file for the 
assignment and sends it to my email inbox.

This project automates the rest of the process by completing the following steps:

1. Logs into my Outlook account, finds new emails from the copier and downloads the PDF attachment.
2. Reads the downloads folder and finds all PDFs from the scanner.
3. Reads my gradebooks to create a list of students who have completed the assignment.
4. Splits each PDF into separate files, one for each student, and saves the files in an uploads folder.
5. Sends the contents of the uploads folder to a folder in my Google Drive account where I store student work.
6. Deletes the contents of the uploads folder.
7. Deletes the PDFs from the downloads folder.

Dependencies:
* PyPDF4