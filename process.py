import os
import shutil

import PyPDF4
from googleapiclient.http import MediaFileUpload
from googleapiclient.discovery import build

from authenticate import Authenticate
from schedule import Schedule
from data import Data


class Process(Schedule):
    """Processes the PDF files."""
    downloads = "/Users/bford/Downloads"
    uploads = "/Users/bford/Documents/uploads"

    def __init__(self, d):
        self.credentials = Authenticate.get_credentials()
        self.file_names = self.get_file_names()  # Files to process. 

        self.date = d.date
        self.semester = d.semester
        self.academic_year = d.academic_year
        self.students = d.data  # Dictionary indexed by periods. 

    def get_file_names(self):
        """Gets file names for all scanned PDFs in the downloads folder."""
        print('Reading downloads folder...')
        scanned_files = []
        files = os.listdir(self.downloads)
        for file in files:
            if file.startswith('scan') and file.endswith('pdf'):
                scanned_files.append(file)
                print('Found: {}'.format(file))
        return scanned_files

    def to_uploads(self):
        """Send PDFs to uploads folder.
        If PDF contains individual work, split the PDF first."""
        print('Processing files...')
        for file in self.file_names:
            file_info = self.file_info(file)
            if file_info[-1] == 'g':
                from_path = self.downloads + '/' + file
                to_path = self.uploads + '/' + self.name(file_info) + '.pdf'
                shutil.copy(from_path, to_path)
                if os.path.isfile(to_path):
                    print('To Uploads: {}'.format(to_path))
            else:
                self.split_pdf(file, file_info)

    def file_info(self, file_name):
        """Split filename string into assignment info.
        Return: list of info."""
        info = file_name.replace('.pdf', '').split('_')  # [scan, period, category, assignment, group/individual]
        info[1] = int(info[1])
        code = self.schedules[self.academic_year][info[1]]['code']
        info.insert(2, code)
        print(info)
        return info[1:]  # [period, course code, category, assignment, group/individual]

    def name(self, file_info):
        """Create file name from scanned file name."""
        code = file_info[1]
        category = file_info[2]
        assignment = file_info[3]
        new_filename = self.academic_year + '_' + code + '_' + category + '_' + assignment
        return new_filename

    def split_pdf(self, file_name, file_info):
        """Split PDF into separate files, one for each student."""
        period = file_info[0]
        assignment = file_info[2].upper() + ' ' + file_info[3].upper()

        df = self.students[period][self.students[period][assignment] != '0']
        df.reset_index(inplace=True)
        path = self.downloads + '/' + file_name
        pdf_file = open(path, 'rb')
        reader = PyPDF4.PdfFileReader(pdf_file)

        if not reader.isEncrypted:
            row = 0  # Row index in students DataFrame
            n = reader.numPages  # Length of PDF
            s = len(df)  # Number of students
            p = int(n/s)  # Number of pages per student
            for i in range(0, n, p):
                writer = PyPDF4.PdfFileWriter()
                for j in range(p):
                    page_object = reader.getPage(i + j)
                    writer.addPage(page_object)
                new_file_name = '{}_{}_{}.pdf'.format(self.name(file_info), df.loc[row]['First'], df.loc[row]['Last'])
                to_path = self.uploads + '/' + new_file_name
                output_file = open(to_path, 'wb')
                writer.write(output_file)
                output_file.close()
                print('To Uploads: {}'.format(to_path))
                row += 1
        pdf_file.close()

    def upload(self):
        print('Uploading to Google Drive...')
        files = os.listdir(self.uploads)
        service = build('drive', 'v3', credentials=self.credentials)
        for file in files:
            if file[-3:] == 'pdf':
                file_metadata = {'name': file, 'parents': ['1eRiWCLf4J-klD0vqHEYJ-w88GeSs-M7b']}
                file_path = self.uploads + '/' + file
                media = MediaFileUpload(file_path, mimetype='application/pdf')
                try:
                    response = service.files().create(body=file_metadata, media_body=media, fields='id').execute()
                except Exception as e:
                    print(e)
                else:
                    print('Uploaded: {} {}'.format(file_path, response))

    def empty_uploads(self):
        """Deletes all files in uploads"""
        # TODO caution: deletes all files regardless of upload success.
        print('Emptying uploads folder...')
        files_to_delete = os.listdir(self.uploads)
        for file in files_to_delete:
            file_path = self.uploads + '/' + file
            if os.path.isfile(file_path):
                os.remove(file_path)
                print('Deleted: {}'.format(file_path))
            else:
                print('No file: {}'.format(file_path))

    def empty_downloads(self):
        """Removes scanned files from Downloads."""
        print('Emptying downloads folder...')
        for file in self.file_names:
            file_path = self.downloads + '/' + file
            if os.path.isfile(file_path):
                os.remove(file_path)
                print('Deleted: {}'.format(file_path))
            else:
                print('No file: {}'.format(file_path))


if __name__ == "__main__":
    data = Data()
    process = Process(data)
    process.to_uploads()
    process.upload()
    process.empty_uploads()
    process.empty_downloads()
