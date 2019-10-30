import pickle
import os.path

from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request


class Authenticate:
    """Authenticate to the Google Sheets and Google Drive APIs.
    This code is adapted from Google's quick start guide."""

    @staticmethod
    def get_credentials():
        """Returns credentials for accessing a Google account"""
        scopes = ['https://www.googleapis.com/auth/drive.file']

        credentials = None
        if os.path.exists('token.pickle'):
            with open('token.pickle', 'rb') as token:
                credentials = pickle.load(token)
        if not credentials or not credentials.valid:
            if credentials and credentials.expired and credentials.refresh_token:
                credentials.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file('credentials.json', scopes)
                credentials = flow.run_local_server(port=0)
            with open('token.pickle', 'wb') as token:
                pickle.dump(credentials, token)
        return credentials
