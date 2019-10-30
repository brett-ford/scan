from datetime import datetime as dt

import pandas as pd
from googleapiclient.discovery import build

from authenticate import Authenticate
from schedule import Schedule


class Data(Schedule):
    """Object that holds information from the grade books."""

    def __init__(self):
        self.date = dt.today()
        self.semester = self.get_semester()
        self.academic_year = self.academic_year()
        self.credentials = Authenticate.get_credentials()
        self.data = self.get_data()  # Dictionary indexed by class periods.
        self.print_results()

    def get_data(self):
        """Returns dictionary of periods and dataframes."""
        print('Reading grade books...')
        data = {}
        gradebook_ids = []
        ss_range = 'Semester {}!B2:BB40'.format(self.semester)  # Spreadsheet range

        service = build('sheets', 'v4', credentials=self.credentials)  # Call the Sheets API
        sheet = service.spreadsheets()

        for p in self.schedules[self.academic_year]:
            gb_id = self.schedules[self.academic_year][p]['gradebook_id']  # Spreadsheet ID
            if gb_id not in gradebook_ids:
                gb_data = []  # Array to hold the gradebook data.
                try:
                    result = sheet.values().get(spreadsheetId=gb_id, range=ss_range).execute()
                    values = result.get('values', [])
                except Exception as e:
                    print('Did not read: period {}'.format(p))
                    print(e)
                else:
                    if not values:
                        print('No data found: period {}'.format(p))
                        continue
                    else:
                        column_headers = values[0]
                        for row in values[1:-1]:
                            gb_data.append(row)
                    data_frame = pd.DataFrame(gb_data, columns=column_headers)
                    gradebook_ids.append(gb_id)
                    data[p] = data_frame
        return data  # Dictionary indexed by class periods.

    def get_semester(self):
        """Returns current semester."""
        s2_start = dt(2020, 1, 22)  # Semester 2 start date.
        if self.date >= s2_start:
            return 2
        return 1

    def academic_year(self):
        """Returns academic year as a string: 2019_2020"""
        if self.semester == 1:
            return str(self.date.year) + '_' + str(self.date.year + 1)
        return str(self.date.year - 1) + '_' + str(self.date.year)

    def print_results(self):
        """Prints summary of data found."""
        for p, df in self.data.items():
            print('Period {}:'.format(p))
            print(df)
            print()


if __name__ == '__main__':
    Data()
