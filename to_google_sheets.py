#!/usr/bin/env python

import gspread
from oauth2client.service_account import ServiceAccountCredentials


class Google_sheet:
    def __init__(self, dataframe, json_file_name, google_workbook_name, sheet_number=0,
                 fillna_with=" ", overwrite_sheet=True, keep_header=False, heading_height=1):
        self.df = dataframe
        self.json_file = json_file_name
        self.workbook_name = google_workbook_name
        self.sheet_number = sheet_number
        self.fillna_with = fillna_with
        self.overwrite_sheet = overwrite_sheet
        self.keep_header = keep_header
        self.heading_height = heading_height

        self.sheet = ""
        self.sheet_instance = ""
        self.starting_row = ""

        self.to_google_sheet()

    # ---------------------------------------------
    # Main function
    # ---------------------------------------------
    def to_google_sheet(self):
        self.__prepare_table()
        self.sheet = self.__get_authorization()
        self.sheet_instance = self.sheet.get_worksheet(self.sheet_number)
        self.starting_row = self.__find_starting_row()
        self.sheet_instance.insert_rows(self.df.values.tolist(),
                                        row=self.starting_row)
        if not self.keep_header and self.overwrite_sheet:
            self.sheet_instance.insert_rows([self.df.columns.tolist()],
                                            row=self.starting_row)

    # ---------------------------------------------
    # Sub functions
    # ---------------------------------------------
    def __prepare_table(self):
        # Convert datetime columns into string
        for column in self.df.columns:
            if self.df[column].dtype in ["datetime64[ns]", "datetime64"]:
                self.df[column] = self.df[column].astype(str)
        # Replace missing values with something else
        self.df.fillna(self.fillna_with, inplace=True)

    def __get_authorization(self):
        # Define the scope
        scope = ['https://spreadsheets.google.com/feeds',
                 'https://www.googleapis.com/auth/drive']
        # Add credentials to the account
        creds = ServiceAccountCredentials.from_json_keyfile_name(self.json_file,
                                                                 scope)
        # Authorize the clientsheet
        client = gspread.authorize(creds)

        # Get the instance of the Spreadsheet
        return client.open(self.workbook_name)

    def __find_starting_row(self):
        if self.overwrite_sheet:
            header_coefficient = 1 + self.keep_header * self.heading_height
            self.sheet_instance.delete_rows(header_coefficient, self.sheet_instance.row_count - 1)
            return header_coefficient
        else:
            return self.sheet_instance.row_count
