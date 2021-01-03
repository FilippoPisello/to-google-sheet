#Author: Filippo Pisello
import gspread
from oauth2client.service_account import ServiceAccountCredentials

import sys
sys.path.append(r"C:\Users\Filippo Pisello\Desktop\Python\Git Projects\Git_Spreadsheet")
from spreadsheet import Spreadsheet

class GoogleSheet(Spreadsheet):
    def __init__(self, dataframe, json_file_name, google_workbook_name,
                 index=False, skip_rows=0, skip_columns=0,
                 correct_lists=True, sheet_number=0):
        super().__init__(dataframe, index, skip_rows, skip_columns, correct_lists)
        self.json_file = json_file_name
        self.workbook_name = google_workbook_name
        self.sheet_number = sheet_number

        self.workbook = None
        self.sheet = None

    # ---------------------------------------------
    # Main function
    # ---------------------------------------------
    def to_google_sheet(self, fill_na_with=" ", clear_content=False, header=True):
        self._prepare_table(fill_na_with)

        self.workbook = self._get_authorization()
        self.sheet = self.workbook.get_worksheet(self.sheet_number)

        if clear_content:
            self.sheet.clear()

        if header:
            self._write_cells(self.header, self.df.columns)
        if self.keep_index:
            self._write_cells(self.index, self.df.index)
        self._write_cells(self.body, self.df)

    # ---------------------------------------------
    # Sub functions
    # ---------------------------------------------
    def _prepare_table(self, fill_na_with):
        # Convert datetime columns into string
        for column in self.df.columns:
            if self.df[column].dtype in ["datetime64[ns]", "datetime64"]:
                self.df[column] = self.df[column].astype(str)
        # Replace missing values with something else
        self.df.fillna(fill_na_with, inplace=True)

    def _get_authorization(self):
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

    def _write_cells(self, spreadsheet_element, table_portion):
        values_list = self._flatten_list(table_portion.values.tolist())
        cells = self.sheet.range(spreadsheet_element.cells_range)

        if len(values_list) != len(cells):
            raise IndexError("Len of cells range and values list do not match")

        for cell, value_ in zip(cells, values_list):
            cell.value = value_

        self.sheet.update_cells(cells)
        return

    @staticmethod
    def _flatten_list(values_list):
        """
        Given iterable, returns a list. If item in iterable is tuple or list, the
        sub elements are added to output, else item is output.
        """
        output = []
        for element in values_list:
            if isinstance(element, (list, tuple, set)):
                output.extend(element)
            else:
                output.append(element)
        return output