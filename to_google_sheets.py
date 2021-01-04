#Author: Filippo Pisello
import gspread
from oauth2client.service_account import ServiceAccountCredentials

import sys
sys.path.append(r"C:\Users\Filippo Pisello\Desktop\Python\Git Projects\Git_Spreadsheet")
from spreadsheet import Spreadsheet

class GoogleSheet(Spreadsheet):
    """
    Class finalized to export a pandas data frame to a Google Sheet workbook.

    ---------------
    The class intakes as main argument a pandas dataframe. Given the json file
    with the authentication keys, it connects with a Google Sheet workbook and
    efficiently uploads to it the data frame content.

    It can be decided where to place the content within the target sheet through
    the
    """
    def __init__(self, dataframe, json_file_name, google_workbook_name,
                 index=False, starting_cell="A1", correct_lists=True, sheet=0):
        super().__init__(dataframe, index, starting_cell, correct_lists)
        self.json_file = json_file_name
        self.workbook_name = google_workbook_name
        self.sheet_id = sheet

        self.workbook = None
        self.sheet = None

    # ---------------------------------------------
    # Main function
    # ---------------------------------------------
    def to_google_sheet(self, fill_na_with=" ", clear_content=False, header=True):
        self._prepare_table(fill_na_with)

        self.workbook = self._get_authorization()
        self.sheet = self.get_sheet()

        if clear_content:
            self.sheet.clear()

        self.sheet.batch_update(self._batch_list(header))

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

    def _get_sheet(self):
        if isinstance(self.sheet_id, str):
            return self.workbook.worksheet(self.sheet_id)
        return self.workbook.get_worksheet(self.sheet_id)

    def _batch_list(self, keep_header):
        output = [{"range" : self.body.cells_range,
                   "values" : self.df.values.tolist()}]

        if keep_header:
            output.append({"range" : self.header.cells_range,
                           "values" : self._columns_for_batch()})
        if self.keep_index:
            output.append({"range" : self.index.cells_range,
                           "values" : self._index_for_batch()})
        return output

    def _columns_for_batch(self):
        if self.indexes_depth[1] > 1:
            output = []
            for level in range(self.indexes_depth[1]):
                output.append([i[level] for i in self.df.columns.values.tolist()])
            return output
        return [self.df.columns.values.tolist()]

    def _index_for_batch(self):
        if self.indexes_depth[0] > 1:
            return self.df.index.values.tolist()
        return [[x] for x in  self.df.index.values.tolist()]
