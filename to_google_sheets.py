#Author: Filippo Pisello
from typing import List, Dict, Union

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
    uploads to it the data frame content.

    It can be decided where to place the content within the target sheet through
    the starting_cell argument.

    Arguments
    ----------------
    dataframe : pandas dataframe object (mandatory)
        Dataframe to be considered.
    json_file_name: str (mandatory)
        The .json file exported from the Google Sheet API which contains the
        authentication key.
    google_workbook_name: str (mandatory)
        The name of the Google Sheet workbook which should receive the data.
    sheet_id: str or int, default=0
        Argument to identify the target sheet within the workbook. If int, it is
        interpreted as sheet index, if str as sheet name.
    index: Bool, default=False
        If True, the index is exported together with the header and body.
    starting_cell: str, default="A1"
        The cell where it will be placed the top left corner of the dataframe.
    correct_lists: Bool, default=True
        If True, the lists stored as the dataframe entries are modified to be more
        readable in the traditional spreadsheet softwares. It helps with the
        Google Sheet compatibility. Type help(GoogleSheet.correct_lists_for_export)
        for further details.
    """
    def __init__(self, dataframe, json_file_name: str, google_workbook_name: str,
                 sheet_id=0, index=False, starting_cell="A1", correct_lists=True):
        super().__init__(dataframe, index, starting_cell, correct_lists)
        self.json_file = json_file_name
        self.workbook_name = google_workbook_name
        self.sheet_id = sheet_id

    # -------------------------------------------------------------------------
    # 1 - Main Elements
    # -------------------------------------------------------------------------
    # 1.1 - Properties
    # --------------------------------
    @property
    def workbook(self) -> gspread.Spreadsheet:
        """
        Gathers the authorization given the json credentials and returns a
        workbook object
        """
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

    @property
    def sheet(self) -> gspread.Worksheet:
        """
        Returns the sheet object corresponding to the workbook's sheet, given
        user input on sheet_id.
        """
        if isinstance(self.sheet_id, str):
            return self.workbook.worksheet(self.sheet_id)
        return self.workbook.get_worksheet(self.sheet_id)

    # --------------------------------
    # 1.2 - Main Methods
    # --------------------------------
    def to_google_sheet(self, fill_na_with=" ", clear_sheet=False, header=True):
        """
        Exports data frame to target sheet within a Google Sheet workbook.

        ---------------
        Before the upload, the function slightly adapts the content of the data
        frame to ensure  compatibility with Google Sheet. Dates are always
        turned to str and  missing values are filled with str. By default, it is
        also applied a correction to lists.

        For the upload, the batch update method is used to ensure the maximum
        efficiency possible. It also limits the number of request fired to the
        Google Sheet API.

        Arguments
        ----------------
        fill_na_with: str, default=" "
            The str missing values should be replaced by. This is necessary to
            avoid errors as Google Sheet does not accept missing values.
        clear_sheet: Bool, default=False
            If True, the whole sheet gets erased before the new data is uploaded
            to it. Note that the value of the destination cells is updated
            anyway.
        header: Bool, default=True
            If False, the header is not exported to the Google Sheet. The other
            table parts will still be placed in the cells as if the header was
            in place. This allows to manually edit the header after an export
            and preserved it after later refreshes.
        """
        self._prepare_table(fill_na_with)

        if clear_sheet:
            self.sheet.clear()

        self.sheet.batch_update(self._batch_list(header))

    # -------------------------------------------------------------------------
    # 2 - Worker Methods
    # -------------------------------------------------------------------------
    def _prepare_table(self, fill_na_with: str):
        """
        Converts datetime columns into str and fills missing values.
        """
        # Convert datetime columns into string
        for column in self.df.columns:
            if self.df[column].dtype in ["datetime64[ns]", "datetime64"]:
                self.df[column] = self.df[column].astype(str)
        # Replace missing values with something else
        self.df.fillna(fill_na_with, inplace=True)

    def _batch_list(self, keep_header: bool) -> List[Dict]:
        """
        Puts together the list of dict to be passed as argument of the
        batch_update method.

        ---------------
        The list is in the following form:
        [{"range":["A1", "A2", ...], "values":[[values row1], [values row2]]}]

        It can contain up to three dictionaries if header index and body are kept.
        """
        # Body is always exported
        output = [{"range" : self.body.cells_range,
                   "values" : self.df.values.tolist()}]
        # If header is kept, add to batch list
        if keep_header:
            output.append({"range" : self.header.cells_range,
                           "values" : self._columns_for_batch()})
        # If index is kept, add to batch list
        if self.keep_index:
            output.append({"range" : self.index.cells_range,
                           "values" : self._index_for_batch()})
        return output

    def _columns_for_batch(self) -> Union[List[List], List[str]]:
        """
        It reshapes the output of dataframe.columns.values.tolist() making it
        adapt for the batch_update() method.

        ---------------
        The key here is that each row of values should be on a separated list,
        otherwise the batch_update() method will throw an error. Thus, in case
        of multicolumns, the content should be reshaped in n list, where n is
        the multicolumns's depth.
        """
        # Handling case with multicolumns
        if self.indexes_depth[1] > 1:
            output = []
            for level in range(self.indexes_depth[1]):
                output.append([i[level] for i in self.df.columns.values.tolist()])
            return output
        # Handling case with simple columns
        return [self.df.columns.values.tolist()]

    def _index_for_batch(self) -> Union[List[List], List[str]]:
        """
        It reshapes the output of dataframe.index.values.tolist() making it
        adapt for the batch_update() method.

        ---------------
        The key here is that each row of values should be on a separated list,
        otherwise the batch_update() method will throw an error. Thus, if there
        is no multiindex, each index value should be placed in its own list.
        """
        # Handling case with multiindex
        if self.indexes_depth[0] > 1:
            return self.df.index.values.tolist()
        # Handling case with simple index
        return [[x] for x in  self.df.index.values.tolist()]
