# 1. Documentation
## 1.1. Code structure
The code is designed to convey a hierarchical division of the class' methods. Different code portions are introduced by comment blocks which are made of two compact lines of "#" having a number in between.

The sections are structured as follows:
- **Part 1**: main elements
  - **Part 1.1**: properties
  - **Part 1.2**: main methods
- **Part 2**: worker methods

The methods' are ordered so that if method B is invoked by method A, then B will be below A. This structure should hopefully help the reader to understand how the simple pieces are assembled to construct more complex items.

The individual methods are designed to follow as closely as possible the **single-responsibility principle**. Some of them are **protected** - their name is preceded by an underscore. This is done for two main reasons. First, not to clutter excessively the help text of the class, since protected methods are not displayed in this output. This allows the focus to be kept on the most important elements. Second, protection is in place to clearly signal which are the methods meant to be used only internally.

## 1.2. GoogleSheet Class
The GoogleSheet class it the only class in the script. Here are further details on its usage.

### 1.2.1. Arguments
The class inherits four arguments from the Spreadsheet class:
- **dataframe** : pandas data frame object (mandatory)
  - The pandas data frame to be considered.
- **keep_index** : Bool, default=False
  - If True, it is taken into account that the first column of the spreadsheet will be occupied by the index. All the dimensions will be adjusted as consequence.
- **starting_cell**: str, default="A1"
  - The cell where it will be placed the top left corner of the dataframe.
- **correct_lists**: Bool, default=False
  - If True, the lists stored as the data frame entries are modified to be more readable in the traditional spreadsheet softwares. This happens in four ways. (1) Empty lists are replaced by missing values. (2) Missing values are removed from within the lists. (3) Lists of len 1 are replaced by the single element they contain. (4) Lists are replaced by str formed by their elements separated by commas.

There are then three native arguments of the class:
- **google_workbook_id**: str, (mandatory)
  - Either the title or the key of the workbook the data should be pushed to.
- **sheet_id**: str or int, default=0
  - Argument to identify the target sheet within the workbook. If int, it is interpreted as sheet index, if str as the sheet name. If str matches no existing sheet, a new one is created.
- **auth_keys**: None or str or dict, default=None
  - If None, it is assumed that the json file for the authentication is in
  the default folder "~/.config/gspread/your_file.json". If str, it is the custom path of the json authentication file. If dict, it contains the parsed content of the authentication file.

### 1.2.2. Attributes
The CustomExcel object inherits four attributes from the Spreadsheet class:
- **self.df** : pandas data frame object
- **self.keep_index** : Bool
- **self.skip_rows**: int
- **self.skip_cols**: int

There are then two native attributes:
- **self.sheet**: gspread.Spreadsheet object
- **self.workbook**: gspread.Worksheet object

### 1.2.3. Properties
The CustomExcel object inherits eight properties from the Spreadsheet class:
- **self.indexes_depth**: [int, int]
- **self.header_coordinates**: [[int, int], [int, int]]
- **self.index_coordinates**: [[int, int], [int, int]]
- **self.body_coordinates**: [[int, int], [int, int]]
- **self.header**: SpreadsheetElement object
- **self.index**: SpreadsheetElement object
- **self.body**: SpreadsheetElement object
- **self.table**: SpreadsheetElement object

Details on these can be found at this [_link_](https://github.com/FilippoPisello/Spreadsheet).

### 1.2.4. Methods
This section just includes the methods which are meant to be accessed by the user. These can be found in part 1. For further info on the worker methods please consult their docstrings.

#### 1.2.4.1. to_google_sheet()
```python
obj.to_google_sheet(self, fill_na_with=" ", clear_sheet=False, header=True)
```
Exports data frame to target sheet within a Google Sheet workbook.

Before the upload, the function slightly adapts the content of the data frame to ensure  compatibility with Google Sheet. Dates and categories are turned to str and missing values are filled with str. By default, it is also applied a correction to lists.

For the upload, the batch update method is used to ensure the maximum efficiency possible. It also limits the number of request fired to the Google Sheet API.

**Arguments**
- **fill_na_with**: str, default=" "
  - The str missing values should be replaced by. This is necessary to avoid errors as Google Sheet does not accept missing values.
- **clear_sheet**: Bool, default=False
  - If True, the whole sheet gets erased before the new data is uploaded to it. Note that the value of the destination cells is updated anyway.
- **header**: Bool, default=True
  - If False, the header is not exported to the Google Sheet. The other table parts will still be placed in the cells as if the header was in place. This allows to manually edit the header after an export and preserved it after later refreshes.