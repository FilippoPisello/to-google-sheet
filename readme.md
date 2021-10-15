[![Code style: black](https://img.shields.io/badge/code%20style-black-000000.svg)](https://github.com/psf/black)
[![Python 3.8+](https://img.shields.io/badge/python-3.5+-blue.svg)](https://www.python.org/downloads/release/python-360/)
# Export pandas data frames to Google Sheet <!-- omit in toc -->
This module allows you to export a pandas data frame into Google Sheet in just two lines of code.

## Table of Contents <!-- omit in toc -->
- [1. Overview.](#1-overview)
  - [1.1. Goal of the project](#11-goal-of-the-project)
  - [1.2. Code structure](#12-code-structure)
- [2. Usage examples](#2-usage-examples)
  - [2.1. Usage with implicit json oauth file call](#21-usage-with-implicit-json-oauth-file-call)
  - [2.2. Usage with explicit json oauth file call](#22-usage-with-explicit-json-oauth-file-call)
- [2. Required packages](#2-required-packages)
- [3. Class elements](#3-class-elements)
  - [3.1. Arguments](#31-arguments)
  - [3.2. Attributes](#32-attributes)
  - [3.3. Properties](#33-properties)
  - [3.4. Methods](#34-methods)
    - [3.4.1. to_google_sheet()](#341-to_google_sheet)

## 1. Overview.
The GoogleSheet class allows the user to rapidly export a pandas data frame to a sheet within a Google Sheet workbook. It relies mostly on the gspread module which works as an intermediary with the Google Sheet API.

To be able to complete these operations, some quick steps on the Google Console should be carried out. [This article from the gspread library ](https://gspread.readthedocs.io/en/latest/oauth2.html) provides a detailed explanation on what to do. In summary:
- Obtain a json file with the credentials through the Google Developers Console.
- Place the file in the right folder of your device if you want to call it freely from any script.

CustomExcel is a subclass of the Spreadsheet class, whose code and doc can be found [at this link](https://github.com/FilippoPisello/Spreadsheet).

### 1.1. Goal of the project
The aim of this class is to speed up the upload process to Google Sheet, making it as similar as possible to a regular output to excel, as it happens in pandas through the built-in method dataframe.to_excel().

## 2. Usage examples
As an example, suppose that you want to send the pandas data frame *df* to the first sheet of the Google Sheet workbook named *"MyWork"*. For this to work correctly, the workbook **needs to be shared with edit permission with the email specified in the json authentication file**.

### 2.1. Usage with implicit json oauth file call
If you choose to save the json oauth file as "credentials.json" in the default folder specified by the gspread library ([doc](https://gspread.readthedocs.io/en/latest/oauth2.html)) - in my case it was *"C:\Users\MyUser\AppData\Roaming\gspread"* - the usage is as follows:
```python
sheet = GoogleSheet(dataframe=df, google_workbook_id="MyWork")
sheet.to_google_sheet()
```
### 2.2. Usage with explicit json oauth file call
The json file can also be kept in any other arbitrary folder. In this case, the path to the json file needs to be passed for the parameter **auth_keys**. For the following example suppose that the json file was named  *"credentials.json"* and placed in the *"auth"* folder inside the working directory. The the usage is the following:
```python
sheet = GoogleSheet(dataframe=df, google_workbook_id="MyWork", auth_keys="auth\credentials.json")
sheet.to_google_sheet()
```
## 3. Required packages
CustomExcel requires the following custom module created by me:
- **Spreadsheet** [_link_](https://github.com/FilippoPisello/Spreadsheet)

This class relies on the following built-in packages:
- **typing**
- **string** _[by Spreadsheet class]_

And on the following additional packages:
- **gspread**
- **numpy** _[by Spreadsheet class]_

## 3. Class elements
### 3.1. Arguments
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

### 3.2. Attributes
The CustomExcel object inherits four attributes from the Spreadsheet class:
- **self.df** : pandas data frame object
- **self.keep_index** : Bool
- **self.skip_rows**: int
- **self.skip_cols**: int

There are then two native attributes:
- **self.sheet**: gspread.Spreadsheet object
- **self.workbook**: gspread.Worksheet object

### 3.3. Properties
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

### 3.4. Methods
This section just includes the methods which are meant to be accessed by the user. These can be found in part 1. For further info on the worker methods please consult their docstrings.

#### 3.4.1. to_google_sheet()
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