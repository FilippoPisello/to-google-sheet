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
- [3. Required packages](#3-required-packages)
- [4. Class elements](#4-class-elements)
  - [4.1. Arguments](#41-arguments)
  - [4.2. Attributes](#42-attributes)
  - [4.3. Properties](#43-properties)
  - [4.4. Methods](#44-methods)
    - [4.4.1. to_google_sheet()](#441-to_google_sheet)

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

# Documentation
Consult the documentation [here](docs.md).