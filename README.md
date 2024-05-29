# ExcelXmlManager
# Project Title: ExcelXmlPacket Utility

# Description:
The ExcelXmlPacket Utility is a collection of functions designed to interact with Excel files stored in XML format. It provides various functionalities to read data from Excel sheets, manipulate the data, and extract specific information such as sheet names, cell values, last used columns, and more.

# Features:

1.) Reading Excel Data: The utility can read data from Excel sheets stored in XML format.

2.) Sheet Information: It provides functions to retrieve information about the sheets present in the Excel file, such as names of the sheets.

3.) Cell Value Extraction: Users can extract the value of a specific cell by providing its address.

4.) Last Used Column: The utility determines the index of the last used column in a sheet.

5.) Error Handling: Proper error handling is implemented to ensure smooth execution and error reporting.


# Purpose: 
This code can be wrapped as an object or library in any RPA technology, hence providing us with methods to manage and interact with an Excel file as XML. This prevents the time wasted waiting for Excel to open up and performing traditional Excel VBO.
The Global Code contains a VB.Net  class named ExcelXmlPacket is designed to process Excel files. It takes the path to an Excel file as input and extracts information from it. The class creates a copy of the Excel file with a .zip extension and extracts its contents. Then, it reads data from the extracted XML files like worksheets and shared strings. Finally, it stores the extracted data in a structured format for further processing.

# Technical Architecture:
The program contains 3 Classes which are :

### 1. **ExcelXmlPacket Class**

The `ExcelXmlPacket` class is designed to setup the excel in XML format. it contains functions that implements setting up and converting the excel to xml. These functions simply implement the process of changing the extension of the given excel to .zip, unzipping the file and holding relevant data that would be used in other classes in vairables. 

#### Key Components:

- **Properties:**
  - `ExcelFilePath`, `xmlContentArchivedPath`, `WorksheetData`, `CellFormatTypes`, `SharedStringsData`, `ExcelSheetNames`: These properties provide access to the various paths and data structures used within the class.

- **Constructor:**
  - `New(excelFilePath As String)`: This constructor initializes the class with the path to an Excel file, sets up necessary file paths, and calls methods to generate XML files, read worksheet data, and shared strings, and get cell format types.

- **Methods:**
  - `GenerateExcelXml()`: Deletes any existing XML folder, creates a new one, copies the Excel file as a .zip file, and unzips it to extract its contents.
  - `CreateAFolder(folderPath As String)`: Creates a directory if it doesn't exist.
  - `DeleteAFolder(folderPath As String)`: Deletes a directory and its contents if it exists.
  - `CreateACopyOfAFileByChangingExtension(filePath As String, folderPath As String)`: Copies the Excel file to a new location with a .zip extension.
  - `UnzipAFile(zipFilePath As String, extractionFolderPath As String)`: Unzips the .zip file to the specified location.
  - `GetNameSpaceUri(xmlFilePath As String)`: Retrieves the namespace URI from an XML file.
  - `GetWorkSheetData(relFilePath As String, workbookXmlFilePath As String, xlFolderPath As String)`: Reads the worksheet data from the extracted XML files.
  - `GetCellFormatTypes(workbookXmlFilePath As String)`: Retrieves cell format types from the styles.xml file.

- **Destructor:**
  - `Dispose()`: Cleans up by deleting the generated XML folder.

### 2. **Reusables Class**

The `Reusables` class contains utility functions that are commonly used for various tasks, particularly related to Excel operations.

#### Key Components:

- **Methods:**
  - `ColumnLetterToIndex(columnLetter As String) As Integer`: Converts an Excel column letter (e.g., "A") to its corresponding index (0-based).
  - `ColumnIndexToLetter(columnIndex As Integer) As String`: Converts a column index (0-based) to its corresponding Excel column letter.
  - `SeparateColumnLetterAndRowNumber(cellAddress As String) As Object()`: Splits a cell address (e.g., "A1") into its column letter and row number components.
  - `InitializeTable(numberOfCols As Integer) As DataTable`: Initializes a DataTable with a specified number of columns, using default column names.
  - `InitializeTable(columnNames As String()) As DataTable`: Initializes a DataTable with specified column names.

### 3. **ExcelXmlAction Class**

provides a suite of static methods for various excel related operation i.e read cell value, last used range, last row, read excel to datatable e.t.c. This class is dependent on the `ExcelXmlPacket` class . Here's a detailed breakdown of the class and its methods

The third class, `ExcelXmlAction`, provides a suite of static methods to read and manipulate Excel data stored in XML format. This class is dependent on the `ExcelXmlPacket` class for extracting and managing the XML data. Here's a detailed breakdown of the class and its methods:

### Overview

The `ExcelXmlAction` class facilitates various operations on Excel files represented in XML format. It includes methods to read Excel data into a `DataTable`, fetch sheet names and indices, and determine the used range, last used row, and last used column in a sheet.

### Class Definition

```vb
Imports System.Data
Imports System.IO
Imports System.Xml
Imports ExcelXmlManager

Public Class ExcelXmlAction
```

### Methods

1. **ReadExcelToTable(excelFullFilePath As String, sheetName As String, hasHeader As Boolean) As DataTable**
   - Reads the specified sheet from the Excel file into a `DataTable`.
   - Parameters:
     - `excelFullFilePath`: Full path to the Excel file.
     - `sheetName`: Name of the sheet to read.
     - `hasHeader`: Indicates if the first row contains headers.
   - Processes shared strings and sheet data XML to populate the `DataTable`.

2. **ReadExcelToTable(excelFullFilePath As String, sheetIndex As Integer, hasHeader As Boolean) As DataTable**
   - An overload of the previous method, which reads the sheet based on its index.
   - Parameters:
     - `excelFullFilePath`: Full path to the Excel file.
     - `sheetIndex`: Index of the sheet to read.
     - `hasHeader`: Indicates if the first row contains headers.
   - Similar processing of shared strings and sheet data XML.

3. **GetSheetNameBySheetIndex(excelFullFilePath As String, sheetIndex As Integer) As String**
   - Retrieves the name of the sheet at the specified index.
   - Parameters:
     - `excelFullFilePath`: Full path to the Excel file.
     - `sheetIndex`: Index of the sheet.
   - Returns the name of the sheet or throws an error if not found.

4. **GetSheetIndexBySeetName(excelFullFilePath As String, sheetName As String) As Integer**
   - Retrieves the index of the sheet with the specified name.
   - Parameters:
     - `excelFullFilePath`: Full path to the Excel file.
     - `sheetName`: Name of the sheet.
   - Returns the index of the sheet or throws an error if not found.

5. **GetUsedRange(excelFullFilePath As String, sheetName As String) As String**
   - Retrieves the used range of cells in the specified sheet.
   - Parameters:
     - `excelFullFilePath`: Full path to the Excel file.
     - `sheetName`: Name of the sheet.
   - Returns the used range in A1 notation.

6. **GetLastUsedRow(excelFullFilePath As String, sheetName As String) As Integer**
   - Retrieves the index of the last used row in the specified sheet.
   - Parameters:
     - `excelFullFilePath`: Full path to the Excel file.
     - `sheetName`: Name of the sheet.
   - Returns the last used row index or throws an error if an issue arises.

7. **GetLastUsedColumn(excelFullFilePath As String, sheetName As String) As Integer**
   - Retrieves the index of the last used column in the specified sheet.
   - Parameters:
     - `excelFullFilePath`: Full path to the Excel file.
     - `sheetName`: Name of the sheet.
   - Returns the last used column index or throws an error if an issue arises.

8. **GetExcelSheetNames(excelFullFilePath As String) As DataTable**
   - Retrieves the names of all sheets in the Excel file.
   - Parameters:
     - `excelFullFilePath`: Full path to the Excel file.
   - Returns a `DataTable` containing the sheet names.

9. **GetCellValue(excelFullFilePath As String, sheetName As String, cellAddress As String) As String**
   - Retrieves the value of a specified cell.
   - Parameters:
     - `excelFullFilePath`: Full path to the Excel file.
     - `sheetName`: Name of the sheet.
     - `cellAddress`: Address of the cell in A1 notation.
   - Returns the cell value as a string or throws an error if an issue arises.

### Detailed Method Functionality

Each method follows a similar pattern:
1. **Instantiation**: Create an instance of `ExcelXmlPacket` to access the XML data.
2. **Error Handling**: Check for the existence of the specified sheet or index and handle errors appropriately.
3. **XML Processing**: Load and manipulate XML data to extract the required information.
4. **Return Results**: Return the extracted data or throw a `SystemException` if an error occurs.

### Summary

The `ExcelXmlAction` class is a utility for reading and manipulating Excel data stored in XML format. It provides methods for accessing sheet names and indices, reading cell values, and determining the structure and content of Excel sheets. This class relies heavily on the `ExcelXmlPacket` class for parsing and handling the XML data.


Overall, the ExcelXmlPacket Utility streamlines the process of working with Excel files in XML format, providing developers with essential functions for data extraction and manipulation.

# Developer : David Oku
Download code at : https://github.com/MasterOfLogic1/ExcelXmlManager





