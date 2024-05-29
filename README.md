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

The `ExcelXmlPacket` class is designed to handle and manipulate Excel files in XML format. This class provides functionalities to extract data from Excel files, read various components like worksheets and shared strings, and format cells. it contains functions that implements setting up and converting the excel to xml. when initialized, the excel 

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

### 3. **ExcelWriter Class**

The `ExcelWriter` class provides a method to write a DataTable to an Excel file by generating XML content that conforms to the OpenXML spreadsheet format.

#### Key Components:

- **Method:**
  - `WriteTableToNewExcel(dt As DataTable, excelFilePath As String, sheetName As String)`: This method generates an XML document for the worksheet data from the given DataTable and saves it as an Excel file.
    - It creates a worksheet XML structure with namespaces.
    - It adds rows and cells to the worksheet based on the DataTable.
    - It saves the XML document to a specified file path and renames it to have an `.xlsx` extension.

### Summary

- The **ExcelXmlPacket** class focuses on reading and extracting data from an existing Excel file.
- The **Reusables** class provides utility functions for common Excel-related operations.
- The **ExcelWriter** class (added in the latest explanation) focuses on writing data back to a new Excel file by generating the necessary XML content. 

These classes together provide a comprehensive set of functionalities for handling Excel files, from reading and manipulating data to writing it back into the correct format.


Overall, the ExcelXmlPacket Utility streamlines the process of working with Excel files in XML format, providing developers with essential functions for data extraction and manipulation.

# Developer : David Oku
Download code at : https://github.com/MasterOfLogic1/ExcelXmlManager





