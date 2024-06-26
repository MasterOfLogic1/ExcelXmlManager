# ExcelXmlPowerPack Library

## Overview
The `ExcelXmlMain` namespace provides functionality for interacting with Excel files through the `ExcelXmlAction` class. This class allows you to read and manipulate Excel sheets by performing various actions such as reading cell values, adding or deleting sheets, and more.

## Public Class: `ExcelXmlAction`

### Constructor

#### `New(filePath As String)`
Initializes a new instance of the `ExcelXmlAction` class with the specified file path.

- **Parameters:**
  - `filePath`: A string representing the path to the Excel file.

### Public Methods

#### `ReadCellValue(sheetName As String, cellReference As String) As String`
Returns the value of a specified cell in the given sheet.

- **Parameters:**
  - `sheetName`: The name of the sheet.
  - `cellReference`: The cell reference (e.g., "A1").

#### `GetAllSheetNames() As String[]`
Returns an array of all sheet names in the Excel file.

- **Parameters:** None

#### `GetSheetByIndex(index As Integer) As String`
Returns the name of the sheet at the specified index.

- **Parameters:**
  - `index`: The index of the sheet (0-based).

#### `GetSheetIndexByName(sheetName As String) As Integer?`
Returns the index of the sheet with the specified name.

- **Parameters:**
  - `sheetName`: The name of the sheet.

#### `GetLastUsedRow(sheetName As String) As Integer`
Returns the index of the last used row in the specified sheet.

- **Parameters:**
  - `sheetName`: The name of the sheet.

#### `GetLastUsedColumn(sheetName As String) As Object[]`
Returns the letter and index of the last used column in the specified sheet.

- **Parameters:**
  - `sheetName`: The name of the sheet.

#### `GetUsedRange(sheetName As String) As Object[]`
Returns the used range of the specified sheet as an array of two cell references.

- **Parameters:**
  - `sheetName`: The name of the sheet.

#### `DeleteSheet(sheetName As String)`
Deletes the specified sheet from the Excel file.

- **Parameters:**
  - `sheetName`: The name of the sheet.

#### `AddSheet(sheetName As String)`
Adds a new sheet with the specified name to the Excel file.

- **Parameters:**
  - `sheetName`: The name of the new sheet.

#### `RenameSheet(oldSheetName As String, newSheetName As String)`
Renames an existing sheet to a new name.

- **Parameters:**
  - `oldSheetName`: The current name of the sheet.
  - `newSheetName`: The new name for the sheet.

#### `HideSheet(sheetName As String)`
Hides the specified sheet.

- **Parameters:**
  - `sheetName`: The name of the sheet to hide.

#### `UnhideSheet(sheetName As String)`
Unhides the specified sheet.

- **Parameters:**
  - `sheetName`: The name of the sheet to unhide.

#### `AddColorToRange(sheetName As String, colorHex As String, Optional cellRange As String = Nothing)`
Adds color to a specified range of cells in the sheet. If no range is specified, it colors the entire used range.

- **Parameters:**
  - `sheetName`: The name of the sheet.
  - `colorHex`: The color in hex format (e.g., "FF0000" for red).
  - `cellRange` (Optional): The cell range to color (e.g., "A1:C5"). If omitted, the entire used range is colored.

#### `DeleteRange(sheetName As String, Optional cellRange As String = Nothing)`
Deletes a specified range of cells in the sheet. If no range is specified, it deletes the entire used range.

- **Parameters:**
  - `sheetName`: The name of the sheet.
  - `cellRange` (Optional): The cell range to delete (e.g., "A1:C5"). If omitted, the entire used range is deleted.

## Usage
To use the `ExcelXmlAction` class, create an instance by passing the path to your Excel file, and then call the desired methods on the instance.

```vb
Dim excelAction As New ExcelXmlAction("path_to_excel_file.xlsx")
Dim cellValue As String = excelAction.ReadCellValue("Sheet1", "A1")
Dim sheetNames As String() = excelAction.GetAllSheetNames()
```

This README provides an overview of the public methods available in the `ExcelXmlAction` class. Each method is described with its parameters and what it does.
