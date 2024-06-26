# ExcelXmlPowerPack Library

## Overview
The `ExcelXmlMain` namespace provides functionality for interacting with Excel files through the `ExcelXmlAction` class. This class allows you to read and manipulate Excel sheets by performing various actions such as reading cell values, adding or deleting sheets, and more.

## Public Class: `ExcelXmlAction`

### Constructor

#### `New(filePath As String)`
Initializes a new instance of the `ExcelXmlAction` class with the specified file path.

- **Parameters:**
  - `filePath`: A string representing the path to the Excel file.

## Public Methods

### 1. `ReadCellValue(sheetName As String, cellReference As String) As String`
Returns the value of a specified cell in the given sheet.

- **Parameters:**
  - `sheetName`: The name of the sheet.
  - `cellReference`: The cell reference (e.g., "A1").

- **Example:**
  ```vb
  Dim cellValue As String = excelAction.ReadCellValue("Sheet1", "A1")
  ```

### 2. `GetAllSheetNames() As String[]`
Returns an array of all sheet names in the Excel file.

- **Parameters:** None

- **Example:**
  ```vb
  Dim sheetNames As String() = excelAction.GetAllSheetNames()
  ```

### 3. `GetSheetByIndex(index As Integer) As String`
Returns the name of the sheet at the specified index.

- **Parameters:**
  - `index`: The index of the sheet (0-based).

- **Example:**
  ```vb
  Dim sheetName As String = excelAction.GetSheetByIndex(0)
  ```

### 4. `GetSheetIndexByName(sheetName As String) As Integer?`
Returns the index of the sheet with the specified name.

- **Parameters:**
  - `sheetName`: The name of the sheet.

- **Example:**
  ```vb
  Dim sheetIndex As Integer? = excelAction.GetSheetIndexByName("Sheet1")
  ```

### 5. `GetLastUsedRow(sheetName As String) As Integer`
Returns the index of the last used row in the specified sheet.

- **Parameters:**
  - `sheetName`: The name of the sheet.

- **Example:**
  ```vb
  Dim lastUsedRow As Integer = excelAction.GetLastUsedRow("Sheet1")
  ```

### 6. `GetLastUsedColumn(sheetName As String) As Object[]`
Returns the letter and index of the last used column in the specified sheet.

- **Parameters:**
  - `sheetName`: The name of the sheet.

- **Example:**
  ```vb
  Dim lastUsedColumn As Object() = excelAction.GetLastUsedColumn("Sheet1")
  ```

### 7. `GetUsedRange(sheetName As String) As Object[]`
Returns the used range of the specified sheet as an array of two cell references.

- **Parameters:**
  - `sheetName`: The name of the sheet.

- **Example:**
  ```vb
  Dim usedRange As Object() = excelAction.GetUsedRange("Sheet1")
  ```

### 8. `DeleteSheet(sheetName As String)`
Deletes the specified sheet from the Excel file.

- **Parameters:**
  - `sheetName`: The name of the sheet.

- **Example:**
  ```vb
  excelAction.DeleteSheet("Sheet1")
  ```

### 9. `AddSheet(sheetName As String)`
Adds a new sheet with the specified name to the Excel file.

- **Parameters:**
  - `sheetName`: The name of the new sheet.

- **Example:**
  ```vb
  excelAction.AddSheet("NewSheet")
  ```

### 10. `RenameSheet(oldSheetName As String, newSheetName As String)`
Renames an existing sheet to a new name.

- **Parameters:**
  - `oldSheetName`: The current name of the sheet.
  - `newSheetName`: The new name for the sheet.

- **Example:**
  ```vb
  excelAction.RenameSheet("OldSheetName", "NewSheetName")
  ```

### 11. `HideSheet(sheetName As String)`
Hides the specified sheet.

- **Parameters:**
  - `sheetName`: The name of the sheet to hide.

- **Example:**
  ```vb
  excelAction.HideSheet("Sheet1")
  ```

### 12. `UnhideSheet(sheetName As String)`
Unhides the specified sheet.

- **Parameters:**
  - `sheetName`: The name of the sheet to unhide.

- **Example:**
  ```vb
  excelAction.UnhideSheet("Sheet1")
  ```

### 13. `AddColorToRange(sheetName As String, colorHex As String, Optional cellRange As String = Nothing)`
Adds color to a specified range of cells in the sheet. If no range is specified, it colors the entire used range.

- **Parameters:**
  - `sheetName`: The name of the sheet.
  - `colorHex`: The color in hex format (e.g., "FF0000" for red).
  - `cellRange` (Optional): The cell range to color (e.g., "A1:C5"). If omitted, the entire used range is colored.

- **Example:**
  ```vb
  excelAction.AddColorToRange("Sheet1", "FF0000", "A1:C5")
  ```

### 14. `DeleteRange(sheetName As String, Optional cellRange As String = Nothing)`
Deletes a specified range of cells in the sheet. If no range is specified, it deletes the entire used range.

- **Parameters:**
  - `sheetName`: The name of the sheet.
  - `cellRange` (Optional): The cell range to delete (e.g., "A1:C5"). If omitted, the entire used range is deleted.

- **Example:**
  ```vb
  excelAction.DeleteRange("Sheet1", "A1:C5")
  ```

### 15. `ReadSheetToDataTable(sheetName As String, Optional cellRange As String = Nothing, Optional hasHeader As Boolean = False)`
Reads an Excel sheet into a `DataTable`.

- **Parameters:**
  - `sheetName`: The name of the sheet to read.
  - `cellRange` (Optional): The range of cells to read (e.g., "A1:C10"). Defaults to the entire sheet if not specified.
  - `hasHeader` (Optional): Indicates if the first row should be treated as a header. Defaults to `False`.

- **Returns:**
  - `System.Data.DataTable`: A DataTable containing the sheet data.

- **Example:**
  ```vb
  Dim dataTable As System.Data.DataTable = excelAction.ReadSheetToDataTable("Sheet1", "A1:C10", True)
  ```

### 16. `WriteDataTableToSheet(filePath As String, sheetName As String, dataTable As System.Data.DataTable, Optional startCell As String = "A1", Optional AddHeader As Boolean = True)`
Writes a `DataTable` to an Excel sheet, starting from a specified cell.

- **Parameters:**
  - `filePath`: The file path of the Excel document.
  - `sheetName`: The name of the sheet to write to.
  - `dataTable`: The DataTable to write.
  - `startCell` (Optional): The cell reference to start writing from (e.g., "A1"). Defaults to "A1".
  - `AddHeader` (Optional): Indicates if column headers should be written. Defaults to `True`.

- **Example:**
  ```vb
  excelAction.WriteDataTableToSheet("path/to/excel.xlsx", "Sheet1", dataTable, "A1", True)
  ```

### 17. `AppendDataTableToSheet(filePath As String, sheetName As String, dataTable As System.Data.DataTable)`
Appends a `DataTable` to an existing Excel sheet.

- **Parameters:**
  - `filePath`: The file path of the Excel document.
  - `sheetName`: The name of the sheet to append to.
  - `dataTable`: The DataTable to append.

- **Example:**
  ```vb
  excelAction.AppendDataTableToSheet("path/to/excel.xlsx", "Sheet1", dataTable)
  ```

## Example Usage In Blue Prism

Certainly! Here's the updated section for Blue Prism usage, including references to required DLLs and namespace imports:

---

## Example Usage In Blue Prism

This library offers a comprehensive suite of methods for interacting with Excel files programmatically, making it ideal for automating Excel tasks in applications like Blue Prism for robotic process automation (RPA). It supports functionalities such as reading and writing cell values, managing sheets, and applying styles. The VBO ensures robust operation and provides clear diagnostics in case of failures.

### External References Required:

1. **DocumentFormat.OpenXml.dll**
2. **DocumentFormat.OpenXml.Framework.dll**
3. **DocumentFormat.OpenXml.Features.dll**
4. **ExcelXmlPowerPack.dll**

### Namespace Imports:

- `ExcelXmlPowerPack.ExcelXmlMain`

### Examples Implemented in Blue Prism:

To use these functionalities in Blue Prism, import the necessary DLLs into your VBO and implement each example usage in a code stage on a designated page.

1. **Reading Cell Value:**

   ```vb
   Dim excelAction As New ExcelXmlAction("path_to_excel_file.xlsx")
   Dim cellValue As String = excelAction.ReadCellValue("Sheet1", "A1")
   ```

2. **Getting All Sheet Names:**

   ```vb
   Dim excelAction As New ExcelXmlAction("path_to_excel_file.xlsx")
   Dim sheetNames As String() = excelAction.GetAllSheetNames()
   ```

3. **Adding a New Sheet:**

   ```vb
   Dim excelAction As New ExcelXmlAction("path_to_excel_file.xlsx")
   excelAction.AddSheet("NewSheet")
   ```

4. **Writing DataTable to Sheet:**

   ```vb
   Dim dataTable As System.Data.DataTable = New System.Data.DataTable()
   ' Populate the DataTable with data

   Dim excelAction As New ExcelXmlAction("path_to_excel_file.xlsx")
   excelAction.WriteDataTableToSheet("path_to_excel_file.xlsx", "Sheet1", dataTable, "A1", True)
   ```

5. **Deleting a Range of Cells:**

   ```vb
   Dim excelAction As New ExcelXmlAction("path_to_excel_file.xlsx")
   excelAction.DeleteRange("Sheet1", "A1:C5")
   ```

---

This section outlines how to integrate and utilize the `ExcelXmlPowerPack` VBO within Blue Prism, ensuring that you have the necessary dependencies and demonstrating typical use cases through implemented examples.

