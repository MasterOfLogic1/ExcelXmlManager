Imports System.Data
Imports System.IO
Imports System.IO.Compression
Imports System.Text.RegularExpressions

'The Reusable class contains sets of reusable actions that efficiently assist the ExcelXmlPacket class. While the class does not import the XML namespace directly,
'it is highly efficient due to its methods. These methods perform actions that enable the ExcelXmlPacket to carry out its tasks with minimal 
Public Class Reusables

    'when a letter is given to this function it retruns an excel column index - index starts from 0
    Public Shared Function ColumnLetterToIndex(columnLetter As String) As Integer
        ' Ensure the column letter is in uppercase
        columnLetter = columnLetter.ToUpper()
        Dim sum As Integer = 0
        Dim length As Integer = columnLetter.Length
        For i As Integer = 0 To length - 1
            ' Convert each character to its corresponding number (A=0, B=1, ..., Z=25)
            Dim charValue As Integer = Asc(columnLetter(i)) - Asc("A"c)
            ' Calculate the column index
            sum = sum * 26 + charValue
        Next
        Return sum
    End Function

    'when an index is given to this function it retruns an excel column letter - index starts from 0
    Public Shared Function ColumnIndexToLetter(columnIndex As Integer) As String
        Dim columnLetter As String = String.Empty
        Dim tempIndex As Integer = columnIndex
        While tempIndex >= 0
            Dim remainder As Integer = tempIndex Mod 26
            columnLetter = Chr(65 + remainder) & columnLetter
            tempIndex = Math.Floor(tempIndex / 26) - 1
        End While
        Return columnLetter
    End Function

    'this function seperates cell letter from row number in a cell address
    Public Shared Function SeparateColumnLetterAndRowNumber(cellAddress As String) As Object()
        Try
            ' Define a regular expression to match the column letters and row numbers
            Dim regex As New Regex("([A-Z]+)([0-9]+)")
            Dim match As Match = regex.Match(cellAddress)
            If match.Success Then
                Return {CObj(match.Groups(1).Value), CObj(CInt(match.Groups(2).Value))}
            Else
                ' Handle invalid input if necessary
                Throw New SystemException("Invalid cell Address. Please provide a valid excel cell address i.e A1")
            End If
        Catch
            Throw New SystemException("Invalid cell Address. Please provide a valid excel cell address i.e A1")
        End Try
    End Function


    ' Overloaded method for initializing table with header and column names or column count and default column names
    Public Shared Function InitializeTable(numberOfCols As Integer) As DataTable
        Dim dt As New DataTable
        Dim i As Integer = 1
        While (i <= numberOfCols)
            dt.Columns.Add("COLUMN_" + i.ToString(), GetType(String))
            i = i + 1
        End While
        Return dt
    End Function

    ' Overloaded method for initializing table with header and column names or column count and default column names
    Public Shared Function InitializeTable(columnNames As String()) As DataTable
        Dim dt As New DataTable
        Dim n As Integer = 0
        For Each colName As String In columnNames
            If String.IsNullOrEmpty(colName) Or String.IsNullOrWhiteSpace(colName) Then
                colName = "Column_" + n.ToString()
            End If
            dt.Columns.Add(colName, GetType(String))
            n = n + 1
        Next
        Return dt
    End Function
    Public Shared Sub CreateAFolder(folderPath As String)
        If Not Directory.Exists(folderPath) Then
            Directory.CreateDirectory(folderPath)
        End If
    End Sub

    'deletes a folder 
    Public Shared Sub DeleteAFolder(folderPath As String)
        If Directory.Exists(folderPath) Then
            Directory.Delete(folderPath, True)
        End If
    End Sub


    'deletes a file
    Public Shared Sub DeleteAFile(filePath As String)
        If File.Exists(filePath) Then
            File.Delete(filePath)
        End If
    End Sub


    'zips files and folders to a given zip file
    Public Shared Sub ZipAFile(sourceDirectoryPath As String, zipFilePath As String)
        If Directory.Exists(sourceDirectoryPath) Then
            If File.Exists(zipFilePath) Then
                'remove residue
                Reusables.DeleteAFile(zipFilePath)
            End If
            ZipFile.CreateFromDirectory(sourceDirectoryPath, zipFilePath)
        Else
            Throw New SystemException("the folder to zip was not found")
        End If
    End Sub

    'revert an patrent folder xml back to excel
    Public Shared Function RevertXmlParentFolderToExcelFile(excelFilePath As String, XmlParentFolderPath As String) As String
        Dim writeExcelFolderPath = Path.Combine(Path.GetPathRoot(Environment.SystemDirectory), "Automation", "XML_Processor", "write")
        'Delete the write folder path:
        Reusables.DeleteAFolder(writeExcelFolderPath)
        'create the write folder path inside of the processor folder:
        Reusables.CreateAFolder(writeExcelFolderPath)
        'first zip modified contents with the name of the zip file matching that of the excel
        Dim zipfilePath As String = Path.Combine(writeExcelFolderPath, "ex" + Now.ToString("ddMMyyHHmmss") + ".zip")
        'the generated excel file would be same as the zip file name:
        Dim newExcel As String = Path.Combine(writeExcelFolderPath, Path.GetFileNameWithoutExtension(zipfilePath) + Path.GetExtension(excelFilePath))
        'now create the zip file
        Reusables.ZipAFile(XmlParentFolderPath, zipfilePath)
        'now rename the zip file back to excel format:
        File.Copy(zipfilePath, newExcel)
        Reusables.DeleteAFile(zipfilePath)
        Reusables.DeleteAFile(excelFilePath)
        File.Move(newExcel, excelFilePath)
        Return newExcel
    End Function



    'creates a copy of an excel file into a target location and changing its extension from .xlsx to .zip
    Public Shared Function CreateACopyOfAFileByChangingExtension(filePath As String, folderPath As String) As String
        If File.Exists(filePath) AndAlso Directory.Exists(folderPath) Then
            Dim newFilePath As String = Path.Combine(folderPath, Path.GetFileNameWithoutExtension(filePath) + ".zip")
            File.Copy(filePath, Path.Combine(folderPath, newFilePath))
            Return newFilePath
        Else
            Throw New SystemException("intended file path and destination folder must be valid")
        End If
    End Function



    'unzips a given zip into a desired location
    Public Shared Sub UnzipAFile(zipFilePath As String, extractionFolderPath As String)
        If Directory.Exists(extractionFolderPath) Then
            If File.Exists(zipFilePath) Then
                ZipFile.ExtractToDirectory(zipFilePath, extractionFolderPath)
            Else
                Throw New SystemException("The zip file " + zipFilePath + " does not exist.")
            End If
        Else
            Throw New SystemException("the extarction destination path does not exists.")
        End If
    End Sub


End Class
