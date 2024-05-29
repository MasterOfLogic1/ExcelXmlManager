Imports System.Data
Imports System.IO
Imports System.Xml
Imports ExcelXmlManager

Public Class ExcelXmlAction
    'Reads an excel to a datatable
    Public Shared Function ReadExcelToTable(excelFullFilePath As String, sheetName As String, hasHeader As Boolean) As DataTable
        Dim dt As DataTable = Nothing
        Dim errorMessage As String = String.Empty

        Try
            ' Instantiate custom xml class
            Dim ExcelXmlPacket As New ExcelXmlPacket(excelFullFilePath)

            ' Throw error if sheet not present
            If Not ExcelXmlPacket.WorksheetData.Keys().Contains(sheetName) Then
                Throw New SystemException("Sheet not found: [" & sheetName & "]")
            End If

            ' Get all xml data that are used to make up the data found in the given excel sheet
            Dim SheetData As String = ExcelXmlPacket.WorksheetData(sheetName)("data").ToString()
            Dim sharedStringData As String = ExcelXmlPacket.SharedStringsData("data").ToString()
            Dim nameSpaceUri As String = ExcelXmlPacket.SharedStringsData("namespaceuri").ToString()

            ' Process the shared string data which include all string data that make up the sheet
            Dim sharedStringsMapper As New Dictionary(Of Integer, String)
            Dim sharedStringsXml As New XmlDocument()
            sharedStringsXml.LoadXml(sharedStringData)

            ' Create a namespace manager for shared string
            Dim nsManager As New XmlNamespaceManager(sharedStringsXml.NameTable)
            nsManager.AddNamespace("ns", nameSpaceUri)

            ' Select nodes with the namespace prefix in the shared strings 
            Dim nodesOfSharedStrings As XmlNodeList = sharedStringsXml.SelectNodes("//ns:t", nsManager)

            ' Process the selected nodes into the shared strings Mapper: map the data on the shared strings file to the sheet.xml file found in worksheets
            Dim nodeIndex As Integer = 0
            For Each node As XmlNode In nodesOfSharedStrings
                Dim stringValue As String = node.InnerText
                ' Enter shared string into shared string mapper
                sharedStringsMapper(nodeIndex) = stringValue
                ' Increment node index
                nodeIndex += 1
            Next

            ' Process the sheet data which include all non-string data that make up the sheet
            Dim sheetXml As New XmlDocument()
            sheetXml.LoadXml(SheetData)
            ' Get all the row nodes
            Dim nl As List(Of XmlNode) = (From rn As XmlNode In sheetXml.SelectNodes("//ns:row", nsManager).Cast(Of XmlNode)() Select rn).ToList()

            ' Find the row node with the highest number of "ns:c" child nodes and get the count of those child nodes
            Dim maxColsNode As XmlNode = nl.OrderByDescending(Function(x) x.SelectNodes("ns:c", nsManager).Count).FirstOrDefault()

            ' Check if maxColsNode is not Nothing (to handle cases where nl might be empty)
            Dim numberOfCols As Integer = If(maxColsNode IsNot Nothing, maxColsNode.SelectNodes("ns:c", nsManager).Count, 0)

            ' Iterate over each row node
            For Each rowNode As XmlNode In sheetXml.SelectNodes("//ns:row", nsManager)
                Dim rowIndex As Integer = Integer.Parse(rowNode.Attributes("r").Value)

                If (hasHeader) Then
                    rowIndex = rowIndex - 1
                End If

                While dt IsNot Nothing AndAlso dt.Rows.Count < rowIndex - 1
                    ' Add empty rows until the rowIndex is reached
                    dt.Rows.Add(New String(numberOfCols - 1) {})
                End While

                Dim rowData As New List(Of String)(New String(numberOfCols - 1) {})
                ' Select each column here
                For Each cellNode As XmlNode In rowNode.SelectNodes("ns:c", nsManager)
                    Dim nd As String = String.Empty
                    Dim cellValue As String = String.Empty
                    Dim colRef As String = cellNode.Attributes("r").Value
                    Dim colLetter As String = Reusables.SeparateColumnLetterAndRowNumber(colRef)(0)
                    Dim colIndex As Integer = Reusables.ColumnLetterToIndex(colLetter) ' Function to convert column reference to index

                    ' Trying to get cell value
                    Dim dataType As String
                    nd = cellNode.SelectSingleNode("ns:v", nsManager)?.InnerText
                    If Not String.IsNullOrEmpty(nd) Then
                        dataType = cellNode.Attributes("t")?.Value
                        If Not String.IsNullOrEmpty(dataType) Then
                            'Since it is "s" and "s" represents string, then get actual value from shared strings dictionary
                            cellValue = sharedStringsMapper(CInt(nd))
                        Else
                            'cell value might be in some other format
                            Dim targetFmtIndex As String = cellNode.Attributes("s")?.Value
                            Try
                                Dim fmt As String = ExcelXmlPacket.CellFormatTypes(CInt(targetFmtIndex))

                                Select Case fmt
                                    Case "Double"
                                        cellValue = CDbl(nd)
                                    Case "General"
                                        cellValue = nd
                                    Case "Unknown"
                                        cellValue = nd
                                    Case Else
                                        cellValue = DateTime.FromOADate(nd).ToString(fmt)
                                        cellValue = DateTime.ParseExact(cellValue, fmt, Nothing).ToString("dd/MM/yyyy HH:mm:ss")
                                End Select


                            Catch
                                cellValue = nd
                            End Try

                        End If

                    Else
                        'cell value is found as blank

                        cellValue = ""
                    End If
                    'add cell value to row list
                    rowData(colIndex) = cellValue

                Next

                If dt Is Nothing Then
                    If hasHeader Then
                        ' If excel has header then initialize table with first row as header
                        dt = Reusables.InitializeTable(rowData.ToArray())
                    Else
                        ' Create generic headers
                        dt = Reusables.InitializeTable(rowData.Count)
                        dt.Rows.Add(rowData.ToArray())
                    End If
                Else
                    dt.Rows.Add(rowData.ToArray())
                End If
            Next

            ExcelXmlPacket.Dispose()

        Catch ex As Exception
            errorMessage = ex.Message
        End Try

        If Not String.IsNullOrEmpty(errorMessage) Then
            Throw New SystemException(errorMessage)
        End If

        Return dt
    End Function


    'Overrides the first function
    Public Shared Function ReadExcelToTable(excelFullFilePath As String, sheetIndex As Integer, hasHeader As Boolean) As DataTable
        Dim dt As DataTable = Nothing
        Dim errorMessage As String = String.Empty

        Try
            ' Instantiate custom xml class
            Dim ExcelXmlPacket As New ExcelXmlPacket(excelFullFilePath)

            ' Return all the available sheet indices
            Dim indices As Integer() = ExcelXmlPacket.WorksheetData.Values().Select(Function(v) CInt(v("sheetindex").ToString())).ToArray()

            ' Throw error if sheet not present
            If Not indices.Contains(sheetIndex) Then
                Throw New SystemException("No sheet found at index: [" & sheetIndex.ToString() & "]")
            End If

            ' Get the sheet name related to the given index
            Dim sheetName As String = ExcelXmlPacket.WorksheetData.First(Function(p) Convert.ToInt32(p.Value("sheetindex")) = sheetIndex).Key()

            ' Get all xml data that are used to make up the data found in the given excel sheet
            ' Get all xml data that are used to make up the data found in the given excel sheet
            Dim SheetData As String = ExcelXmlPacket.WorksheetData(sheetName)("data").ToString()
            Dim sharedStringData As String = ExcelXmlPacket.SharedStringsData("data").ToString()
            Dim nameSpaceUri As String = ExcelXmlPacket.SharedStringsData("namespaceuri").ToString()

            ' Process the shared string data which include all string data that make up the sheet
            Dim sharedStringsMapper As New Dictionary(Of Integer, String)
            Dim sharedStringsXml As New XmlDocument()
            sharedStringsXml.LoadXml(sharedStringData)

            ' Create a namespace manager for shared string
            Dim nsManager As New XmlNamespaceManager(sharedStringsXml.NameTable)
            nsManager.AddNamespace("ns", nameSpaceUri)

            ' Select nodes with the namespace prefix in the shared strings 
            Dim nodesOfSharedStrings As XmlNodeList = sharedStringsXml.SelectNodes("//ns:t", nsManager)

            ' Process the selected nodes into the shared strings Mapper: map the data on the shared strings file to the sheet.xml file found in worksheets
            Dim nodeIndex As Integer = 0
            For Each node As XmlNode In nodesOfSharedStrings
                Dim stringValue As String = node.InnerText
                ' Enter shared string into shared string mapper
                sharedStringsMapper(nodeIndex) = stringValue
                ' Increment node index
                nodeIndex += 1
            Next

            ' Process the sheet data which include all non-string data that make up the sheet
            Dim sheetXml As New XmlDocument()
            sheetXml.LoadXml(SheetData)
            ' Get all the row nodes
            Dim nl As List(Of XmlNode) = (From rn As XmlNode In sheetXml.SelectNodes("//ns:row", nsManager).Cast(Of XmlNode)() Select rn).ToList()

            ' Find the row node with the highest number of "ns:c" child nodes and get the count of those child nodes
            Dim maxColsNode As XmlNode = nl.OrderByDescending(Function(x) x.SelectNodes("ns:c", nsManager).Count).FirstOrDefault()

            ' Check if maxColsNode is not Nothing (to handle cases where nl might be empty)
            Dim numberOfCols As Integer = If(maxColsNode IsNot Nothing, maxColsNode.SelectNodes("ns:c", nsManager).Count, 0)

            ' Iterate over each row node
            For Each rowNode As XmlNode In sheetXml.SelectNodes("//ns:row", nsManager)
                Dim rowIndex As Integer = Integer.Parse(rowNode.Attributes("r").Value)

                If (hasHeader) Then
                    rowIndex = rowIndex - 1
                End If

                While dt IsNot Nothing AndAlso dt.Rows.Count < rowIndex - 1
                    ' Add empty rows until the rowIndex is reached
                    dt.Rows.Add(New String(numberOfCols - 1) {})
                End While

                Dim rowData As New List(Of String)(New String(numberOfCols - 1) {})
                ' Select each column here
                For Each cellNode As XmlNode In rowNode.SelectNodes("ns:c", nsManager)
                    Dim nd As String = String.Empty
                    Dim cellValue As String = String.Empty
                    Dim colRef As String = cellNode.Attributes("r").Value
                    Dim colLetter As String = Reusables.SeparateColumnLetterAndRowNumber(colRef)(0)
                    Dim colIndex As Integer = Reusables.ColumnLetterToIndex(colLetter) ' Function to convert column reference to index

                    ' Trying to get cell value
                    Dim dataType As String
                    nd = cellNode.SelectSingleNode("ns:v", nsManager)?.InnerText
                    If Not String.IsNullOrEmpty(nd) Then
                        dataType = cellNode.Attributes("t")?.Value
                        If Not String.IsNullOrEmpty(dataType) Then
                            'Since it is "s" and "s" represents string, then get actual value from shared strings dictionary
                            cellValue = sharedStringsMapper(CInt(nd))
                        Else
                            'cell value might be in some other format
                            Dim targetFmtIndex As String = cellNode.Attributes("s")?.Value
                            Try
                                Dim fmt As String = ExcelXmlPacket.CellFormatTypes(CInt(targetFmtIndex))

                                Select Case fmt
                                    Case "Double"
                                        cellValue = CDbl(nd)
                                    Case "General"
                                        cellValue = nd
                                    Case "Unknown"
                                        cellValue = nd
                                    Case Else
                                        cellValue = DateTime.FromOADate(nd).ToString(fmt)
                                        cellValue = DateTime.ParseExact(cellValue, fmt, Nothing).ToString("dd/MM/yyyy HH:mm:ss")
                                End Select


                            Catch
                                cellValue = nd
                            End Try

                        End If

                    Else
                        'cell value is found as blank

                        cellValue = ""
                    End If
                    'add cell value to row list
                    rowData(colIndex) = cellValue

                Next

                If dt Is Nothing Then
                    If hasHeader Then
                        ' If excel has header then initialize table with first row as header
                        dt = Reusables.InitializeTable(rowData.ToArray())
                    Else
                        ' Create generic headers
                        dt = Reusables.InitializeTable(rowData.Count)
                        dt.Rows.Add(rowData.ToArray())
                    End If
                Else
                    dt.Rows.Add(rowData.ToArray())
                End If
            Next

            ExcelXmlPacket.Dispose()

        Catch ex As Exception
            errorMessage = ex.Message
        End Try

        If Not String.IsNullOrEmpty(errorMessage) Then
            Throw New SystemException(errorMessage)
        End If

        Return dt
    End Function

    'returns the sheetname of an excel when given a sheet index
    Public Shared Function GetSheetNameBySheetIndex(excelFullFilePath As String, sheetIndex As Integer) As String
        Dim errorMessage As String = String.Empty
        Dim sheetName As String = Nothing

        Try
            ' Instantiate custom xml class
            Dim ExcelXmlPacket As New ExcelXmlPacket(excelFullFilePath)

            ' Return all the available sheet indices
            Dim indices As Integer() = ExcelXmlPacket.WorksheetData.Values().Select(Function(v) CInt(v("sheetindex").ToString())).ToArray()

            ' Throw error if sheet not present
            If Not indices.Contains(sheetIndex) Then
                Throw New SystemException("No sheet found at index: [" & sheetIndex.ToString() & "]")
            End If

            ' Get the sheet name related to the given index
            sheetName = ExcelXmlPacket.WorksheetData.First(Function(p) Convert.ToInt32(p.Value("sheetindex")) = sheetIndex).Key()

            ExcelXmlPacket.Dispose()
        Catch ex As Exception
            errorMessage = ex.Message
        End Try

        If Not String.IsNullOrEmpty(errorMessage) Then
            Throw New SystemException(errorMessage)
        End If

        Return sheetName
    End Function
    'returns the sheet index of an excel when given a sheet name
    Public Shared Function GetSheetIndexBySeetName(excelFullFilePath As String, sheetName As String) As Integer
        Dim errorMessage As String = String.Empty
        Dim sheetIndex As Integer = -1

        Try
            ' Instantiate custom xml class
            Dim ExcelXmlPacket As New ExcelXmlPacket(excelFullFilePath)

            ' Throw error if sheet not present
            If Not ExcelXmlPacket.WorksheetData.Keys().Contains(sheetName) Then
                Throw New SystemException("Sheet not found: [" & sheetName & "]")
            End If

            ' Get sheet index
            sheetIndex = CInt(ExcelXmlPacket.WorksheetData(sheetName)("sheetindex"))

            ExcelXmlPacket.Dispose()
        Catch ex As Exception
            errorMessage = ex.Message
        End Try

        If Not String.IsNullOrEmpty(errorMessage) Then
            Throw New SystemException(errorMessage)
        End If

        Return sheetIndex
    End Function

    'returns the range used in an excel
    Public Shared Function GetUsedRange(excelFullFilePath As String, sheetName As String) As String
        Dim errorMessage As String = String.Empty
        Dim usedRange As String = Nothing

        Try
            ' Instantiate custom xml class
            Dim ExcelXmlPacket As New ExcelXmlPacket(excelFullFilePath)

            ' Throw error if sheet not present
            If Not ExcelXmlPacket.WorksheetData.Keys().Contains(sheetName) Then
                Throw New SystemException("Sheet not found: [" & sheetName & "]")
            End If

            ' Get all xml data that are used to make up the data found in the given excel sheet
            Dim sheetData As String = ExcelXmlPacket.WorksheetData(sheetName)("data").ToString()
            ' Get expected namespace URI
            Dim namespaceUri As String = ExcelXmlPacket.SharedStringsData("namespaceuri").ToString()

            Dim doc As New XmlDocument()
            doc.LoadXml(sheetData)

            ' Create a namespace manager
            Dim nsManager As New XmlNamespaceManager(doc.NameTable)
            nsManager.AddNamespace("ns", namespaceUri)

            ' Select nodes with the namespace prefix in the dimension 
            Dim dimensionNode As XmlNode = doc.SelectSingleNode("//ns:worksheet/ns:dimension", nsManager)

            usedRange = dimensionNode.Attributes("ref").Value

            ExcelXmlPacket.Dispose()
        Catch ex As Exception
            errorMessage = ex.Message
        End Try

        If Not String.IsNullOrEmpty(errorMessage) Then
            Throw New SystemException(errorMessage)
        End If

        Return usedRange
    End Function

    'returns the last used row in an excel
    Public Shared Function GetLastUsedRow(excelFullFilePath As String, sheetName As String) As Integer
        Dim errorMessage As String = String.Empty
        Dim lastRow As Integer = -1

        Try
            ' Instantiate custom xml class
            Dim ExcelXmlPacket As New ExcelXmlPacket(excelFullFilePath)

            ' Get sheet data and namespace URI
            Dim sheetData As String = ExcelXmlPacket.WorksheetData(sheetName)("data").ToString()
            Dim namespaceUri As String = ExcelXmlPacket.SharedStringsData("namespaceuri").ToString()

            ' Load XML document
            Dim doc As New XmlDocument()
            doc.LoadXml(sheetData)

            ' Create a namespace manager
            Dim nsManager As New XmlNamespaceManager(doc.NameTable)
            nsManager.AddNamespace("ns", namespaceUri)

            ' Select all row nodes under sheetData
            Dim rowNodes As XmlNodeList = doc.SelectNodes("//ns:worksheet/ns:sheetData/ns:row", nsManager)

            ' Find the last row with non-empty cells by iterating through row nodes
            For Each rowNode As XmlNode In rowNodes
                Dim rowNumber As Integer = Integer.Parse(rowNode.Attributes("r").Value)
                Dim cellNodes As XmlNodeList = rowNode.SelectNodes("ns:c", nsManager)

                Dim rowHasData As Boolean = False
                For Each cellNode As XmlNode In cellNodes
                    Dim cellValueNode As XmlNode = cellNode.SelectSingleNode("ns:v", nsManager)
                    If cellValueNode IsNot Nothing AndAlso Not String.IsNullOrEmpty(cellValueNode.InnerText) Then
                        rowHasData = True
                        Exit For
                    End If
                Next

                If rowHasData AndAlso rowNumber > lastRow Then
                    lastRow = rowNumber
                End If
            Next

            ExcelXmlPacket.Dispose()
        Catch ex As Exception
            errorMessage = ex.Message
        End Try

        If Not String.IsNullOrEmpty(errorMessage) Then
            Throw New SystemException(errorMessage)
        End If

        Return lastRow
    End Function
    'returns the last used column in an excel
    Public Shared Function GetLastUsedColumn(excelFullFilePath As String, sheetName As String) As Integer
        Dim errorMessage As String = String.Empty
        Dim lastUsedColumn As Integer = 0

        Try
            ' Instantiate custom xml class
            Dim ExcelXmlPacket As New ExcelXmlPacket(excelFullFilePath)

            ' Get sheet data and namespace URI
            Dim sheetData As String = ExcelXmlPacket.WorksheetData(sheetName)("data").ToString()
            Dim namespaceUri As String = ExcelXmlPacket.SharedStringsData("namespaceuri").ToString()

            ' Load XML document
            Dim doc As New XmlDocument()
            doc.LoadXml(sheetData)

            ' Create a namespace manager
            Dim nsManager As New XmlNamespaceManager(doc.NameTable)
            nsManager.AddNamespace("ns", namespaceUri)

            ' Find all cell references
            Dim cellReferences As New List(Of String)()

            For Each cellNode As XmlNode In doc.SelectNodes("//a:c", nsManager)
                Dim cellValueNode As XmlNode = cellNode.SelectSingleNode("a:v", nsManager)
                If cellValueNode IsNot Nothing AndAlso Not String.IsNullOrEmpty(cellValueNode.InnerText) Then
                    Dim cellRef As String = cellNode.Attributes("r").Value
                    cellReferences.Add(cellRef)
                End If
            Next

            ' Find the last used column
            Dim columnsWithData As New HashSet(Of Integer)

            For Each cellRef In cellReferences
                Dim columnName As String = New String(cellRef.Where(AddressOf Char.IsLetter).ToArray())
                Dim columnNumber As Integer = Reusables.ColumnLetterToIndex(columnName)
                columnsWithData.Add(columnNumber)
            Next

            ' Check from the highest column downwards to find the last used column with data
            If columnsWithData.Count > 0 Then
                lastUsedColumn = columnsWithData.Max()
            End If

            ExcelXmlPacket.Dispose()
        Catch ex As Exception
            errorMessage = ex.Message
        End Try

        If Not String.IsNullOrEmpty(errorMessage) Then
            Throw New SystemException(errorMessage)
        End If

        Return lastUsedColumn
    End Function
    'returns all sheets in an excel
    Public Shared Function GetExcelSheetNames(excelFullFilePath As String) As DataTable
        Dim errorMessage As String = String.Empty
        Dim sheets As DataTable = Nothing

        Try
            ' Instantiate custom xml class
            Dim ExcelXmlPacket As New ExcelXmlPacket(excelFullFilePath)

            ' Get Excel sheet names
            Dim sheetNames As String() = ExcelXmlPacket.ExcelSheetNames

            If sheetNames IsNot Nothing AndAlso sheetNames.Length > 0 Then
                sheets = New DataTable
                sheets.Columns.Add("Sheet_Name", GetType(String))

                For Each sheetName As String In sheetNames
                    sheets.Rows.Add({sheetName})
                Next
            End If

            ExcelXmlPacket.Dispose()
        Catch ex As Exception
            errorMessage = ex.Message
        End Try

        If Not String.IsNullOrEmpty(errorMessage) Then
            Throw New SystemException(errorMessage)
        End If

        Return sheets
    End Function
    'gets cell value
    Public Shared Function GetCellValue(excelFullFilePath As String, sheetName As String, cellAddress As String) As String
        Dim cellValue As String = String.Empty
        Dim errorMessage As String = String.Empty

        Try
            ' Instantiate custom XML class
            Dim ExcelXmlPacket As New ExcelXmlPacket(excelFullFilePath)

            ' Throw error if sheet not present
            If Not ExcelXmlPacket.WorksheetData.Keys().Contains(sheetName) Then
                Throw New SystemException("Sheet not found: [" & sheetName & "]")
            End If

            ' Get all XML data that make up the data found in the given Excel sheet
            Dim SheetData As String = ExcelXmlPacket.WorksheetData(sheetName)("data").ToString()
            Dim sharedStringData As String = ExcelXmlPacket.SharedStringsData("data").ToString()
            Dim nameSpaceUri As String = ExcelXmlPacket.SharedStringsData("namespaceuri").ToString()

            ' Process the shared string data which includes all string data that make up the sheet
            Dim sharedStringsMapper As New Dictionary(Of Integer, String)
            Dim sharedStringsXml As New XmlDocument()
            sharedStringsXml.LoadXml(sharedStringData)

            ' Create a namespace manager for shared string
            Dim nsManager As New XmlNamespaceManager(sharedStringsXml.NameTable)
            nsManager.AddNamespace("ns", nameSpaceUri)

            ' Select nodes with the namespace prefix in the shared strings 
            Dim nodesOfSharedStrings As XmlNodeList = sharedStringsXml.SelectNodes("//ns:t", nsManager)

            ' Process the selected nodes into the shared strings Mapper: map the data on the shared strings file to the sheet.xml file found in worksheets
            Dim nodeIndex As Integer = 0
            For Each node As XmlNode In nodesOfSharedStrings
                Dim stringValue As String = node.InnerText
                ' Enter shared string into shared string mapper
                sharedStringsMapper(nodeIndex) = stringValue
                ' Increment node index
                nodeIndex += 1
            Next

            ' Process the sheet data which includes all non-string data that make up the sheet
            Dim sheetXml As New XmlDocument()
            sheetXml.LoadXml(SheetData)

            ' Select the specified cell node
            Dim cellNode As XmlNode = sheetXml.SelectSingleNode("//ns:c[@r='" & cellAddress & "']", nsManager)

            If cellNode IsNot Nothing Then
                ' Trying to get cell value
                Dim dataType As String = cellNode.Attributes("t")?.Value
                Dim nd As String = cellNode.SelectSingleNode("ns:v", nsManager)?.InnerText
                If Not String.IsNullOrEmpty(nd) Then
                    If dataType = "s" Then
                        ' "s" represents string, then get actual value from shared strings dictionary
                        cellValue = sharedStringsMapper(CInt(nd))
                    Else
                        ' Cell value might be in some other format
                        Dim targetFmtIndex As String = cellNode.Attributes("s")?.Value
                        Try
                            Dim fmt As String = ExcelXmlPacket.CellFormatTypes(CInt(targetFmtIndex))

                            Select Case fmt
                                Case "Double"
                                    cellValue = CDbl(nd).ToString()
                                Case "General"
                                    cellValue = nd
                                Case "Unknown"
                                    cellValue = nd
                                Case Else
                                    cellValue = DateTime.FromOADate(nd).ToString(fmt)
                                    cellValue = DateTime.ParseExact(cellValue, fmt, Nothing).ToString("dd/MM/yyyy HH:mm:ss")
                            End Select
                        Catch
                            cellValue = nd
                        End Try
                    End If
                Else
                    ' Cell value is found as blank
                    cellValue = ""
                End If
            Else
                ' Cell not found, return an empty string
                cellValue = ""
            End If

            ExcelXmlPacket.Dispose()

        Catch ex As Exception
            errorMessage = ex.Message
        End Try

        If Not String.IsNullOrEmpty(errorMessage) Then
            Throw New SystemException(errorMessage)
        End If

        Return cellValue
    End Function


    Public Shared Sub WriteTableToNewExcel(dt As DataTable, excelFilePath As String, sheetName As String)
        Try
            ' Create a new XML document for the worksheet
            Dim worksheetXml As New XmlDocument()

            ' Create the worksheet root element
            Dim worksheetElement As XmlElement = worksheetXml.CreateElement("worksheet", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
            worksheetXml.AppendChild(worksheetElement)

            ' Create the sheetData element
            Dim sheetDataElement As XmlElement = worksheetXml.CreateElement("sheetData", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
            worksheetElement.AppendChild(sheetDataElement)

            ' Populate the sheetData with the DataTable
            For Each row As DataRow In dt.Rows
                Dim rowElement As XmlElement = worksheetXml.CreateElement("row", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
                rowElement.SetAttribute("r", (dt.Rows.IndexOf(row) + 1).ToString())

                For Each col As DataColumn In dt.Columns
                    Dim cellElement As XmlElement = worksheetXml.CreateElement("c", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
                    cellElement.SetAttribute("r", Reusables.ColumnIndexToLetter(col.Ordinal) & (dt.Rows.IndexOf(row) + 1))

                    Dim vElement As XmlElement = worksheetXml.CreateElement("v", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
                    vElement.InnerText = row(col).ToString()
                    cellElement.AppendChild(vElement)

                    rowElement.AppendChild(cellElement)
                Next
                sheetDataElement.AppendChild(rowElement)
            Next

            ' Save the XML document
            worksheetXml.Save(excelFilePath)

            ' Now update the file extension to .xlsx
            Dim newExcelFilePath As String = Path.ChangeExtension(excelFilePath, ".xlsx")
            File.Move(excelFilePath, newExcelFilePath)

        Catch ex As Exception
            ' Handle any exceptions
            Throw New SystemException("Error writing DataTable to Excel: " & ex.Message)
        End Try
    End Sub

End Class
