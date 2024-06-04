Imports System.Data
Imports System.IO
Imports System.Xml
Imports ExcelXmlManager

'here we implement some more methods that have to do with reading the excel. now rite action takes here
Public Class ExcelXmlAction
    'Reads an excel to a datatable
    Public Shared Function ReadExcelToTable(excelFullFilePath As String, sheetName As String, hasHeader As Boolean) As DataTable
        Dim dt As DataTable = Nothing
        Dim errorMessage As String = String.Empty

        Try
            ' Instantiate custom xml class
            Dim exl As New ExcelXmlPacket(excelFullFilePath)

            ' Throw error if sheet not present
            If Not exl.WorksheetConfig.Keys().Contains(sheetName) Then
                Throw New SystemException("Sheet not found: [" & sheetName & "]")
            End If

            ' Get all xml data that are used to make up the data found in the given excel sheet
            Dim SheetData As String = exl.WorksheetConfig(sheetName)("data").ToString()
            Dim sharedStringData As String = exl.SharedStringsConfig("data").ToString()

            ' Process the shared string data which include all string data that make up the sheet
            Dim sharedStringsMapper As New Dictionary(Of Integer, String)
            Dim sharedStringsXml As New XmlDocument()
            sharedStringsXml.LoadXml(sharedStringData)

            ' Create a namespace manager for shared string
            Dim nsManager As New XmlNamespaceManager(sharedStringsXml.NameTable)
            exl.AddSharedStringXmlNameSpace(nsManager)

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
                                Dim fmt As String = exl.CellFormatTypesConfig(CInt(targetFmtIndex))

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

            exl.Dispose()

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
            Dim exl As New ExcelXmlPacket(excelFullFilePath)

            ' Return all the available sheet indices
            Dim indices As Integer() = exl.WorksheetConfig.Values().Select(Function(v) CInt(v("sheetindex").ToString())).ToArray()

            ' Throw error if sheet not present
            If Not indices.Contains(sheetIndex) Then
                Throw New SystemException("No sheet found at index: [" & sheetIndex.ToString() & "]")
            End If

            ' Get the sheet name related to the given index
            Dim sheetName As String = exl.WorksheetConfig.First(Function(p) Convert.ToInt32(p.Value("sheetindex")) = sheetIndex).Key()

            ' Get all xml data that are used to make up the data found in the given excel sheet
            ' Get all xml data that are used to make up the data found in the given excel sheet
            Dim SheetData As String = exl.WorksheetConfig(sheetName)("data").ToString()
            Dim sharedStringData As String = exl.SharedStringsConfig("data").ToString()

            ' Process the shared string data which include all string data that make up the sheet
            Dim sharedStringsMapper As New Dictionary(Of Integer, String)
            Dim sharedStringsXml As New XmlDocument()
            sharedStringsXml.LoadXml(sharedStringData)

            ' Create a namespace manager for shared string
            Dim nsManager As New XmlNamespaceManager(sharedStringsXml.NameTable)
            exl.AddSharedStringXmlNameSpace(nsManager)

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
                                Dim fmt As String = exl.CellFormatTypesConfig(CInt(targetFmtIndex))

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

            exl.Dispose()

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
            Dim exl As New ExcelXmlPacket(excelFullFilePath)

            ' Return all the available sheet indices
            Dim indices As Integer() = exl.WorksheetConfig.Values().Select(Function(v) CInt(v("sheetindex").ToString())).ToArray()

            ' Throw error if sheet not present
            If Not indices.Contains(sheetIndex) Then
                Throw New SystemException("No sheet found at index: [" & sheetIndex.ToString() & "]")
            End If

            ' Get the sheet name related to the given index
            sheetName = exl.WorksheetConfig.First(Function(p) Convert.ToInt32(p.Value("sheetindex")) = sheetIndex).Key()

            exl.Dispose()
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
            Dim exl As New ExcelXmlPacket(excelFullFilePath)

            ' Throw error if sheet not present
            If Not exl.WorksheetConfig.Keys().Contains(sheetName) Then
                Throw New SystemException("Sheet not found: [" & sheetName & "]")
            End If

            ' Get sheet index
            sheetIndex = CInt(exl.WorksheetConfig(sheetName)("sheetindex"))

            exl.Dispose()
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
            Dim exl As New ExcelXmlPacket(excelFullFilePath)

            ' Throw error if sheet not present
            If Not exl.WorksheetConfig.Keys().Contains(sheetName) Then
                Throw New SystemException("Sheet not found: [" & sheetName & "]")
            End If

            ' Get all xml data that are used to make up the data found in the given excel sheet
            Dim sheetData As String = exl.WorksheetConfig(sheetName)("data").ToString()
            ' Get expected namespace URI
            Dim namespaceUri As String = exl.SharedStringsConfig("namespaceuri").ToString()

            Dim doc As New XmlDocument()
            doc.LoadXml(sheetData)

            ' Create a namespace manager
            Dim nsManager As New XmlNamespaceManager(doc.NameTable)
            exl.AddSheetXmlNameSpace(nsManager)

            ' Select nodes with the namespace prefix in the dimension 
            Dim dimensionNode As XmlNode = doc.SelectSingleNode("//ns:worksheet/ns:dimension", nsManager)

            usedRange = dimensionNode.Attributes("ref").Value

            exl.Dispose()
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
            Dim exl As New ExcelXmlPacket(excelFullFilePath)

            ' Get sheet data and namespace URI
            Dim sheetData As String = exl.WorksheetConfig(sheetName)("data").ToString()
            Dim namespaceUri As String = exl.SharedStringsConfig("namespaceuri").ToString()

            ' Load XML document
            Dim doc As New XmlDocument()
            doc.LoadXml(sheetData)

            ' Create a namespace manager
            Dim nsManager As New XmlNamespaceManager(doc.NameTable)
            exl.AddSheetXmlNameSpace(nsManager)

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

            exl.Dispose()
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
            Dim exl As New ExcelXmlPacket(excelFullFilePath)

            ' Get sheet data and namespace URI
            Dim sheetData As String = exl.WorksheetConfig(sheetName)("data").ToString()

            ' Load XML document
            Dim doc As New XmlDocument()
            doc.LoadXml(sheetData)

            ' Create a namespace manager
            Dim nsManager As New XmlNamespaceManager(doc.NameTable)
            exl.AddSheetXmlNameSpace(nsManager)

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

            exl.Dispose()
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
            Dim excelXmlPacket As New ExcelXmlPacket(excelFullFilePath)

            ' Get Excel sheet names
            Dim sheetNames As String() = excelXmlPacket.sheetNames

            If sheetNames IsNot Nothing AndAlso sheetNames.Length > 0 Then
                sheets = New DataTable
                sheets.Columns.Add("Sheet_Name", GetType(String))

                For Each sheetName As String In sheetNames
                    sheets.Rows.Add({sheetName})
                Next
            End If

            excelXmlPacket.Dispose()
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
            Dim excelXmlPacket As New ExcelXmlPacket(excelFullFilePath)

            ' Throw error if sheet not present
            If Not excelXmlPacket.WorksheetConfig.Keys().Contains(sheetName) Then
                Throw New SystemException("Sheet not found: [" & sheetName & "]")
            End If

            ' Get all XML data that make up the data found in the given Excel sheet
            Dim SheetData As String = excelXmlPacket.WorksheetConfig(sheetName)("data").ToString()
            Dim sharedStringData As String = excelXmlPacket.SharedStringsConfig("data").ToString()

            ' Process the shared string data which includes all string data that make up the sheet
            Dim sharedStringsMapper As New Dictionary(Of Integer, String)
            Dim sharedStringsXml As New XmlDocument()
            sharedStringsXml.LoadXml(sharedStringData)

            ' Create a namespace manager for shared string
            Dim nsManager As New XmlNamespaceManager(sharedStringsXml.NameTable)
            excelXmlPacket.AddSharedStringXmlNameSpace(nsManager)

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
                            Dim fmt As String = excelXmlPacket.CellFormatTypesConfig(CInt(targetFmtIndex))

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

            excelXmlPacket.Dispose()

        Catch ex As Exception
            errorMessage = ex.Message
        End Try

        If Not String.IsNullOrEmpty(errorMessage) Then
            Throw New SystemException(errorMessage)
        End If

        Return cellValue
    End Function

    Public Shared Sub DeleteSheet(excelFullFilePath As String, sheetName As String)
        Try
            Dim exl As New ExcelXmlPacket(excelFullFilePath)
            exl.DeleteSheet(sheetName)
        Catch ex As Exception
            Console.WriteLine($"Error occurred while deleting the sheet: {ex.Message}")
        End Try
    End Sub


    Public Shared Sub CreateSheet(excelFullFilePath As String, sheetName As String)
        Try
            Dim exl As New ExcelXmlPacket(excelFullFilePath)
            ' Throw error if sheet not present
            If exl.WorksheetConfig.Keys().Contains(sheetName) Then
                Throw New SystemException("Sheet with name [" & sheetName & "] already exists")
            End If
            exl.GenerateNewSheet(sheetName)
            ' Update the Excel file
            exl.Dispose()
        Catch ex As Exception
            Console.WriteLine($"Error occurred while creating the sheet: {ex.Message}")
        End Try
    End Sub

    Public Shared Sub WriteToTableToExcel(excelFullFilePath As String, sheetName As String, dt As DataTable)
        If String.IsNullOrEmpty(sheetName) Then
            'no custom sheet then set sheetname as sheet 1
            sheetName = "Sheet1"
        End If

        'If excel file does not exist then create new excel
        If Not File.Exists(excelFullFilePath) Then
            'create new excel 
            ExcelXmlPacket.CreateNewExcel(excelFullFilePath, sheetName)
        End If

        'now initialize the excelXmlPacket
        Dim exl As New ExcelXmlPacket(excelFullFilePath)

        'confirm  sheet exists if sheet does not create sheet
        If exl.WorksheetConfig.Keys().Contains(sheetName) Then
            CreateSheet(excelFullFilePath, sheetName)
        End If

        exl.Dispose()

    End Sub



End Class
