Imports System.Data
Imports System.IO
Imports System.IO.Compression
Imports System.Text.RegularExpressions
Imports System.Xml


Public Class ExcelXmlPacket
	Private _excelFilePath As String
	Private _generatedXmlFolderPath As String
	Private _xlFolderPath As String
	Private _worksheetsPath As String
	Private _sharedStringsPath As String
	Private _sharedStringsData As New Dictionary(Of String, Object)
	Private _worksheetData As Dictionary(Of String, Dictionary(Of String, Object))
	Private _cellFormatTypes As New Dictionary(Of Integer, Object)
	Private _excelSheetPaths As String()
	Private _excelSheetNames As String()
	Private _stylesPath As String
	Private _workbookXmlPath As String
	Private _xmlContentArchivedPath As String
	Private _relPath As String



	Public ReadOnly Property xmlContentArchivedPath As String
		Get
			Return _xmlContentArchivedPath
		End Get
	End Property

	Public ReadOnly Property ExcelFilePath As String
		Get
			Return _excelFilePath
		End Get
	End Property



	Public ReadOnly Property WorksheetData As Dictionary(Of String, Dictionary(Of String, Object))
		Get
			Return _worksheetData
		End Get
	End Property

	Public ReadOnly Property CellFormatTypes As Dictionary(Of Integer, Object)
		Get
			Return _cellFormatTypes
		End Get
	End Property

	Public ReadOnly Property SharedStringsData As Dictionary(Of String, Object)
		Get
			Return _sharedStringsData 
		End Get
	End Property

	Public ReadOnly Property ExcelSheetNames As String()
		Get
			Return _excelSheetNames
		End Get
	End Property
	Public Sub New(excelFilePath As String)
		'set expected file paths here
		_excelFilePath = excelFilePath
		_generatedXmlFolderPath = Path.Combine(Path.GetPathRoot(Environment.SystemDirectory), "Automation", "XML_Processor")
		_xlFolderPath = Path.Combine(_generatedXmlFolderPath, "xl")
		_relPath = Path.Combine(_generatedXmlFolderPath, "xl", "_rels", "workbook.xml.rels")
		_stylesPath = Path.Combine(_xlFolderPath, "styles.xml")
		_workbookXmlPath = Path.Combine(_xlFolderPath, "workbook.xml")
		_worksheetsPath = Path.Combine(_xlFolderPath, "worksheets")
		_sharedStringsPath = Path.Combine(_xlFolderPath, "sharedStrings.xml")
		'Run method to generate xml files into generated folder path
		_xmlContentArchivedPath = GenerateExcelXml()
		'execute method to get work sheet data which includes sheet id, sheet name , sheet path and sheet data in a single variable as a dictionary
		_worksheetData = GetWorkSheetData(_relPath, _workbookXmlPath, _xlFolderPath)
		'store all the excel sheet names
		_excelSheetNames = _worksheetData.Keys().ToArray()
		'store the shared strings 
		_sharedStringsData("data") = File.ReadAllText(_sharedStringsPath)
		_sharedStringsData("namespaceuri") = GetNameSpaceUri(_sharedStringsPath)
		'get the cell format types
		_cellFormatTypes = GetCellFormatTypes(_stylesPath)
	End Sub
	Private Function GenerateExcelXml() As String
		DeleteAFolder(_generatedXmlFolderPath)
		CreateAFolder(_generatedXmlFolderPath)
		Dim zipfilePath As String = CreateACopyOfAFileByChangingExtension(_excelFilePath, _generatedXmlFolderPath)
		UnzipAFile(zipfilePath, _generatedXmlFolderPath)
		Return zipfilePath
	End Function
	Private Sub CreateAFolder(folderPath As String)
		If Not Directory.Exists(folderPath) Then
			Directory.CreateDirectory(folderPath)
		End If
	End Sub

	'deletes a folder 
	Private Sub DeleteAFolder(folderPath As String)
		If Directory.Exists(folderPath) Then
			Directory.Delete(folderPath, True)
		End If
	End Sub

	'creates a copy of an excel file into a target location and changing its extension from .xlsx to .zip
	Private Function CreateACopyOfAFileByChangingExtension(filePath As String, folderPath As String) As String
		If File.Exists(filePath) AndAlso Directory.Exists(folderPath) Then
			Dim newFilePath As String = Path.Combine(folderPath, Path.GetFileNameWithoutExtension(filePath) + ".zip")
			File.Copy(filePath, Path.Combine(folderPath, newFilePath))
			Return newFilePath
		Else
			Throw New SystemException("intended file path and destination folder must be valid")
		End If
	End Function

	'unzips a given zip into a desired location
	Private Sub UnzipAFile(zipFilePath As String, extractionFolderPath As String)
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
	'returns the name space url of an xml sheet
	Private Function GetNameSpaceUri(xmlFilePath As String) As String
		Dim xmlDoc As New XmlDocument()
		xmlDoc.Load(xmlFilePath)
		' Get the namespace URI
		Dim namespaceUri As String = xmlDoc.DocumentElement.GetAttribute("xmlns")
		Return namespaceUri
	End Function



	'This function when called gets the sheet id, sheet name , sheet path and sheet data in a single variable as a dictionary
	Private Function GetWorkSheetData(relFilePath As String, workbookXmlFilePath As String, xlFolderPath As String) As Dictionary(Of String, Dictionary(Of String, Object))
		Dim RIdToTarget As New Dictionary(Of String, Object)
		Dim xmlRelDoc As New XmlDocument()
		xmlRelDoc.Load(relFilePath)
		' Load Name space
		Dim nsManager As New XmlNamespaceManager(xmlRelDoc.NameTable)
		nsManager.AddNamespace("ns", GetNameSpaceUri(relFilePath))
		'For Each relNode As XmlNode In xmlRelDoc.SelectNodes("//ns:Relationship", nsManager)
		For Each rNode As XmlNode In xmlRelDoc.SelectNodes("//ns:Relationship", nsManager)
			Dim Target As String = rNode.Attributes("Target").Value
			Dim SheetId As String = rNode.Attributes("Id").Value
			RIdToTarget.Add(SheetId, Target)
		Next
		Dim xmlSheetData As New Dictionary(Of String, Dictionary(Of String, Object))
		Dim xmlWorkbookDoc As New XmlDocument()
		xmlWorkbookDoc.Load(workbookXmlFilePath)
		' Load Name space
		nsManager = New XmlNamespaceManager(xmlWorkbookDoc.NameTable)
		nsManager.AddNamespace("ns", GetNameSpaceUri(workbookXmlFilePath))
		For Each sheetNode As XmlNode In xmlWorkbookDoc.SelectNodes("//ns:sheets", nsManager)
			For Each sNode As XmlNode In sheetNode.SelectNodes("ns:sheet", nsManager)
				Try
					Dim actualSheetName As String = sNode.Attributes("name").Value
					Dim SheetId As String = sNode.Attributes("sheetId").Value
					Dim RId As String = sNode.Attributes("r:id").Value
					Dim sheetFullFilePath As String = Path.Combine(xlFolderPath, RIdToTarget(RId).ToString())
					Dim SheetData As String = File.ReadAllText(sheetFullFilePath)
					Dim sheetIndex As Integer = CInt(Path.GetFileNameWithoutExtension(RIdToTarget(RId)).ToString().Replace("sheet", String.Empty)) - 1
					Dim namespaceuri As String = GetNameSpaceUri(sheetFullFilePath)
					xmlSheetData.Add(actualSheetName, New Dictionary(Of String, Object) From {{"sheetid", SheetId}, {"r:id", RId}, {"worksheet", RIdToTarget(RId)}, {"data", SheetData}, {"sheetindex", sheetIndex}, {"namespaceuri", namespaceuri}})
				Catch
				End Try
			Next
		Next
		Return xmlSheetData
		' Next
	End Function

	Private Function GetCellFormatTypes(workbookXmlFilePath As String) As Dictionary(Of Integer, Object)
		Dim numFmtIds As New Dictionary(Of Integer, Object)()

		Dim xmlWorkbookDoc As New XmlDocument()
		xmlWorkbookDoc.Load(workbookXmlFilePath)

		' Load Name space
		Dim nsManager As New XmlNamespaceManager(xmlWorkbookDoc.NameTable)
		nsManager.AddNamespace("ns", GetNameSpaceUri(workbookXmlFilePath))

		' Select the cellXfs node
		Dim cellXfsNode As XmlNode = xmlWorkbookDoc.SelectSingleNode("//ns:cellXfs", nsManager)

		If cellXfsNode IsNot Nothing Then
			' Loop through each xf node inside cellXfs
			Dim nodeIndex As Integer = 0
			For Each xfNode As XmlNode In cellXfsNode.SelectNodes("ns:xf", nsManager)
				' Get the numFmtId attribute value
				Dim numFmtId As Integer = 0 ' Default value if attribute is missing or invalid
				If Integer.TryParse(xfNode.Attributes("numFmtId")?.Value, numFmtId) Then
					' Add numFmtId to dictionary
					Dim formatString = ""
					Select Case numFmtId
						Case 0
							formatString = "General"
						Case 1
							formatString = "General"
						Case 2
							formatString = "Double"
						Case 3
							formatString = "Double"
						Case 4
							formatString = "Double"
						Case 9
							formatString = "Double"
						Case 10
							formatString = "Double"
						Case 14
							formatString = "MM-dd-yy"
						Case 15
							formatString = "d-MMM-yy"
						Case 16
							formatString = "d-MMM"
						Case 17
							formatString = "MMM-yy"
						Case 18
							formatString = "h tt"
						Case 19
							formatString = "h:mm tt"
						Case 20
							formatString = "h"
						Case 21
							formatString = "h:mm"
						Case 22
							formatString = "M/d/yy h"
						Case Else
							formatString = "Unknown"
					End Select

					numFmtIds(nodeIndex) = formatString ' Value is not being used here, you can set it to something meaningful if needed
				End If
				nodeIndex = nodeIndex + 1
			Next
		End If

		Return numFmtIds
	End Function

	'dispose
	Public Sub Dispose()
		Dim generatedXmlFolderPath As String = Path.Combine(Path.GetPathRoot(Environment.SystemDirectory), "Automation", "XML_Processor")
		DeleteAFolder(generatedXmlFolderPath)
	End Sub
End Class
