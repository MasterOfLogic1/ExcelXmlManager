Imports System.Data
Imports System.IO
Imports System.IO.Compression
Imports System.Text.RegularExpressions
Imports System.Xml

'--------------------------------------------------------------------- Concept ---------------------------------------------------------------------------------------------
' This Library facilitates reading, writing, and managing Excel files as XMLs.
' The core idea is to convert the Excel file (.xlsx) to a zip file by simply changing the file extension from '.xlsx' to '.zip'.
' The zip file is then unzipped, revealing the XML files that constitute the Excel file.
' These XML files are manipulated to perform read or write operations.
' A standard Excel file renamed to a zip extension and unzipped will typically contain the following directories and files:
'   _rels
'       .rels
'   docProps
'       app.xml
'       core.xml
'   xl
'       _rels
'           workbook.xml.rels
'       theme
'           theme1.xml
'       worksheets
'           sheet1.xml
'           sheet2.xml
'           ...
'       styles.xml
'       workbook.xml
'       sharedStrings.xml (if the workbook contains any shared strings)
'       calcChain.xml (if the workbook contains formula calculations)
'   [Content_Types].xml
'
' Each of these files and directories plays a specific role in defining the structure and content of the Excel workbook.
' By manipulating these XML files, you can read from and write to the Excel workbook programmatically.
'--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

'Please note that while some read methods can be found here and in the ExcelXmlAction class, it's important to know that all write and update methods are implemented here.
'These include actions such as deleting a sheet, creating a new sheet, etc.

Public Class ExcelXmlPacket
	Private _excelFilePath As String ' Holds the path of the Excel file
	Private _processingFolderPath As String ' Path for the folder where processing takes place
	Private _readXmlFolderPath As String ' Path for the folder where XML content will be read from the unzipped Excel file
	Private _writeExcelFolderPath As String ' When a manipulated xml reverted back to an excel it would be saved here
	Private _xlFolderPath As String ' Path to the 'xl' folder within the unzipped Excel file structure
	Private _worksheetsPath As String ' Path to the 'worksheets' folder within the 'xl' directory
	Private _sharedStringsPath As String ' Path to the 'sharedStrings.xml' file within the 'xl' directory
	Private sharedStringsData As New Dictionary(Of String, Object) ' Dictionary to hold all info found in sharedstrings.xml
	Private worksheetData As Dictionary(Of String, Dictionary(Of String, Object)) ' Dictionary to hold worksheet data, including sheet ID, sheet name, sheet path, and sheet content
	Private cellFormatTypes As New Dictionary(Of Integer, Object) ' Dictionary to hold cell format types extracted from the 'styles.xml' file
	Private _excelSheetPaths As String() ' Array to hold paths of Excel sheets within the workbook
	Private excelSheetNames As String() ' Array to hold names of Excel sheets within the workbook
	Private _stylesXmlPath As String ' Path to the 'styles.xml' file within the 'xl' directory
	Private _workbookXmlPath As String ' Path to the 'workbook.xml' file within the 'xl' directory
	Private _xmlContentArchivedPath As String ' Path to the archived XML content (zipped Excel file)
	Private _relXmlPath As String ' Path to the 'workbook.xml.rels' file within the '_rels' directory inside the 'xl' folder



	Public ReadOnly Property WorksheetConfig As Dictionary(Of String, Dictionary(Of String, Object))
		Get
			Return worksheetData
		End Get
	End Property

	Public ReadOnly Property CellFormatTypesConfig As Dictionary(Of Integer, Object)
		Get
			Return cellFormatTypes
		End Get
	End Property

	Public ReadOnly Property SharedStringsConfig As Dictionary(Of String, Object)
		Get
			Return sharedStringsData
		End Get
	End Property

	Public ReadOnly Property sheetNames As String()
		Get
			Return excelSheetNames
		End Get
	End Property


	Public Sub New(excelFilePath As String)
		'set expected file paths here which is received as soon as this class is intialized:
		_excelFilePath = excelFilePath
		' set Path where processing takes place:
		_processingFolderPath = Path.Combine(Path.GetPathRoot(Environment.SystemDirectory), "Automation", "XML_Processor")
		' set Path where reading and manipulation of xml files would take place .this is same folder were xml would be unzipped into:
		_readXmlFolderPath = Path.Combine(Path.GetPathRoot(Environment.SystemDirectory), "Automation", "XML_Processor", "Read")
		'set path to write to revert xml into excel file
		_writeExcelFolderPath = Path.Combine(Path.GetPathRoot(Environment.SystemDirectory), "Automation", "XML_Processor", "write")
		' set xl folder path as expected 
		_xlFolderPath = Path.Combine(_readXmlFolderPath, "xl")
		'defines file path to the xml found in Automation\XML_Processor\Read\xl\_rels\workbook.xml.rels:
		_relXmlPath = Path.Combine(_xlFolderPath, "_rels", "workbook.xml.rels")
		'defines file path to the xml found in Automation\XML_Processor\Read\xl\styles.xml:
		_stylesXmlPath = Path.Combine(_xlFolderPath, "styles.xml")
		'defines file path to the xml found in Automation\XML_Processor\Read\xl\workbook.xml :
		_workbookXmlPath = Path.Combine(_xlFolderPath, "workbook.xml")
		'defines folder path to the Automation\XML_Processor\Read\xl\worksheets :
		_worksheetsPath = Path.Combine(_xlFolderPath, "worksheets")
		'defines file path to the Automation\XML_Processor\Read\xl\sharedString.xml :
		_sharedStringsPath = Path.Combine(_xlFolderPath, "sharedStrings.xml")
		'Run method to generate xml files into Automation\XML_Processor\Read 
		_xmlContentArchivedPath = GenerateExcelXml()
		'execute method to get work sheet data which includes sheet id, sheet name , sheet path and sheet data in a single variable as a dictionary
		worksheetData = GetWorkSheetData(_relXmlPath, _workbookXmlPath, _xlFolderPath)
		'store all the excel sheet names
		excelSheetNames = worksheetData.Keys().ToArray()
		'store the shared strings 
		sharedStringsData("data") = File.ReadAllText(_sharedStringsPath)
		'get the cell format types from Automation\XML_Processor\Read\xl\styles.xml:
		cellFormatTypes = GetCellFormatTypes(_stylesXmlPath)
	End Sub

	Private Function GenerateExcelXml() As String
		'Delete the processor folder path (Automation\XML_Processor)
		Reusables.DeleteAFolder(_processingFolderPath)
		'create the read folder path inside of the processing folder Automation\XML_Processor\Read  (This also automatically creates the processing folder)
		Reusables.CreateAFolder(_readXmlFolderPath)
		'Copy the target excel file into the [Automation\XML_Processor\Read] with a .zip extension
		Dim zipfilePath As String = Reusables.CreateACopyOfAFileByChangingExtension(_excelFilePath, _readXmlFolderPath)
		'unzip the copied .zip file
		Reusables.UnzipAFile(zipfilePath, _readXmlFolderPath)
		'relocate the zipped file to [Automation\XML_Processor] avoid issues
		File.Move(zipfilePath, Path.Combine(_processingFolderPath, Path.GetFileName(zipfilePath)))
		'set the new file path of the zip file for audit which would be [Automation\XML_Processor\WhateverTheExcelNameWas.zip]
		zipfilePath = Path.Combine(_processingFolderPath, Path.GetFileName(zipfilePath))
		Return zipfilePath
	End Function


	'The functions adds the required namespace for specific xml found in the unzipped file. the name matches the xml they speak to
	Public Sub AddWorbookXmlNameSpace(nsManager As XmlNamespaceManager)
		nsManager.AddNamespace("ns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
		nsManager.AddNamespace("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006")
		nsManager.AddNamespace("x15", "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main")
		nsManager.AddNamespace("xr", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision")
		nsManager.AddNamespace("xr6", "http://schemas.microsoft.com/office/spreadsheetml/2016/revision6")
		nsManager.AddNamespace("xr10", "http://schemas.microsoft.com/office/spreadsheetml/2016/revision10")
		nsManager.AddNamespace("xr2", "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2")
	End Sub


	Public Sub AddSheetXmlNameSpace(nsManager As XmlNamespaceManager)
		nsManager.AddNamespace("default", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
		nsManager.AddNamespace("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
		nsManager.AddNamespace("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006")
		nsManager.AddNamespace("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac")
		nsManager.AddNamespace("xr", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision")
		nsManager.AddNamespace("xr2", "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2")
		nsManager.AddNamespace("xr3", "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3")
	End Sub

	Public Sub AddRelXmlNameSpace(nsManager As XmlNamespaceManager)
		nsManager.AddNamespace("ns", "http://schemas.openxmlformats.org/package/2006/relationships")
		nsManager.AddNamespace("r", "http://schemas.openxmlformats.org/package/2006/relationships")
	End Sub

	Public Sub AddSharedStringXmlNameSpace(nsManager As XmlNamespaceManager)
		nsManager.AddNamespace("ns", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
		nsManager.AddNamespace("default", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
	End Sub

	Public Sub AddDefaultNameSpace(nsManager As XmlNamespaceManager)
		nsManager.AddNamespace("default", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
	End Sub


	'This function when called gets the sheet id, sheet name , sheet path and sheet data in a single variable as a dictionary
	Private Function GetWorkSheetData(relFilePath As String, workbookXmlFilePath As String, xlFolderPath As String) As Dictionary(Of String, Dictionary(Of String, Object))
		Dim RIdToTarget As New Dictionary(Of String, Object)
		Dim xmlRelDoc As New XmlDocument()
		xmlRelDoc.Load(relFilePath)
		' Load Name space
		Dim nsManager As New XmlNamespaceManager(xmlRelDoc.NameTable)
		AddRelXmlNameSpace(nsManager)
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
		AddWorbookXmlNameSpace(nsManager)
		For Each sheetNode As XmlNode In xmlWorkbookDoc.SelectNodes("//ns:sheets", nsManager)
			For Each sNode As XmlNode In sheetNode.SelectNodes("ns:sheet", nsManager)
				Try
					Dim actualSheetName As String = sNode.Attributes("name").Value
					Dim SheetId As String = sNode.Attributes("sheetId").Value
					Dim RId As String = sNode.Attributes("r:id").Value
					Dim sheetFullFilePath As String = Path.Combine(xlFolderPath, RIdToTarget(RId).ToString())
					Dim SheetData As String = File.ReadAllText(sheetFullFilePath)
					Dim sheetIndex As Integer = CInt(Path.GetFileNameWithoutExtension(RIdToTarget(RId)).ToString().Replace("sheet", String.Empty)) - 1
					xmlSheetData.Add(actualSheetName, New Dictionary(Of String, Object) From {{"sheetid", SheetId}, {"r:id", RId}, {"worksheet", RIdToTarget(RId)}, {"data", SheetData}, {"sheetindex", sheetIndex}, {"namespaceuri", Nothing}})
				Catch
				End Try
			Next
		Next
		Return xmlSheetData
		' Next
	End Function

	'This function genrates a dictionary which helps to identify non string datatypes: this is to ensure that dates and numbers are reported efficeintly 
	Private Function GetCellFormatTypes(workbookXmlFilePath As String) As Dictionary(Of Integer, Object)
		Dim numFmtIds As New Dictionary(Of Integer, Object)()

		Dim xmlWorkbookDoc As New XmlDocument()
		xmlWorkbookDoc.Load(workbookXmlFilePath)

		' Load Name space
		Dim nsManager As New XmlNamespaceManager(xmlWorkbookDoc.NameTable)
		AddWorbookXmlNameSpace(nsManager)

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

	'This function would generate a new sheet in the traget excel
	Public Sub GenerateNewSheet(sheetName As String)
		'Generate a new sheetId - remember compensate for sharedstrings and styles (increment by +3)
		Dim sheetId As Integer = worksheetData.Count + 4
		'using new sheet id set sheet path
		Dim sheetXmlPath As String = Path.Combine(_worksheetsPath, "sheet" + sheetId.ToString() + ".xml")
		'now set the sheet rId which would be placed in worbook.xml
		Dim newSheetId As String = $"rId{sheetId}"

		UpdateWorkbookXml(sheetName, newSheetId, sheetId)
		UpdateNewSheetXml(sheetXmlPath)
		UpdateRelXml(sheetXmlPath, newSheetId, sheetId)
		'----now update worsheet dictionary -------
		Dim newSheetData As New Dictionary(Of String, Object) From {
		{"sheetid", sheetId},
		{"r:id", newSheetId},
		{"worksheet", "worksheets/sheet" + sheetId.ToString() + ".xml"},
		{"data", File.ReadAllText(sheetXmlPath)},
		{"sheetindex", sheetId - 1},
		{"namespaceuri", ""}
		}
		worksheetData.Add(sheetName, newSheetData)
		UpdateAppXml(sheetName)
		AddSheetToContentTypes(sheetId)
		Reusables.RevertXmlParentFolderToExcelFile(_excelFilePath, _readXmlFolderPath)
	End Sub

	'This function would delete a given sheet in the target excel
	Public Sub DeleteSheet(sheetName As String)
		Dim sheetToDelete As String = Nothing
		For Each sheet In WorksheetData
			If sheet.Key = sheetName Then
				sheetToDelete = sheet.Value("r:id").ToString()
				Exit For
			End If
		Next

		If sheetToDelete IsNot Nothing Then
			' Remove the sheet from worksheet data
			worksheetData.Remove(sheetName)

			' Remove the corresponding relationship entry
			Dim xmlRelDoc As New XmlDocument()
			xmlRelDoc.Load(_relXmlPath)
			Dim nsManager As New XmlNamespaceManager(xmlRelDoc.NameTable)
			AddRelXmlNameSpace(nsManager)
			Dim relNode As XmlNode = xmlRelDoc.SelectSingleNode($"//ns:Relationship[@Id='{sheetToDelete}']", nsManager)
			If relNode IsNot Nothing Then
				Dim target As String = relNode.Attributes("Target").Value
				Dim sheetFilePath As String = Path.Combine(_xlFolderPath, target)
				If File.Exists(sheetFilePath) Then
					File.Delete(sheetFilePath)
				End If
				relNode.ParentNode.RemoveChild(relNode)
				xmlRelDoc.Save(_relXmlPath)

				' Remove the sheet from workbook XML
				Dim xmlWorkbookDoc As New XmlDocument()
				xmlWorkbookDoc.Load(_workbookXmlPath)
				nsManager = New XmlNamespaceManager(xmlWorkbookDoc.NameTable)
				AddWorbookXmlNameSpace(nsManager)
				Dim sheetNode As XmlNode = xmlWorkbookDoc.SelectSingleNode($"//ns:sheets/ns:sheet[@r:id='{sheetToDelete}']", nsManager)
				If sheetNode IsNot Nothing Then
					sheetNode.ParentNode.RemoveChild(sheetNode)
					xmlWorkbookDoc.Save(_workbookXmlPath)
				End If
			End If
		Else
			Console.WriteLine("Sheet not found.")
		End If
		Reusables.RevertXmlParentFolderToExcelFile(_excelFilePath, _readXmlFolderPath)
	End Sub





	'---------------------when a new sheet is created the following subs would be called :-------------------
	Private Sub UpdateAppXml(sheetName As String)
		Dim appXmlPath As String = Path.Combine(_readXmlFolderPath, "docProps", "app.xml")
		Dim xmlAppDoc As New XmlDocument()
		xmlAppDoc.Load(appXmlPath)

		' Manage namespaces
		Dim nsManager As New XmlNamespaceManager(xmlAppDoc.NameTable)
		nsManager.AddNamespace("vt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes")
		nsManager.AddNamespace("ep", "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties")

		' Update HeadingPairs
		Dim headingPairsNode As XmlNode = xmlAppDoc.SelectSingleNode("//ep:HeadingPairs/vt:vector[vt:variant/vt:lpstr='Worksheets']/vt:variant[2]/vt:i4", nsManager)
		If headingPairsNode IsNot Nothing Then
			headingPairsNode.InnerText = worksheetData.Count.ToString()
		Else
			Console.WriteLine("HeadingPairs node not found.")
		End If

		' Update TitlesOfParts
		Dim titlesOfPartsVectorNode As XmlElement = TryCast(xmlAppDoc.SelectSingleNode("//ep:TitlesOfParts/vt:vector", nsManager), XmlElement)
		If titlesOfPartsVectorNode IsNot Nothing Then
			' Update size attribute for TitlesOfParts vector
			Dim currentSize As Integer = titlesOfPartsVectorNode.ChildNodes.Count
			titlesOfPartsVectorNode.SetAttribute("size", (currentSize + 1).ToString())

			' Add new sheet name
			Dim lpstrNode As XmlElement = xmlAppDoc.CreateElement("vt:lpstr", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes")
			lpstrNode.InnerText = sheetName
			titlesOfPartsVectorNode.AppendChild(lpstrNode)
		Else
			Console.WriteLine("TitlesOfParts vector node not found.")
		End If

		' Save the modified app.xml
		xmlAppDoc.Save(appXmlPath)
	End Sub

	Private Sub AddSheetToContentTypes(sheetId As Integer)
		Dim contentTypesXmlPath As String = Path.Combine(_readXmlFolderPath, "[Content_Types].xml")
		Dim xmlContentTypesDoc As New XmlDocument()
		xmlContentTypesDoc.Load(contentTypesXmlPath)

		' Manage namespaces
		Dim nsManager As New XmlNamespaceManager(xmlContentTypesDoc.NameTable)
		nsManager.AddNamespace("ct", "http://schemas.openxmlformats.org/package/2006/content-types")

		' Add new Override for the new sheet
		Dim newSheetOverride As XmlElement = xmlContentTypesDoc.CreateElement("Override", "http://schemas.openxmlformats.org/package/2006/content-types")
		newSheetOverride.SetAttribute("PartName", "/xl/worksheets/sheet" & sheetId & ".xml")
		newSheetOverride.SetAttribute("ContentType", "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml")

		' Append the new Override to the Types node
		Dim typesNode As XmlNode = xmlContentTypesDoc.SelectSingleNode("//ct:Types", nsManager)
		typesNode.AppendChild(newSheetOverride)

		' Save the modified [Content_Types].xml
		xmlContentTypesDoc.Save(contentTypesXmlPath)
	End Sub


	Private Sub UpdateWorkbookXml(sheetName As String, newSheetId As String, sheetId As Integer)

		'--------------Update workbook XML-----------------------
		Dim xmlWorkbookDoc As New XmlDocument()
		xmlWorkbookDoc.Load(_workbookXmlPath)

		' Namespace URI for the spreadsheet
		Dim namespaceUri As String = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
		Dim relNamespaceUri As String = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
		Dim x15acNamespaceUri As String = "http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac"

		'Create a new <sheet> element
		Dim sheetNode As XmlNode = xmlWorkbookDoc.CreateElement("sheet", namespaceUri)
		sheetNode.Attributes.Append(xmlWorkbookDoc.CreateAttribute("name")).Value = sheetName
		sheetNode.Attributes.Append(xmlWorkbookDoc.CreateAttribute("sheetId")).Value = sheetId
		sheetNode.Attributes.Append(xmlWorkbookDoc.CreateAttribute("r", "id", relNamespaceUri)).Value = newSheetId
		Dim nsManager As New XmlNamespaceManager(xmlWorkbookDoc.NameTable)
		AddWorbookXmlNameSpace(nsManager)
		nsManager.AddNamespace("x15ac", x15acNamespaceUri)

		' Find the <sheets> node
		Dim sheetsNode As XmlNode = xmlWorkbookDoc.SelectSingleNode("//ns:sheets", nsManager)
		' Append the new <sheet> element to the <sheets> element
		sheetsNode.AppendChild(sheetNode)

		' Find or create the <x15ac:absPath> node and set the URL attribute
		Dim absPathNode As XmlNode = xmlWorkbookDoc.SelectSingleNode("//x15ac:absPath", nsManager)
		If absPathNode Is Nothing Then
			absPathNode = xmlWorkbookDoc.CreateElement("x15ac:absPath", "http://schemas.microsoft.com/office/spreadsheetml/2010/11/ac")
			sheetsNode.AppendChild(absPathNode)
		End If
		absPathNode.Attributes.Append(xmlWorkbookDoc.CreateAttribute("url")).Value = _writeExcelFolderPath + "\"
		' Save the modified workbook XML
		xmlWorkbookDoc.Save(_workbookXmlPath)

		'-------------Finished Update workbook XML-----------------------

	End Sub

	Private Sub UpdateNewSheetXml(sheetXmlPath As String)

		'---------------Update the Sheet.xml-----------------------------
		Dim doc As New XmlDocument()

		'Create worksheet element with namespaces
		Dim worksheetElement As XmlElement = doc.CreateElement("worksheet", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
		worksheetElement.SetAttribute("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
		worksheetElement.SetAttribute("xmlns:mc", "http://schemas.openxmlformats.org/markup-compatibility/2006")
		worksheetElement.SetAttribute("xmlns:x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac")
		worksheetElement.SetAttribute("xmlns:xr", "http://schemas.microsoft.com/office/spreadsheetml/2014/revision")
		worksheetElement.SetAttribute("xmlns:xr2", "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2")
		worksheetElement.SetAttribute("xmlns:xr3", "http://schemas.microsoft.com/office/spreadsheetml/2016/revision3")
		worksheetElement.SetAttribute("mc:Ignorable", "x14ac xr xr2 xr3")
		worksheetElement.SetAttribute("xr:uid", "{00000000-0001-0000-0200-000000000000}")

		'Add child elements
		Dim dimensionElement As XmlElement = doc.CreateElement("dimension")
		dimensionElement.SetAttribute("ref", "A1")
		worksheetElement.AppendChild(dimensionElement)

		Dim sheetViewsElement As XmlElement = doc.CreateElement("sheetViews")
		Dim sheetViewElement As XmlElement = doc.CreateElement("sheetView")
		sheetViewElement.SetAttribute("tabSelected", "1")
		sheetViewElement.SetAttribute("workbookViewId", "0")
		sheetViewsElement.AppendChild(sheetViewElement)
		worksheetElement.AppendChild(sheetViewsElement)

		Dim sheetFormatPrElement As XmlElement = doc.CreateElement("sheetFormatPr")
		sheetFormatPrElement.SetAttribute("defaultRowHeight", "15")
		sheetFormatPrElement.SetAttribute("x14ac:dyDescent", "0.25")
		worksheetElement.AppendChild(sheetFormatPrElement)

		worksheetElement.AppendChild(doc.CreateElement("sheetData"))

		Dim pageMarginsElement As XmlElement = doc.CreateElement("pageMargins")
		pageMarginsElement.SetAttribute("left", "0.7")
		pageMarginsElement.SetAttribute("right", "0.7")
		pageMarginsElement.SetAttribute("top", "0.75")
		pageMarginsElement.SetAttribute("bottom", "0.75")
		pageMarginsElement.SetAttribute("header", "0.3")
		pageMarginsElement.SetAttribute("footer", "0.3")
		worksheetElement.AppendChild(pageMarginsElement)

		'Append worksheet element to the document
		doc.AppendChild(worksheetElement)

		doc.Save(sheetXmlPath)
		'------------- Finished Updating Sheet.xml----------------
	End Sub

	Private Sub UpdateRelXml(sheetXmlPath As String, newSheetId As String, sheetId As Integer)

		'--------------Update workbook.xml.rels File-----------------------
		Dim workbookRelsPath As String = _relXmlPath
		Dim xmlWorkbookRelsDoc As New XmlDocument()
		xmlWorkbookRelsDoc.Load(workbookRelsPath)
		Dim nsManagerRels As New XmlNamespaceManager(xmlWorkbookRelsDoc.NameTable)
		AddRelXmlNameSpace(nsManagerRels)


		' Create a new <Relationship> element
		Dim relationshipNode As XmlNode = xmlWorkbookRelsDoc.CreateElement("Relationship", nsManagerRels.LookupNamespace("r"))
		Dim idAttribute As XmlAttribute = xmlWorkbookRelsDoc.CreateAttribute("Id")
		idAttribute.Value = newSheetId
		relationshipNode.Attributes.Append(idAttribute)

		Dim typeAttribute As XmlAttribute = xmlWorkbookRelsDoc.CreateAttribute("Type")
		typeAttribute.Value = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
		relationshipNode.Attributes.Append(typeAttribute)

		Dim targetAttribute As XmlAttribute = xmlWorkbookRelsDoc.CreateAttribute("Target")
		targetAttribute.Value = "worksheets/sheet" + sheetId.ToString() + ".xml"
		relationshipNode.Attributes.Append(targetAttribute)


		' Append the new <Relationship> element to the <Relationships> element
		Dim relationshipsNode As XmlNode = xmlWorkbookRelsDoc.SelectSingleNode("//r:Relationships", nsManagerRels)
		relationshipsNode.AppendChild(relationshipNode)

		' Save the modified workbook.xml.rels
		xmlWorkbookRelsDoc.Save(workbookRelsPath)
		'-------------Finished Update workbook.xml.rels File-----------------------

	End Sub

	Public Shared Sub CreateNewExcel(excelFilePath As String, Optional SheetName As String = "Sheet1")
		Dim _blankWorkbookFolderPath As String = Path.Combine(Path.GetPathRoot(Environment.SystemDirectory), "Automation", "XML_Processor", "blank")
		Dim _writeExcelFolderPath As String = Path.Combine(Path.GetPathRoot(Environment.SystemDirectory), "Automation", "XML_Processor", "write")
		'Delete the write folder path:
		Reusables.DeleteAFolder(_blankWorkbookFolderPath)
		Reusables.DeleteAFolder(_writeExcelFolderPath)
		'create the write folder path inside of the processor folder:
		Reusables.CreateAFolder(_blankWorkbookFolderPath)
		Reusables.DeleteAFolder(_writeExcelFolderPath)
		'create rels folder
		Dim _relsDir As String = Path.Combine(_blankWorkbookFolderPath, "_rels")
		Dim _docPropsDir As String = Path.Combine(_blankWorkbookFolderPath, "docProps")
		Dim _xlDir As String = Path.Combine(_blankWorkbookFolderPath, "xl")
		Dim _xlRelsDir As String = Path.Combine(_xlDir, "_rels")
		Dim _xlThemeDir As String = Path.Combine(_xlDir, "theme")
		Dim _xlWorksheetsDir As String = Path.Combine(_xlDir, "worksheets")

		' Create necessary directories
		Directory.CreateDirectory(_relsDir)
		Directory.CreateDirectory(_docPropsDir)
		Directory.CreateDirectory(_xlDir)
		Directory.CreateDirectory(_xlRelsDir)
		Directory.CreateDirectory(_xlThemeDir)
		Directory.CreateDirectory(_xlWorksheetsDir)

		Dim b As New Dictionary(Of String, Object()) From {
		{"theme1.xml", New Object() {_xlThemeDir, themeXmlString}},
		{"workbook.xml", New Object() {_xlDir, workBookXmlString}},
		{"workbook.xml.rels", New Object() {_xlRelsDir, workbookrelsXmlString}},
		{"sheet1.xml", New Object() {_xlWorksheetsDir, sheetXmlString}},
		{".rels", New Object() {_relsDir, relXml}},
		{"core.xml", New Object() {_docPropsDir, coreXmlString}},
		{"app.xml", New Object() {_docPropsDir, appXmlString}},
		{"[Content_Types].xml", New Object() {_blankWorkbookFolderPath, contentTypesXmlString}},
		{"styles.xml", New Object() {_xlDir, styleXmlString}}
	}

		For Each k As String In b.Keys
			Dim xmlFileName As String = k
			Dim xmlFilePath As String = Path.Combine(b(k)(0), xmlFileName)
			Dim xmlString As String = b(k)(1).ToString().Replace("RobotSheetName", SheetName)
			Dim doc As New XmlDocument()
			doc.LoadXml(xmlString)
			doc.Save(xmlFilePath)
		Next

		Reusables.RevertXmlParentFolderToExcelFile(excelFilePath, _blankWorkbookFolderPath)

	End Sub

	'---------------------end of subs which would be called for new sheet creatioin-------------------

	'The dispose function would reset the cache files and empty all variables: this function is called at the end of each process it would delete the xml_Process folder
	Public Sub Dispose()
		Dim generatedXmlFolderPath As String = Path.Combine(Path.GetPathRoot(Environment.SystemDirectory), "Automation", "XML_Processor")
		Reusables.DeleteAFolder(generatedXmlFolderPath)
	End Sub

	'These public shared readonly members hold the XMLs of a fresh new Excel. These XMLs can be written back into the expected folder structure and then zipped back to produce a new Excel :
	Public Shared ReadOnly themeXmlString As String = "<a:theme xmlns:a='http://schemas.openxmlformats.org/drawingml/2006/main' name='Office Theme'> <a:themeElements> <a:clrScheme name='Office'> <a:dk1> <a:sysClr val='windowText' lastClr='000000'/> </a:dk1> <a:lt1> <a:sysClr val='window' lastClr='FFFFFF'/> </a:lt1> <a:dk2> <a:srgbClr val='44546A'/> </a:dk2> <a:lt2> <a:srgbClr val='E7E6E6'/> </a:lt2> <a:accent1> <a:srgbClr val='5B9BD5'/> </a:accent1> <a:accent2> <a:srgbClr val='ED7D31'/> </a:accent2> <a:accent3> <a:srgbClr val='A5A5A5'/> </a:accent3> <a:accent4> <a:srgbClr val='FFC000'/> </a:accent4> <a:accent5> <a:srgbClr val='4472C4'/> </a:accent5> <a:accent6> <a:srgbClr val='70AD47'/> </a:accent6> <a:hlink> <a:srgbClr val='0563C1'/> </a:hlink> <a:folHlink> <a:srgbClr val='954F72'/> </a:folHlink> </a:clrScheme> <a:fontScheme name='Office'> <a:majorFont> <a:latin typeface='Calibri Light' panose='020F0302020204030204'/> <a:ea typeface=''/> <a:cs typeface=''/> <a:font script='Jpan' typeface='游ゴシック Light'/> <a:font script='Hang' typeface='맑은 고딕'/> <a:font script='Hans' typeface='等线 Light'/> <a:font script='Hant' typeface='新細明體'/> <a:font script='Arab' typeface='Times New Roman'/> <a:font script='Hebr' typeface='Times New Roman'/> <a:font script='Thai' typeface='Tahoma'/> <a:font script='Ethi' typeface='Nyala'/> <a:font script='Beng' typeface='Vrinda'/> <a:font script='Gujr' typeface='Shruti'/> <a:font script='Khmr' typeface='MoolBoran'/> <a:font script='Knda' typeface='Tunga'/> <a:font script='Guru' typeface='Raavi'/> <a:font script='Cans' typeface='Euphemia'/> <a:font script='Cher' typeface='Plantagenet Cherokee'/> <a:font script='Yiii' typeface='Microsoft Yi Baiti'/> <a:font script='Tibt' typeface='Microsoft Himalaya'/> <a:font script='Thaa' typeface='MV Boli'/> <a:font script='Deva' typeface='Mangal'/> <a:font script='Telu' typeface='Gautami'/> <a:font script='Taml' typeface='Latha'/> <a:font script='Syrc' typeface='Estrangelo Edessa'/> <a:font script='Orya' typeface='Kalinga'/> <a:font script='Mlym' typeface='Kartika'/> <a:font script='Laoo' typeface='DokChampa'/> <a:font script='Sinh' typeface='Iskoola Pota'/> <a:font script='Mong' typeface='Mongolian Baiti'/> <a:font script='Viet' typeface='Times New Roman'/> <a:font script='Uigh' typeface='Microsoft Uighur'/> <a:font script='Geor' typeface='Sylfaen'/> </a:majorFont> <a:minorFont> <a:latin typeface='Calibri' panose='020F0502020204030204'/> <a:ea typeface=''/> <a:cs typeface=''/> <a:font script='Jpan' typeface='游ゴシック'/> <a:font script='Hang' typeface='맑은 고딕'/> <a:font script='Hans' typeface='等线'/> <a:font script='Hant' typeface='新細明體'/> <a:font script='Arab' typeface='Arial'/> <a:font script='Hebr' typeface='Arial'/> <a:font script='Thai' typeface='Tahoma'/> <a:font script='Ethi' typeface='Nyala'/> <a:font script='Beng' typeface='Vrinda'/> <a:font script='Gujr' typeface='Shruti'/> <a:font script='Khmr' typeface='DaunPenh'/> <a:font script='Knda' typeface='Tunga'/> <a:font script='Guru' typeface='Raavi'/> <a:font script='Cans' typeface='Euphemia'/> <a:font script='Cher' typeface='Plantagenet Cherokee'/> <a:font script='Yiii' typeface='Microsoft Yi Baiti'/> <a:font script='Tibt' typeface='Microsoft Himalaya'/> <a:font script='Thaa' typeface='MV Boli'/> <a:font script='Deva' typeface='Mangal'/> <a:font script='Telu' typeface='Gautami'/> <a:font script='Taml' typeface='Latha'/> <a:font script='Syrc' typeface='Estrangelo Edessa'/> <a:font script='Orya' typeface='Kalinga'/> <a:font script='Mlym' typeface='Kartika'/> <a:font script='Laoo' typeface='DokChampa'/> <a:font script='Sinh' typeface='Iskoola Pota'/> <a:font script='Mong' typeface='Mongolian Baiti'/> <a:font script='Viet' typeface='Arial'/> <a:font script='Uigh' typeface='Microsoft Uighur'/> <a:font script='Geor' typeface='Sylfaen'/> </a:minorFont> </a:fontScheme> <a:fmtScheme name='Office'> <a:fillStyleLst> <a:solidFill> <a:schemeClr val='phClr'/> </a:solidFill> <a:gradFill rotWithShape='1'> <a:gsLst> <a:gs pos='0'> <a:schemeClr val='phClr'> <a:lumMod val='110000'/> <a:satMod val='105000'/> <a:tint val='67000'/> </a:schemeClr> </a:gs> <a:gs pos='50000'> <a:schemeClr val='phClr'> <a:lumMod val='105000'/> <a:satMod val='103000'/> <a:tint val='73000'/> </a:schemeClr> </a:gs> <a:gs pos='100000'> <a:schemeClr val='phClr'> <a:lumMod val='105000'/> <a:satMod val='109000'/> <a:tint val='81000'/> </a:schemeClr> </a:gs> </a:gsLst> <a:lin ang='5400000' scaled='0'/> </a:gradFill> <a:gradFill rotWithShape='1'> <a:gsLst> <a:gs pos='0'> <a:schemeClr val='phClr'> <a:satMod val='103000'/> <a:lumMod val='102000'/> <a:tint val='94000'/> </a:schemeClr> </a:gs> <a:gs pos='50000'> <a:schemeClr val='phClr'> <a:satMod val='110000'/> <a:lumMod val='100000'/> <a:shade val='100000'/> </a:schemeClr> </a:gs> <a:gs pos='100000'> <a:schemeClr val='phClr'> <a:lumMod val='99000'/> <a:satMod val='120000'/> <a:shade val='78000'/> </a:schemeClr> </a:gs> </a:gsLst> <a:lin ang='5400000' scaled='0'/> </a:gradFill> </a:fillStyleLst> <a:lnStyleLst> <a:ln w='6350' cap='flat' cmpd='sng' algn='ctr'> <a:solidFill> <a:schemeClr val='phClr'/> </a:solidFill> <a:prstDash val='solid'/> <a:miter lim='800000'/> </a:ln> <a:ln w='12700' cap='flat' cmpd='sng' algn='ctr'> <a:solidFill> <a:schemeClr val='phClr'/> </a:solidFill> <a:prstDash val='solid'/> <a:miter lim='800000'/> </a:ln> <a:ln w='19050' cap='flat' cmpd='sng' algn='ctr'> <a:solidFill> <a:schemeClr val='phClr'/> </a:solidFill> <a:prstDash val='solid'/> <a:miter lim='800000'/> </a:ln> </a:lnStyleLst> <a:effectStyleLst> <a:effectStyle> <a:effectLst/> </a:effectStyle> <a:effectStyle> <a:effectLst/> </a:effectStyle> <a:effectStyle> <a:effectLst> <a:outerShdw blurRad='57150' dist='19050' dir='5400000' algn='ctr' rotWithShape='0'> <a:srgbClr val='000000'> <a:alpha val='63000'/> </a:srgbClr> </a:outerShdw> </a:effectLst> </a:effectStyle> </a:effectStyleLst> <a:bgFillStyleLst> <a:solidFill> <a:schemeClr val='phClr'/> </a:solidFill> <a:solidFill> <a:schemeClr val='phClr'> <a:tint val='95000'/> <a:satMod val='170000'/> </a:schemeClr> </a:solidFill> <a:gradFill rotWithShape='1'> <a:gsLst> <a:gs pos='0'> <a:schemeClr val='phClr'> <a:tint val='93000'/> <a:satMod val='150000'/> <a:shade val='98000'/> <a:lumMod val='102000'/> </a:schemeClr> </a:gs> <a:gs pos='50000'> <a:schemeClr val='phClr'> <a:tint val='98000'/> <a:satMod val='130000'/> <a:shade val='90000'/> <a:lumMod val='103000'/> </a:schemeClr> </a:gs> <a:gs pos='100000'> <a:schemeClr val='phClr'> <a:shade val='63000'/> <a:satMod val='120000'/> </a:schemeClr> </a:gs> </a:gsLst> <a:lin ang='5400000' scaled='0'/> </a:gradFill> </a:bgFillStyleLst> </a:fmtScheme> </a:themeElements> <a:objectDefaults/> <a:extraClrSchemeLst/> <a:extLst> <a:ext uri='{05A4C25C-085E-4340-85A3-A5531E510DB2}'> <thm15:themeFamily xmlns:thm15='http://schemas.microsoft.com/office/thememl/2012/main' name='Office Theme' id='{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}' vid='{4A3C46E8-61CC-4603-A589-7422A47A8E4A}'/> </a:ext> </a:extLst> </a:theme>"
	Public Shared ReadOnly workBookXmlString As String = "<workbook xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main' xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships' xmlns:mc='http://schemas.openxmlformats.org/markup-compatibility/2006' xmlns:x15='http://schemas.microsoft.com/office/spreadsheetml/2010/11/main' mc:Ignorable='x15'> <fileVersion appName='xl' lastEdited='6' lowestEdited='6' rupBuild='14420'/> <workbookPr defaultThemeVersion='164011'/> <bookViews> <workbookView xWindow='0' yWindow='0' windowWidth='22260' windowHeight='12645'/> </bookViews> <sheets> <sheet name='RobotSheetName' sheetId='1' r:id='rId1'/> </sheets> <calcPr calcId='162913'/> <extLst> <ext xmlns:x15='http://schemas.microsoft.com/office/spreadsheetml/2010/11/main' uri='{140A7094-0E35-4892-8432-C4D2E57EDEB5}'> <x15:workbookPr chartTrackingRefBase='1'/> </ext> </extLst> </workbook>"
	Public Shared ReadOnly styleXmlString As String = "<styleSheet xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main' xmlns:mc='http://schemas.openxmlformats.org/markup-compatibility/2006' xmlns:x14ac='http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac' xmlns:x16r2='http://schemas.microsoft.com/office/spreadsheetml/2015/02/main' mc:Ignorable='x14ac x16r2'> <fonts count='1' x14ac:knownFonts='1'> <font> <sz val='11'/> <color theme='1'/> <name val='Calibri'/> <family val='2'/> <scheme val='minor'/> </font> </fonts> <fills count='2'> <fill> <patternFill patternType='none'/> </fill> <fill> <patternFill patternType='gray125'/> </fill> </fills> <borders count='1'> <border> <left/> <right/> <top/> <bottom/> <diagonal/> </border> </borders> <cellStyleXfs count='1'> <xf numFmtId='0' fontId='0' fillId='0' borderId='0'/> </cellStyleXfs> <cellXfs count='1'> <xf numFmtId='0' fontId='0' fillId='0' borderId='0' xfId='0'/> </cellXfs> <cellStyles count='1'> <cellStyle name='Normal' xfId='0' builtinId='0'/> </cellStyles> <dxfs count='0'/> <tableStyles count='0' defaultTableStyle='TableStyleMedium2' defaultPivotStyle='PivotStyleLight16'/> <extLst> <ext xmlns:x14='http://schemas.microsoft.com/office/spreadsheetml/2009/9/main' uri='{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}'> <x14:slicerStyles defaultSlicerStyle='SlicerStyleLight1'/> </ext> <ext xmlns:x15='http://schemas.microsoft.com/office/spreadsheetml/2010/11/main' uri='{9260A510-F301-46a8-8635-F512D64BE5F5}'> <x15:timelineStyles defaultTimelineStyle='TimeSlicerStyleLight1'/> </ext> </extLst> </styleSheet>"
	Public Shared ReadOnly sheetXmlString As String = "<worksheet xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main' xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships' xmlns:mc='http://schemas.openxmlformats.org/markup-compatibility/2006' xmlns:x14ac='http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac' mc:Ignorable='x14ac'> <dimension ref='A1'/> <sheetViews> <sheetView tabSelected='1' workbookViewId='0'/> </sheetViews> <sheetFormatPr defaultRowHeight='15' x14ac:dyDescent='0.25'/> <sheetData/> <pageMargins left='0.7' right='0.7' top='0.75' bottom='0.75' header='0.3' footer='0.3'/> </worksheet>"
	Public Shared ReadOnly workbookrelsXmlString As String = "<?xml version='1.0' encoding='UTF-8' standalone='yes'?> <Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId3' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles' Target='styles.xml'/><Relationship Id='rId2' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme' Target='theme/theme1.xml'/><Relationship Id='rId1' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet' Target='worksheets/sheet1.xml'/></Relationships>"
	Public Shared ReadOnly appXmlString As String = "<Properties xmlns='http://schemas.openxmlformats.org/officeDocument/2006/extended-properties' xmlns:vt='http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes'> <Application>Microsoft Excel</Application> <DocSecurity>0</DocSecurity> <ScaleCrop>false</ScaleCrop> <HeadingPairs> <vt:vector size='2' baseType='variant'> <vt:variant> <vt:lpstr>Worksheets</vt:lpstr> </vt:variant> <vt:variant> <vt:i4>1</vt:i4> </vt:variant> </vt:vector> </HeadingPairs> <TitlesOfParts> <vt:vector size='1' baseType='lpstr'> <vt:lpstr>RobotSheetName</vt:lpstr> </vt:vector> </TitlesOfParts> <Company/> <LinksUpToDate>false</LinksUpToDate> <SharedDoc>false</SharedDoc> <HyperlinksChanged>false</HyperlinksChanged> <AppVersion>16.0300</AppVersion> </Properties>"
	Public Shared ReadOnly coreXmlString As String = "<cp:coreProperties xmlns:cp='http://schemas.openxmlformats.org/package/2006/metadata/core-properties' xmlns:dc='http://purl.org/dc/elements/1.1/' xmlns:dcterms='http://purl.org/dc/terms/' xmlns:dcmitype='http://purl.org/dc/dcmitype/' xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance'> <dc:creator>Robot</dc:creator> <cp:lastModifiedBy/> <dcterms:created xsi:type='dcterms:W3CDTF'>2015-06-05T18:17:20Z</dcterms:created> <dcterms:modified xsi:type='dcterms:W3CDTF'>2015-06-05T18:17:26Z</dcterms:modified> </cp:coreProperties>"
	Public Shared ReadOnly contentTypesXmlString As String = "<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'> <Default Extension='rels' ContentType='application/vnd.openxmlformats-package.relationships+xml'/> <Default Extension='xml' ContentType='application/xml'/> <Override PartName='/xl/workbook.xml' ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'/> <Override PartName='/xl/worksheets/sheet1.xml' ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'/> <Override PartName='/xl/theme/theme1.xml' ContentType='application/vnd.openxmlformats-officedocument.theme+xml'/> <Override PartName='/xl/styles.xml' ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml'/> <Override PartName='/docProps/core.xml' ContentType='application/vnd.openxmlformats-package.core-properties+xml'/> <Override PartName='/docProps/app.xml' ContentType='application/vnd.openxmlformats-officedocument.extended-properties+xml'/> </Types>"
	Public Shared ReadOnly relXml As String = "<?xml version='1.0' encoding='UTF-8' standalone='yes'?> <Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId3' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties' Target='docProps/app.xml'/><Relationship Id='rId2' Type='http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties' Target='docProps/core.xml'/><Relationship Id='rId1' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument' Target='xl/workbook.xml'/></Relationships>"


End Class




