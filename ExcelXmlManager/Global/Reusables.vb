Imports System.Data
Imports System.Text.RegularExpressions
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

End Class
