Imports System
Imports System.Data

Module Program
    Sub Main(args As String())
        Console.WriteLine("Testing Functions !")
        Dim dt As DataTable = Action.ReadExcelToTable("C:\Automation\Test\Test.xlsx", "Sheet1", True)
    End Sub
End Module
