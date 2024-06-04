Imports System
Imports System.Data

Module Program
    Sub Main(args As String())
        Console.WriteLine("Testing Functions !")
        ExcelXmlAction.CreateSheet("C:\Automation\Test\Test.xlsx", "Man")

        Dim dt As DataTable = ExcelXmlAction.ReadExcelToTable("C:\Automation\Test\Test.xlsx", "Sheet1", False)
        'ExcelXmlAction.WriteTableToNewExcel(dt, "C:\Automation\Robo.xlsx", "Sheet1")
        Dim sheetName As String = ExcelXmlAction.GetSheetNameBySheetIndex("C:\Automation\Test\Test.xlsx", 0)
        Dim c As String = ExcelXmlAction.GetCellValue("C:\Automation\Test\Test.xlsx", "Sheet1", "G16")


    End Sub
End Module
