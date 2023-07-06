Attribute VB_Name = "Module1"
Sub ExportSheetsAsFiles()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim exportFolder As String
    
    ' Set the workbook object
    Set wb = Workbooks.Open("C:\Users\gis\Desktop\CVD_Deaths_Project\Root_table.xlsx")
    
    ' Set the export folder path
    exportFolder = "C:\Users\gis\Desktop\CVD_Deaths_Project\" ' Specify the desired export folder path
    
    ' Loop through each sheet in the workbook
    For Each ws In wb.Sheets
        ' Save the sheet as a separate file in .xlsx format
        ws.Copy
        ActiveWorkbook.SaveAs exportFolder & ws.Name & ".xlsx", FileFormat:=xlOpenXMLWorkbook
        ActiveWorkbook.Close SaveChanges:=False
    Next ws
    
    ' Close the original workbook
    wb.Close SaveChanges:=False
    
    ' Display a message box indicating the completion
    MsgBox "All sheets exported successfully.", vbInformation
End Sub

