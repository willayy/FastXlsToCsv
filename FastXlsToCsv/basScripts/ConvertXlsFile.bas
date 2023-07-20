Attribute VB_Name = "Module1"
Public Sub ConvertXlsFileToCsv(inputFile As String, outputDir As String)
    Dim xlsFile As String
    Dim csvFile As String
    Dim wb As Workbook
    
    ' Disable screen updating for faster processing
    Application.ScreenUpdating = False
    
    ' Loop through each file in the input folder
    Set fso = CreateObject("Scripting.FileSystemObject")
    xlsFile = File(Name, inputFile)
    
    if  fso.GetExtensionName(xlsFile) = "xls"
        ' Create the output CSV file name with the same name as the .xls file
        csvFile = outputDir & "/" & Replace(xlsFile, ".xls", ".csv", , , vbTextCompare)
        
        ' Open the Excel file and save as CSV
        Set wb = Workbooks.Open(inputDir & "/" & xlsFile)
        wb.SaveAs Filename:=csvFile, FileFormat:=xlCSVUTF8, CreateBackup:=False
        wb.Close SaveChanges:=True
    ElseIf fso.GetExtensionName(xlsFile) = "xlsx"
        ' Create the output CSV file name with the same name as the .xls file
        csvFile = outputDir & "/" & Replace(xlsFile, ".xlsx", ".csv", , , vbTextCompare)
        
        ' Open the Excel file and save as CSV
        Set wb = Workbooks.Open(inputDir & "/" & xlsFile)
        wb.SaveAs Filename:=csvFile, FileFormat:=xlCSVUTF8, CreateBackup:=False
        wb.Close SaveChanges:=True
    End if
    
    ' Enable screen updating after processing
    Application.ScreenUpdating = True
End Sub