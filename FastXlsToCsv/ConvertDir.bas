Attribute VB_Name = "Module1"
Public Sub ConvertXlsToCsv(inputDir As String, outputDir As String)
    Dim inputFolder As String
    Dim outputFolder As String
    Dim xlsFile As String
    Dim csvFile As String
    Dim wb As Workbook
    
    ' Disable screen updating for faster processing
    Application.ScreenUpdating = False
    
    ' Loop through each file in the input folder
    Debug.Print inputDir
    xlsFile = Dir(inputDir + "/")
    Do While xlsFile <> ""
        ' Create the output CSV file name with the same name as the .xls file
        csvFile = outputDir & "/" & Replace(xlsFile, ".xls", ".csv", , , vbTextCompare)
        
        ' Open the Excel file and save as CSV
        Set wb = Workbooks.Open(inputDir & "/" & xlsFile)
        wb.SaveAs Filename:=csvFile, FileFormat:=xlCSVUTF8, CreateBackup:=False
        wb.Close SaveChanges:=True
        
        ' Move to the next file
        xlsFile = Dir
    Loop
    
    ' Enable screen updating after processing
    Application.ScreenUpdating = True
End Sub
