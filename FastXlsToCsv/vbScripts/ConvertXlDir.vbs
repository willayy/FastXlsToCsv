Dim objExcel, objWorkbook, fso, file

' Check if the required arguments are provided
If WScript.Arguments.Count < 2 Then
    WScript.Echo "Usage: cscript ConvertXlDir.vbs <InputDir> <OutputDir>"
    WScript.Quit(1)
End If

' Get the Excel and CSV file paths from the command line arguments
inputDir = WScript.Arguments(0)
outputDir = WScript.Arguments(1)

' Function to get the filename without extension
Function RemoveExcelExtension(fileName)
    fileName = Replace(fileName, ".xlsx", "", 1, -1, vbTextCompare) ' Replace .xlsx with an empty string
    fileName = Replace(fileName, ".xls", "", 1, -1, vbTextCompare) ' Replace .xls with an empty string
    RemoveExcelExtension = fileName
End Function

' Create an Excel application object
Set objExcel = CreateObject("Excel.Application")

' Hide Excel application window (optional)
objExcel.Visible = False

Set fso = CreateObject("Scripting.FileSystemObject")

' Iterating through the files in a folder
For Each excelFile In fso.GetFolder(inputDir & "/").Files
    if fso.getExtensionName(excelFile) = "xls" OR fso.getExtensionName(excelFile) = "xlsx" Then
        Set file = fso.GetFile(excelFile)
        excelFileName = RemoveExcelExtension(file.Name)

        csvFile = outputDir & "/" & excelFileName & ".csv"
        
        ' Open the Excel file
        Set objWorkbook = objExcel.Workbooks.Open(excelFile)

        ' Save the workbook as CSV
        objWorkbook.SaveAs csvFile, 6 ' 6 is the CSV file format
    End if
Next

' Close the workbook and Excel application
objWorkbook.Close False ' False means don't save changes
objExcel.Quit

' Release the objects from memory
Set objWorkbook = Nothing
Set objExcel = Nothing
Set fso = Nothing
Set file = Nothing