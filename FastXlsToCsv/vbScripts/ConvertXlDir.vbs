Dim objExcel, objWorkbook, fso, file

' Check if the required arguments are provided
If WScript.Arguments.Count < 2 Then
    WScript.Echo "Usage: cscript ConvertXlDir.vbs <InputDir> <OutputDir>"
    WScript.Quit(1)
End If

' Get the Excel and CSV file paths from the command line arguments
inputDir = WScript.Arguments(0)
outputDir = WScript.Arguments(1)

excelFile = Dir(inputDir & "/")
