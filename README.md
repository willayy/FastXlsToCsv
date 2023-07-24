# FastXlsToCsv
### Made by William Norland, 2023
Fast way to convert large .xls/.xlsx files to CSV by saving them via .vbs and runnin the vbs with windows script host.
## Dependencies
Excel and a Windows operating system, no other python module dependencies!
## Usage
### Import
![Screenshot (52)](https://github.com/willayy/FastXlsToCsv/assets/117913560/4ff08183-e8db-4c71-941e-e9864462c9f3)
### Use

```python
input = r"c:\Users\someone\Desktop\ExcelFiles\excelFileThatWantsTobeCsv.xlsx"
output =  r"c:\Users\someone\Desktop\FolderForExportedCsvs"
XlsConverter.convertXlsFileToCsv(input, output)
```


