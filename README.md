# FastXlsToCsv
### Made by William Norland, 2023
Fast way to convert large .xls/.xlsx files to CSV by saving them via a vbs and running the vbs with windows script host.
## Dependencies
* Excel (Only tested on excel 2016) and 
* Windows operating system (Only tested on Windows 10)
* Only tested on Python 3.11, should probably work on any python that has os and subprocess
* Windows script host (Exists on pretty much every windows ever)

## Usage
### Import
![Screenshot (52)](https://github.com/willayy/FastXlsToCsv/assets/117913560/4ff08183-e8db-4c71-941e-e9864462c9f3)
### Use

```python
input = r"c:\Users\someone\Desktop\ExcelFiles\excelFileThatWantsTobeCsv.xlsx"
output =  r"c:\Users\someone\Desktop\FolderForExportedCsvs"
XlsConverter.convertXlsFileToCsv(input, output)
```


