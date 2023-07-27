# FastXlsToCsv
### Made by William Norland, 2023

Fast way to convert large .xls/.xlsx files to CSV by saving them via a vbs running via windows script host.
## Why?
During a project i did i noticed how slow pandas read excel files to dataframe, because pandas reads csv to dataframe much faster i decied to release my solution
for turning excel into csv!

## Dependencies
* Excel (Only tested on excel 2016) and .
* Windows operating system (Only tested on Windows 10).
* Only tested on Python 3.11, should probably work on any python that has os and subprocess.
* Windows script host (Exists on pretty much every windows ever).

## Usage
### Install
![Screenshot (2)](https://github.com/willayy/FastXlsToCsv/assets/117913560/219b6291-15c9-4b90-8d90-21404c50acfa)
### Import
![Screenshot (3)](https://github.com/willayy/FastXlsToCsv/assets/117913560/c73a81b7-c1d7-4e13-9980-8a5e7d6a7217)
### Use
```python
inputFile: str = r"c:\Users\someone\Desktop\ExcelFiles\excelFileThatWantsTobeCsv.xlsx"
inputDir: str = r"c:\Users\someone\Desktop\ExcelFiles"
outputDir: str =  r"c:\Users\someone\Desktop\FolderForExportedCsvs"
XlsConverter.convertXlFileToCsv(inputFile, output)
XlsConverter.convertXlDir(inputDir, output)
```

## License
Relased under the MIT License, check FastXlsToCsv/LICENSE for more information.

