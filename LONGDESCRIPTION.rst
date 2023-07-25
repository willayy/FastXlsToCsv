FastXlsToCsv
--------------
*Made by William Norland, 2023*

Fast way to convert large .xls/.xlsx files to CSV by saving them via a vbs running via windows script host.

Dependencies
=============
* Excel (Only tested on excel 2016)
* Windows operating system (Only tested on Windows 10).
* Only tested on Python 3.11, should probably work on any python that has os and subprocess.
* Windows script host (Exists on pretty much every windows ever).


Usage
==========
==========
Install
==========

pip install FastXlsToCsv

==========
Import
==========

from FastXlToCsv import XlsConverter

======
Use
======

| inputFile: str = r"c:\\Users\\someone\\Desktop\\ExcelFiles\\excelFileThatWantsTobeCsv.xlsx"

| inputDir: str = r"c:\Users\\someone\\Desktop\\ExcelFiles"

| outputDir: str =  r"c:\\Users\\someone\\Desktop\\FolderForExportedCsvs"

| XlsConverter.convertXlFileToCsv(inputFile, output)

| XlsConverter.convertXlDir(inputDir, output)

License
=========
Relased under the MIT License, check FastXlsToCsv/LICENSE for more information.