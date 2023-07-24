"""
## XlsConverter Module
#### Made by:
William Norland, 2023, https://github.com/willayy
### Requirements:
You have Excel installed on your computer, only tested with Microsoft Excel 2016.
### How to use?:
Call methods convertXlFile() and convertXlDir() from this module to convert excel files to csv,
### Github
https://github.com/willayy/FastXlsToCsv
"""

from FastXlsToCsv._XlsToCsvExceptions import *
import subprocess
import os

def convertXlDirToCsv(inputDir: str, outputDir: str) -> None:
    """
    ## Convert Xls/Xlsx Dir to Csv
    ### inputDir: 
    takes str as argument and will attempt to convert all .xlsx and .xls
    files in the directory corresponding to the str. The input may contain other files they will be ignored.
    ### outputDir: 
    takes str as argument and will attempt to export the csv files converted
    to the directory corresponding to the str.
    ## Raises:
    #### InputIsNotFileException: 
    If the input dir wasnt found, likely a problem with the str path.
    #### FastXlsToCsvModuleException: 
    If something within the module doesnt go/work as excpected, please check input and if that doesnt fix it
    contact developer https://github.com/willayy
    """

    if __checkIfPathIsDir(inputDir) is False:
        raise InputIsNotDirException()

    scriptPath: str = os.path.abspath(r"FastXlsToCsv\vbScripts\ConvertXlDir.vbs")

    try: 
        subprocess.run(["cscript", scriptPath, inputDir, outputDir], check = True)
    except:
        raise FastXlsToCsvModuleException()

def convertXlFileToCsv(inputFile: str, outputDir: str) -> None:
    """
    ## Convert Xls/Xlsx File to Csv
    ## inputFile:
    takes str argument and will attempt to convert only this xls/xlsx file.
    ## OutputDir:
    takes str argument and will attempt to exprot the csv files converted to the
    directory corresponding to the str.
    ## Raises:
    #### InputIsNotFileException: 
    If the input file wasnt found, likely a problem with the str path.
    #### InputIsNotXlFileException: 
    If the input file wasnt xls/xlsx, likely a file with the wrong exception.
    #### FastXlsToCsvModuleException: 
    If something doesnt go/work as excpected, please check input and if that doesnt fix it
    contact developer https://github.com/willayy
    """

    if __checkIfFileExists(inputFile) is False:
        raise InputIsNotFileException()
    
    if __checkIfFileIsXlFile(inputFile) is False:
        raise InputIsNotXlFileException()
    
    scriptPath: str = os.path.abspath(r"FastXlsToCsv\vbScripts\ConvertXlFile.vbs")
    
    try: 
        subprocess.run(["cscript", scriptPath, inputFile, outputDir])
    except:
        raise FastXlsToCsvModuleException()

def __checkIfPathIsDir(dirPath: str) -> bool:
    return os.path.isdir(dirPath)

def __checkIfFileExists(filePath: str) -> bool:
    return os.path.exists(filePath)

def __checkIfFileIsXlFile(filePath: str) -> bool:
    extension = os.path.splitext(filePath)[1]
    return (extension == ".xls" or extension == ".xlsx")