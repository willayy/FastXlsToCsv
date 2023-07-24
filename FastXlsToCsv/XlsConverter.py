from XlsToCsvExceptions import *
import subprocess
import os

class XlsConverter:
    """
    ## XlsConverter class
    ### Requirements:
    You have Excel installed on your computer, right now its only tested with Microsoft Excel 2016.
    ### How to use?:
    Instantiate an object of this class to gain access to conversion methods, you can choose to exclude a excelPath argument in the constructor.
    Excluding excelPath will make XlsConverter try to find your excel installation by trying from a small set of options. As you may understand this is not very safe so 
    providing your own path is recommended.
    """
    excelPath: str

    def __init__(self) -> None:
        self
        

    def convertXlsDirToCsv(self, inputDir: str, outputDir: str) -> None:
        """
        ## Convert Xls/Xlsx Dir to Csv
        ### inputDir: 
        takes str as argument and will attempt to convert all .xlsx and .xls
        files in the directory corresponding to the str. The input may contain other files.
        ### outputDir: 
        takes str as argument and will attempt to export the csv files converted
        to the directory corresponding to the str.
        """
        if self.__checkIfPathIsDir(inputDir) is False:
            InputIsNotDirException()

        scriptPath: str = os.path.abspath(r"FastXlsToCsv\vbScripts\ConvertXlDir.vbs")

        try: 
            subprocess.run(["cscript", scriptPath, inputDir, outputDir], check = True)
        except Exception:
            raise FastXlsToCsvModuleError()

    def convertXlsFileToCsv(self, inputFile: str, outputDir: str) -> None:
        """
        ## Convert Xls/Xlsx File to Csv
        ## inputFile:
        takes str argument and will attempt to convert only this xls/xlsx file.
        ## outputDir:
        takes str argument and will attempt to exprot the csv files converted to the
        directory corresponding to the str
        """

    def convertXlsFileToCsv(self, inputFile: str, outPutDir: str) -> None:
        pass

    def __checkIfPathIsDir(self, dirPath: str) -> bool:
        return os.path.isdir(dirPath)

    def __checkIfFileExists(self, filePath: str) -> bool:
        return os.path.exists(filePath)

    def __checkIfFileIsXlFile(self, filePath: str) -> bool:
        extension = os.path.splitext(filePath)[1]
        return (extension == ".xls" or extension == ".xlsx")

   
    