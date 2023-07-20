from XlsToCsvExceptions import *
import subprocess
import os

class XlsConverter:
    """
    ## XlsConverter class

    ### How to use?
    Instantiate an object of this class to gain access to conversion methods, you can choose to exclude a excelPath argument in the constructor.
    Excluding excelPath will make XlsConverter try to find your excel installation by trying from a small set of options. As you may understand this is not very safe so 
    providing your own path is recommended.
    """
    excelPath: str

    def __init__(self, excelPath: str = "AutoFind") -> None:
        if self.excelPath == "AutoFind":
            self.excelPath = self.__findExcelWindows()
        else:
            self.excelPath = excelPath

    def convertXlsDirToCsv(self, inputDir: str, outputDir: str) -> None:
        """
        ## Convert Xls Dir to Csv
        ### inputDir: 
        takes str as argument and will attempt to convert all .xlsx and .xls
        files in the directory corresponding to the str.
        ### outputDir: 
        takes str as argument and will attempt to export the csv files converted
        to the directory corresponding to the str.
        ## Errors
        """
        if self.__checkIfPathIsDir(self, inputDir) is False:
            InputIsNotDirException()

        excelPath: str = self.excelPath
        scriptPath: str = "FastXlsToCsv\basScripts\ConvertXlsDir.bas"
        arg1: str = inputDir
        arg2: str = outputDir

        cmd = [
        excelPath,
        "/e",                   # Start Excel without displaying the UI
        "/NoSplash",            # Disable the splash screen
        "/x", scriptPath,       # Path to your VBA script
        arg1,                   # First argument
        arg2                    # Second argument
        ]
        subprocess.run(cmd)

        
    def convertXlsFileToCsv(self, inputFile: str, outPutDir: str) -> None:
        pass

    def __checkIfPathIsDir(self, dirPath: str) -> bool:
        return os.path.isdir(dirPath)

    def __checkIfFileExists(self, filePath: str) -> bool:
        return os.path.exists(filePath)

    def __checkIfFileIsXlFile(self, filePath: str) -> bool:
        extension = os.path.splitext(filePath)[1]
        return (extension == ".xls" or extension == ".xlsx")

    def __findExcelWindows(self):
        "Tries to find Excel on windows, raises error if it doesnt, refuses to elaborate."

        possible_paths = [
            "C:\\Program Files\\Microsoft Office\\root\\Office16\\EXCEL.EXE",
            "C:\\Program Files (x86)\\Microsoft Office\\root\\Office16\\EXCEL.EXE",
        ]

        for path in possible_paths:
            if os.path.exists(path):
                return path

        raise AutoFindExcelException()
    