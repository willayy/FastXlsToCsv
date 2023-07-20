class InputIsNotDirException(Exception):
    def __init__(self, message = "The provided input tp inputDir is not a str path to a directory") -> None:
        super().__init__(message)

class InputIsNotFileException(Exception):
    def __init__(self, message = "The provided input to inputFile is not a str path to a directory") -> None:
        super().__init__(message)

class OutputIsNotDirException(Exception):
    def __init__(self, message = "The provided input to outputDir is not a str path to a directory ") -> None:
        super().__init__(message)

class AutoFindExcelException(Exception):
    def __init__(self, message = "No path to excel was found when XlsConverter tried to look in normal installation locations. Please provide a path to EXCEL.EXE") -> None:
        super().__init__(*args)