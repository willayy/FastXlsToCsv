class InputIsNotDirException(Exception):
    def __init__(self, message = "The provided input tp inputDir is not a str path to a directory") -> None:
        super().__init__(message)

class InputIsNotFileException(Exception):
    def __init__(self, message = "The provided input to inputFile is not a str path to a directory") -> None:
        super().__init__(message)

class OutputIsNotDirException(Exception):
    def __init__(self, message = "The provided input to outputDir is not a str path to a directory ") -> None:
        super().__init__(message)

class FastXlsToCsvModuleError(Exception):
    def __init__(self, message = "There was a script error from within the FastXlsToCsv module, please check for bad input and then report to module developer") -> None:
        super().__init__(message)