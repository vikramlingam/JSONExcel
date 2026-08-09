class JsonExcelError(Exception):
    """Base exception for jsonexcel."""


class InvalidInputError(JsonExcelError):
    """Input cannot be interpreted as records."""


class ConversionError(JsonExcelError):
    """Workbook conversion failed safely."""


class ExcelLimitError(JsonExcelError):
    """Data exceeds an Excel worksheet limit."""


class LongTextError(ExcelLimitError):
    """A cell value exceeds Excel's 32,767-character text limit."""


class ConfigurationError(JsonExcelError):
    """An option is invalid or incompatible."""
