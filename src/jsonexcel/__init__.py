"""Simple, safe JSON-to-Excel conversion."""

from .api import convert, convert_folder, to_bytes, to_excel
from .analytics import analyze, pivot, summarize
from .compare import compare
from .dashboard import dashboard
from .exceptions import ConfigurationError, ConversionError, ExcelLimitError, InvalidInputError, JsonExcelError, LongTextError
from .excel_reader import inspect_json_column
from .limits import diagnose_limits
from .profile import profile, schema
from .roundtrip import from_excel
from .validation import inspect, validate

__all__ = ["convert", "to_excel", "to_bytes", "convert_folder", "analyze", "summarize", "pivot", "compare", "dashboard", "profile", "schema", "from_excel", "diagnose_limits", "inspect", "inspect_json_column", "validate", "JsonExcelError", "InvalidInputError", "ConversionError", "ExcelLimitError", "LongTextError", "ConfigurationError"]
__version__ = "0.1.0"
