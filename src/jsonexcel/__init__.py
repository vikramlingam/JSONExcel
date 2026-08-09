"""Simple, safe JSON-to-Excel conversion."""

from .analytics import analyze, pivot, summarize
from .api import convert, convert_folder, to_bytes, to_excel
from .compare import compare
from .dashboard import dashboard
from .excel_reader import inspect_json_column
from .exceptions import (
    ConfigurationError,
    ConversionError,
    ExcelLimitError,
    InvalidInputError,
    JsonExcelError,
    LongTextError,
)
from .limits import diagnose_limits
from .profile import profile, schema
from .roundtrip import from_excel
from .validation import inspect, validate

__all__ = ["ConfigurationError", "ConversionError", "ExcelLimitError", "InvalidInputError", "JsonExcelError", "LongTextError", "analyze", "compare", "convert", "convert_folder", "dashboard", "diagnose_limits", "from_excel", "inspect", "inspect_json_column", "pivot", "profile", "schema", "summarize", "to_bytes", "to_excel", "validate"]
__version__ = "0.1.0"
