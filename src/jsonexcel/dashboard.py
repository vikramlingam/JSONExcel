"""Small native-Excel dashboard generator."""

from collections import defaultdict
from io import BytesIO
from pathlib import Path
from typing import Any

import xlsxwriter

from .exceptions import ConfigurationError, ExcelLimitError
from .inference import prepare
from .normalize import flatten, load_records
from .writer import MAX_COLS, MAX_ROWS, write_safe_cell

DASHBOARD_GROUP_START_ROW = 7


def dashboard(data: Any, output: Any = None, *, metrics: list[str] | None = None, group_by: str | None = None, date: str | None = None, charts: list[str] | None = None) -> bytes | None:
    rows = [flatten(row) for row in load_records(data)]
    requested_charts = charts or (["line"] if date and not group_by else ["bar"] if group_by or date else [])
    invalid_charts = [chart for chart in requested_charts if chart not in {"bar", "line"}]
    if invalid_charts:
        raise ConfigurationError(f"Unsupported dashboard chart type: {invalid_charts[0]}. Use 'bar' or 'line'.")
    fields = _fields(rows)
    if len(rows) + 1 > MAX_ROWS:
        raise ExcelLimitError(
            f"Dashboard Data worksheet requires {len(rows) + 1:,} rows including the header, "
            f"exceeding Excel's {MAX_ROWS:,}-row limit."
        )
    if len(fields) > MAX_COLS:
        raise ExcelLimitError(
            f"Dashboard Data worksheet discovered {len(fields):,} columns, "
            f"exceeding Excel's {MAX_COLS:,}-column limit."
        )
    candidates = metrics if metrics is not None else fields
    numeric = [field for field in candidates if any(isinstance(row.get(field), (int, float)) and not isinstance(row.get(field), bool) for row in rows)][:3]
    grouping = group_by or date
    groups = _ordered([row.get(grouping) for row in rows]) if grouping else []
    if grouping and numeric:
        grouped_rows_required = DASHBOARD_GROUP_START_ROW + 1 + len(groups)
        if grouped_rows_required > MAX_ROWS:
            raise ExcelLimitError(
                f"Dashboard grouped section requires {grouped_rows_required:,} rows, "
                f"exceeding Excel's {MAX_ROWS:,}-row worksheet limit."
            )
    buffer = BytesIO()
    workbook = xlsxwriter.Workbook(buffer, {"in_memory": True})
    dashboard_sheet = workbook.add_worksheet("Dashboard")
    data_sheet = workbook.add_worksheet("Data")
    title = workbook.add_format({"bold": True, "font_size": 16, "font_color": "#1F4E78"})
    kpi = workbook.add_format({"bold": True, "font_size": 14, "bg_color": "#D9EAF7", "align": "center"})
    date_fmts: dict[str, Any] = {}
    write_safe_cell(dashboard_sheet, 0, 0, "jsonexcel dashboard", title, sheet_name="Dashboard", header="title")
    for index, metric in enumerate(numeric):
        values: list[float | int] = [value for row in rows if isinstance((value := row.get(metric)), (int, float)) and not isinstance(value, bool)]
        write_safe_cell(dashboard_sheet, 2, index * 2, metric, kpi, sheet_name="Dashboard", header="metric")
        write_safe_cell(dashboard_sheet, 3, index * 2, sum(values) if values else 0, kpi, sheet_name="Dashboard", header=metric)
    for col_index, field in enumerate(fields):
        write_safe_cell(data_sheet, 0, col_index, field, sheet_name="Data", header=str(field))
    for row_index, row in enumerate(rows, start=1):
        for col_index, field in enumerate(fields):
            _write_prepared(workbook, data_sheet, row_index, col_index, row.get(field), date_fmts, "Data", field)
    if grouping and numeric:
        grouped: defaultdict[Any, defaultdict[str, float]] = defaultdict(lambda: defaultdict(float))
        for row in rows:
            for metric in numeric:
                value = row.get(metric)
                if isinstance(value, (int, float)) and not isinstance(value, bool):
                    grouped[row.get(grouping)][metric] += value
        start = DASHBOARD_GROUP_START_ROW
        for col_index, value in enumerate([grouping, *numeric]):
            write_safe_cell(dashboard_sheet, start, col_index, value, sheet_name="Dashboard", header="grouped header")
        for row_index, group in enumerate(groups, start=start + 1):
            _write_prepared(workbook, dashboard_sheet, row_index, 0, group, date_fmts, "Dashboard", grouping)
            for col_index, metric in enumerate(numeric, start=1):
                write_safe_cell(dashboard_sheet, row_index, col_index, grouped[group][metric], sheet_name="Dashboard", header=metric)
        for chart_type in requested_charts:
            chart = workbook.add_chart({"type": chart_type})
            chart.add_series({"name": ["Dashboard", start, 1], "categories": ["Dashboard", start + 1, 0, start + len(groups), 0], "values": ["Dashboard", start + 1, 1, start + len(groups), 1]})
            chart.set_title({"name": "Dashboard chart"})
            chart.set_legend({"none": True})
            dashboard_sheet.insert_chart("A12", chart, {"x_scale": 1.3, "y_scale": 1.2})
    workbook.close()
    content = buffer.getvalue()
    if output is None:
        return content
    if hasattr(output, "write"):
        output.write(content)
    else:
        Path(output).write_bytes(content)
    return None


def _fields(rows: list[dict[str, Any]]) -> list[str]:
    return list(dict.fromkeys(field for row in rows for field in row))


def _ordered(values: list[Any]) -> list[Any]:
    return list(dict.fromkeys(values))


def _write_prepared(workbook: Any, worksheet: Any, row: int, column: int, value: Any, formats: dict[str, Any], sheet_name: str, header: str) -> None:
    prepared, format_code = prepare(value)
    cell_format = formats.setdefault(format_code, workbook.add_format({"num_format": format_code})) if format_code else None
    write_safe_cell(worksheet, row, column, prepared, cell_format, sheet_name=sheet_name, header=header)
