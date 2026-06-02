"""
PBIP TMDL Injector - ThoughtSpot -> Power BI
===========================================

Use this file as:
    backend/src/powerbi/pbip_tmdl_injector.py

Polished version:
- Uses safe measure table name: "DAX Measures".
- Never creates unsupported table named "Measures".
- Deduplicates tables, columns, and measures.
- Generates Power BI-friendly TMDL tables.
- Adds one safe placeholder row when source data is empty.
- Keeps numeric columns numeric so Power BI can use Sum/Average, not only Count.
- Adds better summarizeBy defaults:
    Sales / Profit / Cost / Quantity / Target -> sum
    Discount / Rate / Margin / Percentage -> average
    ID / Name / Category / Region -> none
- Adds measure formatting:
    Sales/Profit/Cost/Target -> currency style
    Margin/Rate/Discount/Percentage -> percentage style
    Count/Quantity -> whole number style
- Converts raw ThoughtSpot formulas into safer Power BI DAX.
- Avoids fake fallback table names like SourceTable.
- Writes relationships.tmdl and updates model.tmdl idempotently.
"""

from __future__ import annotations

import math
import re
import uuid
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Tuple

from loguru import logger

try:
    import pandas as pd
    _PANDAS_AVAILABLE = True
except ImportError:  # pragma: no cover
    pd = None
    _PANDAS_AVAILABLE = False


MEASURE_TABLE_NAME = "DAX Measures"
INVALID_MEASURE_TABLE_NAMES = {"Measures", "measures", "DAX_Measures"}


# -----------------------------------------------------------------------------
# Name helpers
# -----------------------------------------------------------------------------

def _clean_name(name: Any, fallback: str = "Field") -> str:
    value = str(name or "").strip()
    value = re.sub(r'["`]', "", value)
    value = re.sub(r"\s+", " ", value).strip()
    return value or fallback


def _safe_filename(name: Any) -> str:
    safe = re.sub(r'[<>:"/\\|?*]', "_", _clean_name(name, "Table"))
    safe = safe.strip(" .")
    return safe or "Table"


def _escape_tmdl_quoted(value: Any) -> str:
    return str(value).replace("'", "''")


def _tmdl_identifier(name: Any, fallback: str = "Field") -> str:
    return f"'{_escape_tmdl_quoted(_clean_name(name, fallback))}'"


def _escape_dax_string(value: Any) -> str:
    return str(value).replace('"', '""')


def _escape_dax_single_quoted(value: Any) -> str:
    return str(value).replace("'", "''")


def _dax_column_ref(table: str, column: str) -> str:
    return f"'{_escape_dax_single_quoted(table)}'[{_clean_name(column, 'Column')}]"


def _dax_sum_ref(table: str, column: str) -> str:
    return f"SUM({_dax_column_ref(table, column)})"


def _dax_avg_ref(table: str, column: str) -> str:
    return f"AVERAGE({_dax_column_ref(table, column)})"


def _dedupe_names(names: Iterable[Any], fallback: str = "Column") -> List[str]:
    seen: Dict[str, int] = {}
    result: List[str] = []

    for raw in names:
        base = _clean_name(raw, fallback)
        key = base.lower()
        seen[key] = seen.get(key, 0) + 1
        result.append(base if seen[key] == 1 else f"{base}_{seen[key]}")

    return result


def _clean_table_name(name: Any, index: int = 1) -> str:
    table = _clean_name(name, f"Table_{index}")
    if table in INVALID_MEASURE_TABLE_NAMES:
        table = MEASURE_TABLE_NAME if "dax" in table.lower() else f"{table} Table"
    return table


def _clean_measure_name(name: Any, fallback: str = "Measure") -> str:
    return _clean_name(name, fallback)


# -----------------------------------------------------------------------------
# Type / formatting helpers
# -----------------------------------------------------------------------------

def _looks_numeric_column(column_name: str) -> bool:
    """Infer whether a column is a numeric business metric.

    This is intentionally conservative for ID/name/date/category fields, because
    Power BI should group by those fields. It is intentionally aggressive for
    Sales/Profit/Cost/Quantity/Discount even when pandas reads them as object,
    because otherwise Power BI can show only Count of Sales.
    """
    low = str(column_name or "").lower().strip()
    if not low:
        return False

    text_tokens = [
        " id", "_id", "id ", "key", "code", "name", "category",
        "segment", "region", "country", "state", "city", "date",
        "product", "customer", "order date",
    ]
    if any(t in low for t in text_tokens):
        # Keep explicit numeric count fields numeric, but do not make IDs numeric.
        if not any(t in low for t in ["count", "quantity", "qty", "units"]):
            return False

    numeric_tokens = [
        "sales", "profit", "revenue", "cost", "amount", "price",
        "quantity", "qty", "discount", "margin", "rate", "target",
        "score", "value", "total", "avg", "average", "count", "units"
    ]
    return any(t in low for t in numeric_tokens)


def _is_percentage_column(column_name: str) -> bool:
    low = column_name.lower()
    return any(token in low for token in ["discount", "margin", "rate", "percent", "percentage", "%"])


def _is_whole_number_column(column_name: str) -> bool:
    low = column_name.lower()
    return any(token in low for token in ["quantity", "qty", "count", "units", "number"])


def _infer_tmdl_datatype_from_name(column_name: str) -> str:
    low = column_name.lower()
    if any(token in low for token in ["date", "month", "year"]) and "updated" not in low:
        return "dateTime"
    if _is_whole_number_column(column_name):
        return "int64"
    if _looks_numeric_column(column_name):
        return "double"
    return "string"


def _summarize_by_for_column(column_name: str, tmdl_type: str) -> str:
    low = column_name.lower()

    if tmdl_type in {"string", "dateTime", "boolean"}:
        return "none"

    if any(token in low for token in ["id", "key", "code", "number"]) and not any(
        token in low for token in ["order count", "count", "quantity", "sales", "profit", "cost"]
    ):
        return "none"

    if _is_percentage_column(column_name):
        return "average"

    return "sum"


def _format_string_for_measure(measure_name: str, dax: str = "") -> str:
    low = f"{measure_name} {dax}".lower()

    if any(token in low for token in ["margin", "rate", "discount", "percent", "percentage", "achievement"]):
        return "0.00%"

    if any(token in low for token in ["count", "quantity", "qty", "orders", "units"]):
        return "#,##0"

    if any(token in low for token in ["sales", "profit", "revenue", "cost", "amount", "price", "target", "value"]):
        return "₹#,##0.00"

    return "#,##0.00"


def _auto_measure_name_for_column(column_name: str) -> str:
    base = _clean_name(str(column_name).replace("_", " "), "Value")
    if _is_percentage_column(base):
        return f"Average {base}"
    if any(token in base.lower() for token in ["id", "key", "code"]) and not _looks_numeric_column(base):
        return f"Distinct Count {base}"
    return f"Total {base}"


def _build_auto_numeric_measures(tables: Dict[str, "pd.DataFrame"]) -> List[Dict[str, Any]]:
    """Create DAX measures for every numeric business column.

    Power BI Desktop may still render raw numeric columns as Count in generated
    PBIP visuals. Using explicit measures like [Total Sales] removes that issue.
    """
    result: List[Dict[str, Any]] = []
    seen = set()

    if not _PANDAS_AVAILABLE:
        return result

    for table_name, df in (tables or {}).items():
        if df is None or not hasattr(df, "columns"):
            continue

        for col in df.columns:
            col_name = _clean_name(col, "Column")
            try:
                tmdl_type = _get_tmdl_datatype_for_column(col_name, df[col].dtype)
            except Exception:
                tmdl_type = _infer_tmdl_datatype_from_name(col_name)

            if tmdl_type not in {"int64", "double"}:
                continue

            # Do not create measures for ID/code/key columns.
            low = col_name.lower()
            if any(token in low for token in [" id", "_id", "key", "code"]) and not any(token in low for token in ["sales", "profit", "cost", "quantity", "amount", "revenue"]):
                continue

            measure_name = _auto_measure_name_for_column(col_name)
            key = measure_name.lower()
            if key in seen:
                continue
            seen.add(key)

            dax = _dax_avg_ref(table_name, col_name) if _is_percentage_column(col_name) else _dax_sum_ref(table_name, col_name)
            result.append({
                "name": measure_name,
                "dax": dax,
                "formatString": _format_string_for_measure(measure_name, dax),
                "autoGenerated": True,
            })

    return result


def _clean_dataframe(df: Optional["pd.DataFrame"]) -> Optional["pd.DataFrame"]:
    if df is None or not _PANDAS_AVAILABLE:
        return df

    clean_df = df.copy()
    clean_df.columns = _dedupe_names([str(c) for c in clean_df.columns], "Column")

    # Try to convert obvious numeric/date columns if they arrived as text.
    for col in clean_df.columns:
        col_name = str(col)
        if _looks_numeric_column(col_name):
            try:
                clean_df[col] = pd.to_numeric(clean_df[col], errors="ignore")
            except Exception:
                pass
        elif "date" in col_name.lower():
            try:
                clean_df[col] = pd.to_datetime(clean_df[col], errors="ignore")
            except Exception:
                pass

    # Keep generated PBIP small and stable.
    max_rows = 100
    if len(clean_df) > max_rows:
        clean_df = clean_df.head(max_rows)

    return clean_df


def _get_tmdl_datatype_for_column(column_name: str, dtype: Any) -> str:
    if not _PANDAS_AVAILABLE:
        return _infer_tmdl_datatype_from_name(column_name)

    try:
        if pd.api.types.is_bool_dtype(dtype):
            return "boolean"
        if pd.api.types.is_integer_dtype(dtype):
            return "int64"
        if pd.api.types.is_float_dtype(dtype):
            return "double"
        if pd.api.types.is_datetime64_any_dtype(dtype):
            return "dateTime"
    except Exception:
        pass

    # If pandas still says object/string but name looks numeric, force numeric.
    return _infer_tmdl_datatype_from_name(column_name)


def _dax_datatable_type(tmdl_type: str) -> str:
    return {
        "boolean": "BOOLEAN",
        "int64": "INTEGER",
        "double": "DOUBLE",
        "dateTime": "DATETIME",
        "string": "STRING",
    }.get(tmdl_type, "STRING")


def _format_cell_value(value: Any, tmdl_type: str = "string") -> str:
    if value is None:
        return "BLANK()"

    if _PANDAS_AVAILABLE:
        try:
            if pd.isna(value):
                return "BLANK()"
        except Exception:
            pass

    if tmdl_type == "boolean":
        if isinstance(value, str):
            return "TRUE()" if value.strip().lower() in {"true", "1", "yes", "y"} else "FALSE()"
        return "TRUE()" if bool(value) else "FALSE()"

    if tmdl_type in {"int64", "double"}:
        try:
            number = float(value)
            if math.isnan(number) or math.isinf(number):
                return "BLANK()"
            if tmdl_type == "int64":
                return str(int(number))
            return str(int(number)) if number.is_integer() else repr(number)
        except Exception:
            return "0"

    if tmdl_type == "dateTime":
        if _PANDAS_AVAILABLE and hasattr(value, "strftime"):
            try:
                return f"DATE({value.year}, {value.month}, {value.day})"
            except Exception:
                pass
        if isinstance(value, str):
            m = re.match(r"^(\d{4})-(\d{1,2})-(\d{1,2})", value.strip())
            if m:
                return f"DATE({int(m.group(1))}, {int(m.group(2))}, {int(m.group(3))})"
        return "DATE(1900, 1, 1)"

    return f'"{_escape_dax_string(value)}"'


def _placeholder_cell_for_type(column_name: str, tmdl_type: str) -> str:
    """
    Creates one safe dummy value per type.
    This is not customer data; it only prevents completely empty generated tables.
    """
    low = column_name.lower()

    if tmdl_type == "boolean":
        return "FALSE()"
    if tmdl_type == "int64":
        return "1" if any(token in low for token in ["quantity", "qty", "count"]) else "0"
    if tmdl_type == "double":
        if "sales" in low or "revenue" in low:
            return "1000"
        if "profit" in low:
            return "250"
        if "cost" in low:
            return "750"
        if "discount" in low or "rate" in low or "margin" in low:
            return "0.10"
        if "target" in low:
            return "5000"
        return "0"
    if tmdl_type == "dateTime":
        return "DATE(2026, 1, 1)"
    if "region" in low:
        return '"Demo Region"'
    if "category" in low:
        return '"Demo Category"'
    if "segment" in low:
        return '"Demo Segment"'
    if "name" in low:
        return '"Demo Name"'
    if "id" in low:
        return '"DEMO001"'
    return '"Placeholder"'


def _build_datatable_dax(df: Optional["pd.DataFrame"]) -> str:
    if not _PANDAS_AVAILABLE:
        return '\t\t\tDATATABLE ( "_dummy", STRING, { { "" } } )'

    if df is None or len(df.columns) == 0:
        return '\t\t\tDATATABLE ( "_dummy", STRING, { { "" } } )'

    column_types = {
        col: _get_tmdl_datatype_for_column(str(col), df[col].dtype)
        for col in df.columns
    }

    header_parts: List[str] = []
    for col in df.columns:
        tmdl_type = column_types[col]
        header_parts.append(f'"{_escape_dax_string(col)}", {_dax_datatable_type(tmdl_type)}')

    header = ",\n\t\t\t\t".join(header_parts)

    row_lines: List[str] = []
    for _, row in df.iterrows():
        values = [_format_cell_value(row[col], column_types[col]) for col in df.columns]
        row_lines.append("\t\t\t\t{ " + ", ".join(values) + " }")

    if not row_lines:
        placeholder_values = []
        for col in df.columns:
            tmdl_type = column_types[col]
            placeholder_values.append(_placeholder_cell_for_type(str(col), tmdl_type))
        row_lines.append("\t\t\t\t{ " + ", ".join(placeholder_values) + " }")

    rows = "{\n" + ",\n".join(row_lines) + "\n\t\t\t\t}"

    return (
        "\t\t\tDATATABLE (\n"
        f"\t\t\t\t{header},\n"
        f"\t\t\t\t{rows}\n"
        "\t\t\t)"
    )


# -----------------------------------------------------------------------------
# DAX helpers
# -----------------------------------------------------------------------------

def _mask_dax_protected_parts(expr: str) -> Tuple[str, Dict[str, str]]:
    protected: Dict[str, str] = {}
    masked = str(expr or "")

    patterns = [
        r"'(?:[^']|'')*'\s*\[[^\]]+\]",  # 'Table'[Column]
        r"\[[^\]]+\]",                   # [Measure] or [Column]
        r'"(?:[^"]|"")*"',               # "string"
    ]

    counter = 0
    for pattern in patterns:
        while True:
            match = re.search(pattern, masked)
            if not match:
                break
            key = f"__DAX_PROTECTED_{counter}__"
            protected[key] = match.group(0)
            masked = masked[:match.start()] + key + masked[match.end():]
            counter += 1

    return masked, protected


def _unmask_dax_protected_parts(expr: str, protected: Dict[str, str]) -> str:
    for key, value in protected.items():
        expr = expr.replace(key, value)
    return expr


def _replace_raw_columns_with_aggregation(
    expr: str,
    default_table: Optional[str],
    default_columns: Optional[List[str]] = None,
) -> str:
    if not default_table or not default_columns:
        return expr

    table = _clean_name(default_table, "Table")
    masked, protected = _mask_dax_protected_parts(expr)

    columns = sorted(
        [_clean_name(c, "Column") for c in default_columns if str(c).strip()],
        key=len,
        reverse=True,
    )

    dax_reserved_words = {
        "SUM", "AVERAGE", "MIN", "MAX", "COUNT", "DISTINCTCOUNT",
        "DIVIDE", "IF", "SWITCH", "CALCULATE", "FILTER", "ALL", "ALLEXCEPT",
        "BLANK", "TRUE", "FALSE", "DATE", "YEAR", "MONTH", "DAY",
        "AND", "OR", "NOT", "ROUND", "ROUNDUP", "ROUNDDOWN", "ABS",
        "VAR", "RETURN", "IN", "VALUE", "FORMAT", "CONTAINSSTRING",
    }

    for col in columns:
        if not col or col.upper() in dax_reserved_words:
            continue

        pattern = rf"(?<![\w\]])\b{re.escape(col)}\b(?![\w\[]|\s*\])"
        replacement = _dax_avg_ref(table, col) if _is_percentage_column(col) else _dax_sum_ref(table, col)
        masked = re.sub(pattern, replacement, masked, flags=re.IGNORECASE)

    return _unmask_dax_protected_parts(masked, protected)


def _convert_common_thoughtspot_functions(expr: str, default_table: Optional[str]) -> str:
    if not default_table:
        return expr

    table = _clean_name(default_table, "Table")

    def replace_agg(match: re.Match) -> str:
        fn = match.group(1).lower()
        field = _clean_name(match.group(2), "Column")
        ref = _dax_column_ref(table, field)

        if fn in {"avg", "average"}:
            return f"AVERAGE({ref})"
        if fn in {"count", "countd", "count_distinct", "distinct_count"}:
            return f"DISTINCTCOUNT({ref})" if fn != "count" else f"COUNT({ref})"
        return f"{fn.upper()}({ref})"

    # sum(Sales), average(Discount), count(Order ID)
    expr = re.sub(
        r"\b(sum|avg|average|min|max|count|countd|count_distinct|distinct_count)\s*\(\s*([A-Za-z_][A-Za-z0-9_ ]*)\s*\)",
        replace_agg,
        expr,
        flags=re.IGNORECASE,
    )

    # IF stays IF, but normalize lower-case if.
    expr = re.sub(r"\bif\s*\(", "IF(", expr, flags=re.IGNORECASE)
    return expr


def _normalize_dax_table_references(
    dax: str,
    default_table: Optional[str] = None,
    default_columns: Optional[List[str]] = None,
) -> str:
    """Cleanup common generated DAX patterns and fix raw column references."""
    expr = str(dax or "0").strip() or "0"

    if default_table:
        table = _clean_name(default_table, "Table")

        expr = _convert_common_thoughtspot_functions(expr, table)

        # Convert SUM([Sales]) -> SUM('orders_fact'[Sales])
        expr = re.sub(
            r"\b(SUM|AVERAGE|MIN|MAX|COUNT|DISTINCTCOUNT)\s*\(\s*\[([^\]]+)\]\s*\)",
            lambda m: f"{m.group(1).upper()}({_dax_column_ref(table, m.group(2))})",
            expr,
            flags=re.IGNORECASE,
        )

        # Convert raw fields like Sales - Cost.
        expr = _replace_raw_columns_with_aggregation(expr, table, default_columns)

    simple_div = re.match(r"^(.+?)\s*/\s*(.+?)$", expr, flags=re.DOTALL)
    if simple_div and "DIVIDE" not in expr.upper():
        left = simple_div.group(1).strip()
        right = simple_div.group(2).strip()
        if left and right:
            expr = f"DIVIDE({left}, {right}, 0)"

    return expr


# -----------------------------------------------------------------------------
# TMDL builders
# -----------------------------------------------------------------------------

def _build_table_tmdl(table_name: str, df: Optional["pd.DataFrame"]) -> str:
    table_name = _clean_name(table_name, "Table")
    df = _clean_dataframe(df)

    lines: List[str] = []
    lines.append(f"table {_tmdl_identifier(table_name, 'Table')}")
    lines.append(f"\tlineageTag: {uuid.uuid4()}")
    lines.append("")

    if df is not None and _PANDAS_AVAILABLE:
        for col in df.columns:
            col_name = _clean_name(col, "Column")
            dtype = _get_tmdl_datatype_for_column(col_name, df[col].dtype)
            summarize_by = _summarize_by_for_column(col_name, dtype)
            lines.append(f"\tcolumn {_tmdl_identifier(col_name, 'Column')}")
            lines.append(f"\t\tdataType: {dtype}")
            lines.append(f"\t\tlineageTag: {uuid.uuid4()}")
            lines.append(f"\t\tsummarizeBy: {summarize_by}")
            lines.append(f"\t\tsourceColumn: {_clean_name(col_name, 'Column')}")
            lines.append("")

    lines.append(f"\tpartition {_tmdl_identifier(table_name, 'Table')} = calculated")
    lines.append("\t\tmode: import")
    lines.append("\t\texpression =")
    lines.append(_build_datatable_dax(df))
    lines.append("")
    return "\n".join(lines)


def _build_measure_table_tmdl(
    measures: List[Dict[str, Any]],
    default_table: Optional[str],
    default_columns: Optional[List[str]] = None,
) -> str:
    lines: List[str] = []
    lines.append(f"table {_tmdl_identifier(MEASURE_TABLE_NAME, 'Table')}")
    lines.append(f"\tlineageTag: {uuid.uuid4()}")
    lines.append("")

    lines.append("\tcolumn '_MeasureTableDummy'")
    lines.append("\t\tdataType: string")
    lines.append(f"\t\tlineageTag: {uuid.uuid4()}")
    lines.append("\t\tsummarizeBy: none")
    lines.append("\t\tsourceColumn: _MeasureTableDummy")
    lines.append("\t\tisHidden")
    lines.append("")

    seen = set()
    for raw in measures or []:
        name = _clean_measure_name(
            raw.get("name")
            or raw.get("measure")
            or raw.get("calculated_field")
            or raw.get("source_calculated_field"),
            "Measure",
        )
        dax = str(
            raw.get("dax")
            or raw.get("expression")
            or raw.get("dax_formula")
            or raw.get("converted_dax_formula")
            or "0"
        ).strip() or "0"
        dax = _normalize_dax_table_references(dax, default_table, default_columns)
        fmt = str(raw.get("formatString") or raw.get("format_string") or _format_string_for_measure(name, dax))

        key = name.lower()
        if key in seen:
            logger.warning(f"Skipping duplicate measure: {name}")
            continue
        seen.add(key)

        expr_lines = dax.splitlines()
        if len(expr_lines) == 1:
            lines.append(f"\tmeasure {_tmdl_identifier(name, 'Measure')} = {expr_lines[0]}")
        else:
            lines.append(f"\tmeasure {_tmdl_identifier(name, 'Measure')} =")
            for expr_line in expr_lines:
                lines.append(f"\t\t\t{expr_line}")
        lines.append(f"\t\tformatString: {fmt}")
        lines.append(f"\t\tlineageTag: {uuid.uuid4()}")
        lines.append("")

    lines.append(f"\tpartition {_tmdl_identifier(MEASURE_TABLE_NAME, 'Table')} = calculated")
    lines.append("\t\tmode: import")
    lines.append("\t\texpression =")
    lines.append('\t\t\tDATATABLE ( "_MeasureTableDummy", STRING, { { "" } } )')
    lines.append("")
    return "\n".join(lines)


def _relationship_name(rel: Dict[str, Any]) -> str:
    return _clean_name(
        rel.get("relationship_id")
        or rel.get("name")
        or rel.get("id")
        or f"rel_{rel.get('source_table') or rel.get('from_table')}_{rel.get('target_table') or rel.get('to_table')}",
        "Relationship",
    )


def _build_relationship_tmdl(rel: Dict[str, Any]) -> Optional[str]:
    source_table = _clean_name(rel.get("source_table") or rel.get("from_table"), "")
    target_table = _clean_name(rel.get("target_table") or rel.get("to_table"), "")
    source_column = _clean_name(rel.get("source_column") or rel.get("from_column"), "")
    target_column = _clean_name(rel.get("target_column") or rel.get("to_column"), "")

    if not all([source_table, target_table, source_column, target_column]):
        return None

    name = _relationship_name(rel)
    return (
        f"relationship {_tmdl_identifier(name, 'Relationship')}\n"
        f"\tfromColumn: {_tmdl_identifier(source_table, 'Table')}.{_tmdl_identifier(source_column, 'Column')}\n"
        f"\ttoColumn: {_tmdl_identifier(target_table, 'Table')}.{_tmdl_identifier(target_column, 'Column')}\n"
        "\tcrossFilteringBehavior: oneDirection\n"
        "\tisActive: true\n"
    )


# -----------------------------------------------------------------------------
# Main injector
# -----------------------------------------------------------------------------

class PBIPTmdlInjector:
    """
    Injects tables, measures, and optional relationships into a PBIP SemanticModel.

    Expected folder:
        <ProjectName>.SemanticModel/definition/model.tmdl
        <ProjectName>.SemanticModel/definition/tables/
    """

    def inject(
        self,
        sm_folder: Path,
        tables: Optional[Dict[str, "pd.DataFrame"]] = None,
        measures: Optional[List[Dict[str, Any]]] = None,
        relationships: Optional[List[Dict[str, Any]]] = None,
    ) -> List[str]:
        sm_folder = Path(sm_folder)
        definition_dir = sm_folder / "definition"
        tables_dir = definition_dir / "tables"
        tables_dir.mkdir(parents=True, exist_ok=True)
        definition_dir.mkdir(parents=True, exist_ok=True)

        injected: List[str] = []

        self._remove_invalid_old_measure_files(tables_dir)

        clean_tables = self._prepare_tables(tables or {})
        default_table = self._choose_default_fact_table(clean_tables)
        default_columns = self._get_columns_for_table(clean_tables, default_table)

        for table_name, df in clean_tables.items():
            try:
                content = _build_table_tmdl(table_name, df)
                path = tables_dir / f"{_safe_filename(table_name)}.tmdl"
                path.write_text(content, encoding="utf-8")
                injected.append(table_name)
                row_count = len(df) if df is not None and hasattr(df, "__len__") else 0
                logger.info(f"✓ Wrote polished table TMDL: {path.name} ({row_count} sample rows)")
            except Exception as exc:
                logger.exception(f"✗ Failed to write table TMDL for {table_name}: {exc}")

        # Always add explicit measures for numeric columns (Total Sales, Total Profit, etc.).
        # This is the reliable fix for visuals showing Count of Sales.
        valid_measures = self._prepare_measures((measures or []) + _build_auto_numeric_measures(clean_tables))
        if valid_measures:
            measure_path = tables_dir / f"{_safe_filename(MEASURE_TABLE_NAME)}.tmdl"
            measure_path.write_text(
                _build_measure_table_tmdl(valid_measures, default_table, default_columns),
                encoding="utf-8",
            )
            injected.append(MEASURE_TABLE_NAME)
            logger.info(f"✓ Wrote polished measure table: {measure_path.name} ({len(valid_measures)} measures)")
        else:
            logger.warning("No valid measures found; DAX Measures table skipped")

        self._write_relationships(definition_dir, relationships or [])
        self._update_model_tmdl(sm_folder, injected, relationships or [])
        return injected

    def _prepare_tables(self, tables: Dict[str, "pd.DataFrame"]) -> Dict[str, "pd.DataFrame"]:
        result: Dict[str, "pd.DataFrame"] = {}
        used = set()

        for idx, (raw_name, raw_df) in enumerate(tables.items(), start=1):
            base = _clean_table_name(raw_name, idx)
            name = base
            suffix = 2
            while name.lower() in used:
                name = f"{base}_{suffix}"
                suffix += 1
            used.add(name.lower())
            result[name] = _clean_dataframe(raw_df)

        return result

    def _prepare_measures(self, measures: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        result: List[Dict[str, Any]] = []
        seen = set()

        for raw in measures or []:
            name = _clean_measure_name(
                raw.get("name")
                or raw.get("measure")
                or raw.get("calculated_field")
                or raw.get("source_calculated_field"),
                "",
            )
            dax = str(
                raw.get("dax")
                or raw.get("expression")
                or raw.get("dax_formula")
                or raw.get("converted_dax_formula")
                or ""
            ).strip()

            if not name or not dax:
                continue

            key = name.lower()
            if key in seen:
                continue
            seen.add(key)

            fixed = dict(raw)
            fixed["name"] = name
            fixed["dax"] = dax
            fixed["formatString"] = raw.get("formatString") or raw.get("format_string") or _format_string_for_measure(name, dax)
            result.append(fixed)

        return result

    def _choose_default_fact_table(self, tables: Dict[str, "pd.DataFrame"]) -> Optional[str]:
        if not tables:
            return None

        priority_tokens = ["fact", "orders", "sales", "transaction", "revenue"]
        for token in priority_tokens:
            for name in tables:
                if token in name.lower():
                    return name

        return next(iter(tables.keys()))

    def _get_columns_for_table(
        self,
        tables: Dict[str, "pd.DataFrame"],
        table_name: Optional[str],
    ) -> List[str]:
        if not table_name or table_name not in tables:
            return []

        df = tables.get(table_name)
        if df is None or not hasattr(df, "columns"):
            return []

        return [str(c) for c in df.columns]

    def _remove_invalid_old_measure_files(self, tables_dir: Path) -> None:
        for invalid_name in INVALID_MEASURE_TABLE_NAMES:
            path = tables_dir / f"{invalid_name}.tmdl"
            if path.exists():
                path.unlink()
                logger.info(f"Removed invalid old measure table file: {path.name}")

    def _write_relationships(self, definition_dir: Path, relationships: List[Dict[str, Any]]) -> None:
        rel_blocks: List[str] = []
        seen = set()

        for rel in relationships or []:
            source_table = _clean_name(rel.get("source_table") or rel.get("from_table"), "")
            target_table = _clean_name(rel.get("target_table") or rel.get("to_table"), "")
            source_column = _clean_name(rel.get("source_column") or rel.get("from_column"), "")
            target_column = _clean_name(rel.get("target_column") or rel.get("to_column"), "")
            key = (source_table.lower(), source_column.lower(), target_table.lower(), target_column.lower())
            if not all(key) or key in seen:
                continue
            seen.add(key)

            block = _build_relationship_tmdl(rel)
            if block:
                rel_blocks.append(block)

        if not rel_blocks:
            return

        rel_path = definition_dir / "relationships.tmdl"
        rel_path.write_text("\n".join(rel_blocks).strip() + "\n", encoding="utf-8")
        logger.info(f"✓ Wrote relationships.tmdl ({len(rel_blocks)} relationships)")

    def _update_model_tmdl(
        self,
        sm_folder: Path,
        table_names: List[str],
        relationships: List[Dict[str, Any]],
    ) -> None:
        model_tmdl = Path(sm_folder) / "definition" / "model.tmdl"
        model_tmdl.parent.mkdir(parents=True, exist_ok=True)

        if model_tmdl.exists():
            content = model_tmdl.read_text(encoding="utf-8")
        else:
            content = "model Model\n\tculture: en-US\n"

        refs_to_add: List[str] = []

        for name in table_names:
            ref = f"ref table {_tmdl_identifier(name, 'Table')}"
            if ref not in content and f"ref table {name}" not in content:
                refs_to_add.append(ref)

        if relationships:
            rel_ref = "ref relationships"
            if rel_ref not in content:
                refs_to_add.append(rel_ref)

        if refs_to_add:
            content = content.rstrip() + "\n\n" + "\n".join(refs_to_add) + "\n"
            model_tmdl.write_text(content, encoding="utf-8")
            logger.info(f"✓ Updated model.tmdl with {len(refs_to_add)} references")
        else:
            model_tmdl.write_text(content, encoding="utf-8")
            logger.info("✓ model.tmdl already contains all required references")
