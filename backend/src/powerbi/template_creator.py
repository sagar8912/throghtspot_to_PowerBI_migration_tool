"""
Power BI template/report helper for ThoughtSpot -> Power BI migration.

Use this file as:
    backend/src/powerbi/template_creator.py

Polish version:
- Keeps safe measure table name: "DAX Measures".
- Creates stable report scaffolding without forcing Power BI upgrade.
- Creates clean README and visual_suggestions.json for demo.
- Keeps backward-compatible StarterPBIXCreator.
- Uses VisualConverter when available, but falls back safely if visuals fail.
- Does not expose customer data.
"""

from __future__ import annotations

import json
import re
import tempfile
import zipfile
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional

try:
    from loguru import logger
except Exception:  # pragma: no cover
    import logging
    logging.basicConfig(level=logging.INFO)
    logger = logging.getLogger(__name__)


SAFE_MEASURE_TABLE_NAME = "DAX Measures"
DEFAULT_PAGE_NAME = "Executive Overview"
DEFAULT_REPORT_WIDTH = 1280
DEFAULT_REPORT_HEIGHT = 720


# -----------------------------------------------------------------------------
# Generic helpers
# -----------------------------------------------------------------------------

def _safe_name(value: Any, fallback: str = "Item") -> str:
    text = str(value or fallback).strip()
    text = re.sub(r"[\r\n\t]+", " ", text)
    text = re.sub(r"[^A-Za-z0-9_ .%₹$&()/-]+", "_", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text or fallback


def _safe_file_name(value: Any, fallback: str = "item") -> str:
    text = _safe_name(value, fallback=fallback)
    text = text.replace(" ", "_")
    text = re.sub(r"[^A-Za-z0-9_.-]+", "_", text)
    return text[:90] or fallback


def _write_json(path: Path, data: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")


def _write_text(path: Path, text: str) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(text, encoding="utf-8")


def _normalise_list(value: Any) -> List[Dict[str, Any]]:
    if not value:
        return []
    if isinstance(value, dict):
        return [v for v in value.values() if isinstance(v, dict)]
    if isinstance(value, list):
        return [v for v in value if isinstance(v, dict)]
    return []


def _extract_tables_from_metadata(metadata: Dict[str, Any]) -> List[Dict[str, Any]]:
    for key in (
        "tables",
        "data_tables",
        "source_tables",
        "model_tables",
        "semantic_tables",
        "datasets",
    ):
        tables = _normalise_list(metadata.get(key))
        if tables:
            return tables
    return []


def _extract_worksheets_from_metadata(metadata: Dict[str, Any]) -> List[Dict[str, Any]]:
    for key in (
        "worksheets",
        "answers",
        "sheets",
        "visuals",
        "charts",
        "dashboard_visuals",
    ):
        worksheets = _normalise_list(metadata.get(key))
        if worksheets:
            return worksheets
    return []


def _extract_liveboards_from_metadata(metadata: Dict[str, Any]) -> List[Dict[str, Any]]:
    for key in ("liveboards", "dashboards", "boards"):
        liveboards = _normalise_list(metadata.get(key))
        if liveboards:
            return liveboards
    return []


def _normalise_conversions(conversions: Optional[List[Dict[str, Any]]]) -> List[Dict[str, Any]]:
    result: List[Dict[str, Any]] = []
    seen = set()

    for item in conversions or []:
        if not isinstance(item, dict):
            continue

        name = (
            item.get("name")
            or item.get("measure")
            or item.get("calculated_field")
            or item.get("field_name")
            or item.get("source_calculated_field")
        )
        dax = (
            item.get("dax")
            or item.get("dax_formula")
            or item.get("converted_dax_formula")
            or item.get("expression")
        )

        if not name:
            continue

        fixed = dict(item)
        fixed["name"] = str(name)
        fixed["calculated_field"] = str(name)
        if dax:
            fixed["dax"] = str(dax)

        key = fixed["name"].lower()
        if key in seen:
            continue
        seen.add(key)
        result.append(fixed)

    return result


def _pretty_title(value: Any, fallback: str = "Metric") -> str:
    text = _safe_name(value, fallback)
    text = re.sub(r"^cf_", "", text, flags=re.IGNORECASE)
    text = text.replace("_", " ")
    return " ".join(w.capitalize() for w in text.split())


def _count_from_metadata(metadata: Dict[str, Any], explicit_keys: List[str], fallback_count: int) -> int:
    for key in explicit_keys:
        value = metadata.get(key)
        if isinstance(value, int):
            return value
    return fallback_count


# -----------------------------------------------------------------------------
# Report/package template creator
# -----------------------------------------------------------------------------

class PowerBIReportTemplateCreator:
    """Create lightweight report files for the migration export package."""

    def create_report_folder(
        self,
        report_dir: str | Path,
        metadata: Optional[Dict[str, Any]] = None,
        conversions: Optional[List[Dict[str, Any]]] = None,
        relationships: Optional[List[Dict[str, Any]]] = None,
    ) -> Path:
        """
        Create a .Report folder with report.json and supporting summary files.

        This is intentionally safe:
        - If VisualConverter works, it creates starter visuals.
        - If VisualConverter fails, it creates a blank stable page.
        - Power BI users can still manually create visuals from fields.
        """
        report_dir = Path(report_dir)
        report_dir.mkdir(parents=True, exist_ok=True)

        metadata = metadata or {}
        conversions = _normalise_conversions(conversions or [])
        relationships = relationships or []

        report_name = _safe_name(
            metadata.get("report_name")
            or metadata.get("project_name")
            or metadata.get("name")
            or "ThoughtSpot Migration Report"
        )

        tables = _extract_tables_from_metadata(metadata)
        worksheets = _extract_worksheets_from_metadata(metadata)
        liveboards = _extract_liveboards_from_metadata(metadata)

        visuals: List[Any] = []
        report_json: Dict[str, Any]

        try:
            from .visual_converter import VisualConverter
        except Exception:
            try:
                from visual_converter import VisualConverter  # type: ignore
            except Exception as exc:
                VisualConverter = None  # type: ignore
                logger.warning(f"Could not import VisualConverter: {exc}")

        if VisualConverter:
            try:
                converter = VisualConverter(tables=tables, calculated_fields=conversions)
                if worksheets:
                    visuals = converter.convert_worksheets_to_visuals(worksheets)
                else:
                    visuals = converter.build_default_visuals()

                page = converter.generate_page_json(DEFAULT_PAGE_NAME, visuals)
                report_json = converter.generate_report_json([page])
            except Exception as exc:
                logger.warning(f"VisualConverter failed. Creating safe blank report. Error: {exc}")
                report_json = self._fallback_report_json()
                visuals = []
        else:
            report_json = self._fallback_report_json()

        # Main report files.
        _write_json(report_dir / "report.json", report_json)
        _write_json(report_dir / "Layout", self._report_json_to_legacy_layout(report_json))
        self._write_definition_files(report_dir, report_json)

        # Polished helper artifacts.
        _write_text(
            report_dir / "README.md",
            self._build_readme(
                report_name=report_name,
                metadata=metadata,
                tables=tables,
                worksheets=worksheets,
                liveboards=liveboards,
                conversions=conversions,
                relationships=relationships,
                visual_count=len(visuals),
            ),
        )

        _write_json(
            report_dir / "visual_suggestions.json",
            self.build_visual_suggestions(metadata, conversions, tables),
        )

        _write_json(
            report_dir / "migration_summary.json",
            self._build_migration_summary(
                report_name=report_name,
                metadata=metadata,
                tables=tables,
                worksheets=worksheets,
                liveboards=liveboards,
                conversions=conversions,
                relationships=relationships,
                visual_count=len(visuals),
            ),
        )

        logger.info(f"Created polished report folder: {report_dir} with {len(visuals)} starter visuals")
        return report_dir

    def _fallback_report_json(self) -> Dict[str, Any]:
        return {
            "version": "5.54",
            "themeCollection": {
                "baseTheme": {
                    "name": "CY23SU10",
                    "version": "5.54",
                    "type": 2,
                }
            },
            "activeSectionName": "ReportSection",
            "sections": [
                {
                    "name": "ReportSection",
                    "displayName": DEFAULT_PAGE_NAME,
                    "displayOption": 1,
                    "height": DEFAULT_REPORT_HEIGHT,
                    "width": DEFAULT_REPORT_WIDTH,
                    "ordinal": 0,
                    "filters": "[]",
                    "visualContainers": [],
                    "config": json.dumps(
                        {
                            "objects": {
                                "background": [
                                    {
                                        "properties": {
                                            "color": {
                                                "solid": {
                                                    "color": "#F8FAFC"
                                                }
                                            },
                                            "transparency": 0,
                                        }
                                    }
                                ]
                            }
                        },
                        ensure_ascii=False,
                    ),
                }
            ],
            "config": json.dumps(
                {
                    "version": "5.54",
                    "settings": {
                        "useStylableVisualContainerHeader": True,
                        "exportDataMode": 1,
                    },
                },
                ensure_ascii=False,
            ),
            "layoutOptimization": 0,
        }

    def _report_json_to_legacy_layout(self, report_json: Dict[str, Any]) -> Dict[str, Any]:
        sections = report_json.get("sections") or []
        return {
            "id": 0,
            "sections": sections,
            "config": report_json.get("config", "{}"),
            "layoutOptimization": report_json.get("layoutOptimization", 0),
        }

    def _write_definition_files(self, report_dir: Path, report_json: Dict[str, Any]) -> None:
        """
        Writes PBIP/PBIR-style definition folders for tools that expect them.
        The main report.json remains present for compatibility.
        """
        definition = report_dir / "definition"
        pages_dir = definition / "pages"
        pages_dir.mkdir(parents=True, exist_ok=True)

        _write_json(
            definition / "report.json",
            {
                "version": report_json.get("version", "5.54"),
                "themeCollection": report_json.get("themeCollection", {}),
                "activeSectionName": report_json.get("activeSectionName", "ReportSection"),
            },
        )

        for index, section in enumerate(report_json.get("sections") or []):
            page_name = section.get("name") or f"Page_{index + 1}"
            page_dir = pages_dir / _safe_file_name(page_name, f"Page_{index + 1}")
            page_dir.mkdir(parents=True, exist_ok=True)

            page_json = dict(section)
            visuals = page_json.pop("visualContainers", [])

            _write_json(page_dir / "page.json", page_json)

            visuals_dir = page_dir / "visuals"
            for v_index, visual in enumerate(visuals or []):
                visual_name = visual.get("name") or f"Visual_{v_index + 1}"
                v_dir = visuals_dir / _safe_file_name(visual_name, f"Visual_{v_index + 1}")
                v_dir.mkdir(parents=True, exist_ok=True)
                _write_json(v_dir / "visual.json", visual)

    def _build_readme(
        self,
        report_name: str,
        metadata: Dict[str, Any],
        tables: List[Dict[str, Any]],
        worksheets: List[Dict[str, Any]],
        liveboards: List[Dict[str, Any]],
        conversions: List[Dict[str, Any]],
        relationships: List[Dict[str, Any]],
        visual_count: int,
    ) -> str:
        dashboard_count = _count_from_metadata(metadata, ["total_dashboards", "dashboard_count"], len(liveboards))
        worksheet_count = _count_from_metadata(metadata, ["total_worksheets", "worksheet_count"], len(worksheets))
        table_count = _count_from_metadata(metadata, ["total_tables", "table_count"], len(tables))
        calc_count = _count_from_metadata(metadata, ["total_calculated_fields"], len(conversions))

        top_measures = [_pretty_title(c.get("name") or c.get("calculated_field")) for c in conversions[:8]]
        table_names = [_safe_name(t.get("name") or t.get("table_name") or f"Table {i+1}") for i, t in enumerate(tables[:10])]

        lines = [
            f"# {report_name}",
            "",
            "Generated by **ThoughtSpot → Power BI Migration Tool**.",
            "",
            "## Migration Summary",
            f"- Dashboards / Liveboards: {dashboard_count}",
            f"- Worksheets / Answers: {worksheet_count}",
            f"- Tables: {table_count}",
            f"- Calculated fields: {calc_count}",
            f"- DAX conversions: {len(conversions)}",
            f"- Relationships: {len(relationships)}",
            f"- Starter visuals created: {visual_count}",
            "",
            "## Generated Power BI Assets",
            "- Semantic model files",
            "- DAX measure table",
            "- Table definitions",
            "- Relationship metadata",
            "- Starter report page",
            "- Visual recommendation file",
            "",
            "## Tables",
        ]

        if table_names:
            lines.extend([f"- {name}" for name in table_names])
        else:
            lines.append("- No table metadata detected")

        lines.extend(["", "## Key Measures"])
        if top_measures:
            lines.extend([f"- {name}" for name in top_measures])
        else:
            lines.append("- No calculated fields detected")

        lines.extend(
            [
                "",
                "## Recommended Power BI Validation",
                "1. Open the generated `.pbip` file.",
                "2. Confirm tables are visible in the Data pane.",
                "3. Expand `DAX Measures` and review generated formulas.",
                "4. Create a chart using a dimension and a numeric measure.",
                "5. Avoid clicking `Upgrade report` unless you have saved a backup copy.",
                "",
                "## Demo Note",
                "This package is intended to validate metadata, DAX, model structure, and report scaffolding.",
                "Customer production data should be connected from the approved customer data source.",
            ]
        )

        return "\n".join(lines)

    def _build_migration_summary(
        self,
        report_name: str,
        metadata: Dict[str, Any],
        tables: List[Dict[str, Any]],
        worksheets: List[Dict[str, Any]],
        liveboards: List[Dict[str, Any]],
        conversions: List[Dict[str, Any]],
        relationships: List[Dict[str, Any]],
        visual_count: int,
    ) -> Dict[str, Any]:
        return {
            "report_name": report_name,
            "created_at": datetime.utcnow().isoformat() + "Z",
            "source": metadata.get("source", "ThoughtSpot"),
            "target": "Power BI",
            "summary": {
                "dashboards": _count_from_metadata(metadata, ["total_dashboards", "dashboard_count"], len(liveboards)),
                "worksheets": _count_from_metadata(metadata, ["total_worksheets", "worksheet_count"], len(worksheets)),
                "tables": _count_from_metadata(metadata, ["total_tables", "table_count"], len(tables)),
                "calculated_fields": _count_from_metadata(metadata, ["total_calculated_fields"], len(conversions)),
                "relationships": len(relationships),
                "starter_visuals": visual_count,
            },
            "measure_table": SAFE_MEASURE_TABLE_NAME,
            "status": "generated",
            "notes": [
                "Generated model is suitable for validation and demo.",
                "Production data connection should be configured in customer environment.",
            ],
        }

    def build_visual_suggestions(
        self,
        metadata: Optional[Dict[str, Any]] = None,
        conversions: Optional[List[Dict[str, Any]]] = None,
        tables: Optional[List[Dict[str, Any]]] = None,
    ) -> Dict[str, Any]:
        """
        Rule-based visual suggestions, no Gemini/API required.
        This helps frontend/report package show attractive recommendations.
        """
        metadata = metadata or {}
        conversions = _normalise_conversions(conversions or [])
        tables = tables or _extract_tables_from_metadata(metadata)

        measure_names = [
            c.get("calculated_field") or c.get("name") or c.get("field_name")
            for c in conversions
            if c.get("calculated_field") or c.get("name") or c.get("field_name")
        ]

        all_columns = []
        for table in tables:
            table_name = _safe_name(table.get("name") or table.get("table_name") or "Table")
            for col in table.get("columns") or table.get("fields") or []:
                if isinstance(col, dict):
                    col_name = col.get("name") or col.get("column_name")
                    data_type = str(col.get("data_type") or col.get("type") or "").lower()
                else:
                    col_name = str(col)
                    data_type = ""
                if col_name:
                    all_columns.append(
                        {
                            "table": table_name,
                            "column": _safe_name(col_name),
                            "data_type": data_type,
                        }
                    )

        def find_col(*tokens: str) -> Optional[Dict[str, str]]:
            for col in all_columns:
                low = col["column"].lower()
                if any(t.lower() in low for t in tokens):
                    return col
            return None

        region = find_col("region", "country", "state", "city")
        category = find_col("category", "segment", "sub category")
        date_col = find_col("date", "month", "quarter", "year")
        sales_measure = next((m for m in measure_names if "sales" in m.lower()), None)
        profit_measure = next((m for m in measure_names if "profit" in m.lower()), None)
        margin_measure = next((m for m in measure_names if "margin" in m.lower()), None)
        count_measure = next((m for m in measure_names if "count" in m.lower() or "order" in m.lower()), None)

        suggestions: List[Dict[str, Any]] = []

        for measure in [sales_measure, profit_measure, margin_measure, count_measure]:
            if measure:
                suggestions.append(
                    {
                        "title": _pretty_title(measure),
                        "visual_type": "card",
                        "measure": measure,
                        "table": SAFE_MEASURE_TABLE_NAME,
                        "position": "top_kpi",
                    }
                )

        if region and sales_measure:
            suggestions.append(
                {
                    "title": f"Sales by {region['column']}",
                    "visual_type": "clusteredColumnChart",
                    "axis": region,
                    "measure": sales_measure,
                    "table": SAFE_MEASURE_TABLE_NAME,
                    "position": "main_left",
                }
            )

        if category and profit_measure:
            suggestions.append(
                {
                    "title": f"Profit by {category['column']}",
                    "visual_type": "barChart",
                    "axis": category,
                    "measure": profit_measure,
                    "table": SAFE_MEASURE_TABLE_NAME,
                    "position": "main_right",
                }
            )

        if date_col and sales_measure:
            suggestions.append(
                {
                    "title": f"Sales Trend by {date_col['column']}",
                    "visual_type": "lineChart",
                    "axis": date_col,
                    "measure": sales_measure,
                    "table": SAFE_MEASURE_TABLE_NAME,
                    "position": "bottom_left",
                }
            )

        if category and count_measure:
            suggestions.append(
                {
                    "title": f"Orders by {category['column']}",
                    "visual_type": "donutChart",
                    "category": category,
                    "measure": count_measure,
                    "table": SAFE_MEASURE_TABLE_NAME,
                    "position": "bottom_right",
                }
            )

        if not suggestions:
            for idx, name in enumerate(measure_names[:8]):
                suggestions.append(
                    {
                        "title": _pretty_title(name),
                        "visual_type": "card" if idx < 4 else "clusteredColumnChart",
                        "measure": name,
                        "table": SAFE_MEASURE_TABLE_NAME,
                    }
                )

        return {
            "created_at": datetime.utcnow().isoformat() + "Z",
            "strategy": "rule_based_no_external_api",
            "measure_table": SAFE_MEASURE_TABLE_NAME,
            "suggested_visual_count": len(suggestions),
            "theme": {
                "background": "#F8FAFC",
                "primary": "#2563EB",
                "success": "#16A34A",
                "warning": "#F59E0B",
                "danger": "#DC2626",
            },
            "visuals": suggestions,
            "note": "These suggestions are generated using metadata only. No external AI/API key is required.",
        }


# -----------------------------------------------------------------------------
# Backward compatible starter PBIX creator
# -----------------------------------------------------------------------------

class StarterPBIXCreator:
    """
    Backward-compatible starter template creator.

    Uses "DAX Measures" instead of unsupported "Measures".
    """

    def __init__(self):
        self.template_dir = Path(__file__).parent / "templates"
        self.template_dir.mkdir(exist_ok=True)

    def create_blank_template(
        self,
        output_path: str,
        include_measures_table: bool = True,
        include_date_table: bool = False,
    ) -> Path:
        logger.info(f"Creating blank Power BI template: {output_path}")
        output_path = Path(output_path)
        output_path.parent.mkdir(parents=True, exist_ok=True)

        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            self._create_content_types(temp_path)
            self._create_data_model(temp_path, include_measures_table, include_date_table)
            self._create_data_model_schema(temp_path)
            self._create_version(temp_path)
            self._create_report(temp_path)
            self._create_metadata(temp_path)
            self._create_diagram_state(temp_path)

            with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zf:
                for file_path in temp_path.rglob("*"):
                    if file_path.is_file():
                        zf.write(file_path, file_path.relative_to(temp_path).as_posix())

        logger.info(f"Created template: {output_path}")
        return output_path

    def _create_content_types(self, temp_path: Path) -> None:
        xml = """<?xml version="1.0" encoding="utf-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="json" ContentType="application/json" />
  <Default Extension="xml" ContentType="application/xml" />
  <Override PartName="/DataModel" ContentType="application/x-tmdl-data" />
  <Override PartName="/DataModelSchema" ContentType="application/x-tmdl-metadata" />
  <Override PartName="/DiagramState" ContentType="application/json" />
  <Override PartName="/Report/Layout" ContentType="application/json" />
  <Override PartName="/Metadata" ContentType="application/json" />
  <Override PartName="/Version" ContentType="text/plain" />
</Types>
"""
        _write_text(temp_path / "[Content_Types].xml", xml.strip())

    def _create_data_model(self, temp_path: Path, include_measures_table: bool, include_date_table: bool) -> None:
        tables: List[Dict[str, Any]] = []

        if include_measures_table:
            tables.append(
                {
                    "name": SAFE_MEASURE_TABLE_NAME,
                    "description": "Table for migrated DAX measures",
                    "isHidden": False,
                    "columns": [
                        {
                            "name": "Measure Placeholder",
                            "dataType": "string",
                            "isHidden": True,
                            "sourceColumn": "Measure Placeholder",
                        }
                    ],
                    "partitions": [
                        {
                            "name": f"{SAFE_MEASURE_TABLE_NAME} Partition",
                            "mode": "import",
                            "source": {
                                "type": "m",
                                "expression": "let Source = #table({\"Measure Placeholder\"}, {{\"\"}}) in Source",
                            },
                        }
                    ],
                    "measures": [],
                }
            )

        if include_date_table:
            tables.append(
                {
                    "name": "Calendar",
                    "description": "Optional date table",
                    "columns": [
                        {"name": "Date", "dataType": "dateTime", "sourceColumn": "Date"},
                        {"name": "Year", "dataType": "int64", "sourceColumn": "Year"},
                        {"name": "Month", "dataType": "string", "sourceColumn": "Month"},
                    ],
                    "partitions": [
                        {
                            "name": "Calendar Partition",
                            "mode": "import",
                            "source": {
                                "type": "m",
                                "expression": "let Source = #table({\"Date\",\"Year\",\"Month\"}, {{#date(2026,1,1),2026,\"January\"}}) in Source",
                            },
                        }
                    ],
                }
            )

        model = {
            "name": "SemanticModel",
            "compatibilityLevel": 1567,
            "model": {
                "culture": "en-US",
                "defaultPowerBIDataSourceVersion": "powerBI_V3",
                "sourceQueryCulture": "en-US",
                "dataSources": [],
                "tables": tables,
                "relationships": [],
                "expressions": [],
                "annotations": [
                    {"name": "PBI_ProTooling", "value": "[\"DevMode\"]"},
                    {"name": "GeneratedBy", "value": "ThoughtSpotPowerBIMigration"},
                ],
            },
        }
        _write_json(temp_path / "DataModel", model)

    def _create_report(self, temp_path: Path) -> None:
        report_dir = temp_path / "Report"
        report_dir.mkdir(exist_ok=True)
        layout = {
            "id": 0,
            "sections": [
                {
                    "name": "ReportSection",
                    "displayName": DEFAULT_PAGE_NAME,
                    "filters": "[]",
                    "ordinal": 0,
                    "visualContainers": [],
                    "config": "{}",
                    "displayOption": 0,
                    "width": DEFAULT_REPORT_WIDTH,
                    "height": DEFAULT_REPORT_HEIGHT,
                }
            ],
            "config": "{}",
            "layoutOptimization": 0,
        }
        _write_json(report_dir / "Layout", layout)

    def _create_metadata(self, temp_path: Path) -> None:
        _write_json(
            temp_path / "Metadata",
            {
                "version": "3.0",
                "createdBy": "ThoughtSpotPowerBIMigration",
                "createdAt": datetime.utcnow().isoformat() + "Z",
            },
        )

    def _create_diagram_state(self, temp_path: Path) -> None:
        _write_json(temp_path / "DiagramState", {"version": "1.0", "diagramViewState": {}})

    def _create_data_model_schema(self, temp_path: Path) -> None:
        _write_json(
            temp_path / "DataModelSchema",
            {
                "name": "SemanticModel",
                "compatibilityLevel": 1567,
                "model": {"defaultPowerBIDataSourceVersion": "powerBI_V3"},
            },
        )

    def _create_version(self, temp_path: Path) -> None:
        _write_text(temp_path / "Version", "2.0")


# -----------------------------------------------------------------------------
# Public functions used by other backend files
# -----------------------------------------------------------------------------

def create_report_template(
    report_dir: str | Path,
    metadata: Optional[Dict[str, Any]] = None,
    conversions: Optional[List[Dict[str, Any]]] = None,
    relationships: Optional[List[Dict[str, Any]]] = None,
) -> Path:
    creator = PowerBIReportTemplateCreator()
    return creator.create_report_folder(report_dir, metadata, conversions, relationships)


def create_default_templates() -> None:
    creator = StarterPBIXCreator()
    templates_dir = Path("./templates")
    templates_dir.mkdir(parents=True, exist_ok=True)

    creator.create_blank_template(
        output_path=str(templates_dir / "blank_template.pbix"),
        include_measures_table=False,
        include_date_table=False,
    )
    creator.create_blank_template(
        output_path=str(templates_dir / "standard_template.pbix"),
        include_measures_table=True,
        include_date_table=False,
    )
    logger.info("Default templates created in ./templates/")


if __name__ == "__main__":
    create_default_templates()
