"""
visual_converter.py
Power BI visual/report helper for ThoughtSpot -> Power BI migration.

Use this file as: backend/src/powerbi/visual_converter.py

Polished version fixes/improves:
- Keeps table/column names exactly same as semantic model/TMDL.
- Uses measure table name: "DAX Measures".
- Creates better default visuals when worksheet metadata is weak/missing.
- Prefer SUM for numeric business columns instead of Count.
- Builds cleaner chart titles.
- Adds safer visual styling: title, data labels, legend, category labels, and default colors.
- Avoids broken/unsupported custom visual JSON.
- Keeps backward-compatible helper functions.
"""

from __future__ import annotations

import json
import re
import uuid
from dataclasses import dataclass, field
from enum import Enum
from typing import Any, Dict, Iterable, List, Optional, Tuple

try:
    from loguru import logger
except Exception:  # pragma: no cover
    import logging

    logging.basicConfig(level=logging.INFO)
    logger = logging.getLogger(__name__)


MEASURE_TABLE_NAME = "DAX Measures"


class PowerBIVisualType(str, Enum):
    CARD = "card"
    TABLE = "tableEx"
    MATRIX = "pivotTable"
    CLUSTERED_BAR_CHART = "clusteredBarChart"
    CLUSTERED_COLUMN_CHART = "clusteredColumnChart"
    STACKED_BAR_CHART = "barChart"
    STACKED_COLUMN_CHART = "columnChart"
    LINE_CHART = "lineChart"
    AREA_CHART = "areaChart"
    PIE_CHART = "pieChart"
    DONUT_CHART = "donutChart"
    SCATTER_CHART = "scatterChart"
    MAP = "map"
    SLICER = "slicer"


@dataclass
class VisualLayout:
    x: float
    y: float
    width: float
    height: float
    z_index: int = 0


@dataclass
class PowerBIFieldRef:
    table: str
    column: str
    aggregation: Optional[str] = None
    is_measure: bool = False


@dataclass
class PowerBIVisual:
    visual_type: PowerBIVisualType
    name: str
    title: str
    layout: VisualLayout
    data_roles: Dict[str, List[PowerBIFieldRef]]
    filters: List[Dict[str, Any]] = field(default_factory=list)


class VisualConverter:
    CANVAS_WIDTH = 1280
    CANVAS_HEIGHT = 720
    DEFAULT_VISUAL_WIDTH = 390
    DEFAULT_VISUAL_HEIGHT = 245

    # Simple, professional palette. These are embedded in report visual objects only.
    THEME_COLORS = [
        "#118DFF",  # blue
        "#12239E",  # navy
        "#E66C37",  # orange
        "#6B007B",  # purple
        "#E044A7",  # pink
        "#744EC2",  # violet
        "#D9B300",  # gold
        "#D64550",  # red
        "#197278",  # teal
        "#1AAB40",  # green
    ]

    def __init__(
        self,
        tables: Optional[List[Dict[str, Any]]] = None,
        calculated_fields: Optional[List[Dict[str, Any]]] = None,
    ):
        self.tables = tables or []
        self.calculated_fields = calculated_fields or []
        self._field_to_table = self._build_field_to_table_index(self.tables)
        self._measure_table_name = MEASURE_TABLE_NAME

    # ----------------------------- public API -----------------------------

    def convert_worksheets_to_visuals(
        self,
        worksheets: List[Dict[str, Any]],
        auto_layout: bool = True,
    ) -> List[PowerBIVisual]:
        logger.info(f"Converting {len(worksheets or [])} worksheets to Power BI visuals")
        visuals: List[PowerBIVisual] = []
        seen_names: set[str] = set()

        for worksheet in worksheets or []:
            try:
                visual = self._convert_single_worksheet(worksheet or {})
                if not visual:
                    continue
                visual.name = self._unique_name(self._safe_visual_name(visual.name), seen_names)
                visual.title = self._clean_title(visual.title)
                if auto_layout:
                    visual.layout = self._calculate_auto_layout(len(visuals), max(len(worksheets), 1))
                visuals.append(visual)
            except Exception as exc:
                logger.warning(f"Failed to convert worksheet {worksheet.get('name', 'Unknown')}: {exc}")

        # If source worksheets are table-only or weak, add useful default visuals also.
        if not visuals:
            visuals = self.build_default_visuals()
        elif self._looks_like_weak_visual_set(visuals):
            defaults = self.build_default_visuals()
            # Keep maximum 6 visuals so the report is not crowded.
            visuals = (defaults + visuals)[:6]

        if auto_layout:
            for i, visual in enumerate(visuals):
                visual.layout = self._calculate_auto_layout(i, len(visuals))

        logger.info(f"Converted {len(visuals)} visuals")
        return visuals

    def build_default_visuals(self) -> List[PowerBIVisual]:
        """
        Creates useful default visuals from the available semantic model.
        This prevents a blank/boring report when source worksheet metadata is missing.
        """
        visuals: List[PowerBIVisual] = []

        measure_refs = self._measure_refs_from_calculated_fields()
        dimension_refs = self._dimension_refs_from_tables()
        numeric_refs = self._numeric_refs_from_tables()

        # Prefer business columns if available.
        sales = self._find_ref(["sales", "revenue", "amount"], numeric_refs, default_agg="sum")
        profit = self._find_ref(["profit", "margin"], numeric_refs, default_agg="sum")
        quantity = self._find_ref(["quantity", "qty", "units"], numeric_refs, default_agg="sum")
        discount = self._find_ref(["discount"], numeric_refs, default_agg="avg")

        region = self._find_ref(["region", "state", "country", "city"], dimension_refs)
        category = self._find_ref(["category", "segment", "sub category", "sub_category"], dimension_refs)
        date = self._find_ref(["date", "order date", "created"], dimension_refs)
        customer = self._find_ref(["customer", "customer name", "customer id"], dimension_refs)
        product = self._find_ref(["product", "product name", "product id"], dimension_refs)

        # KPI cards: use existing DAX measures first, then important numeric columns.
        kpi_candidates = measure_refs[:4]
        for fallback in [sales, profit, quantity, discount]:
            if fallback and len(kpi_candidates) < 4:
                kpi_candidates.append(fallback)
        kpi_candidates = self._dedupe_refs(kpi_candidates)[:4]

        for idx, ref in enumerate(kpi_candidates):
            visuals.append(
                PowerBIVisual(
                    visual_type=PowerBIVisualType.CARD,
                    name=f"kpi_{idx + 1}_{ref.column}",
                    title=self._friendly_metric_title(ref),
                    layout=VisualLayout(0, 0, 280, 130),
                    data_roles={"Values": [ref]},
                )
            )

        # Sales by Region
        if region and sales:
            visuals.append(
                PowerBIVisual(
                    visual_type=PowerBIVisualType.CLUSTERED_COLUMN_CHART,
                    name="sales_by_region",
                    title="Sales by Region",
                    layout=VisualLayout(0, 0, 600, 260),
                    data_roles={"Category": [region], "Y": [sales]},
                )
            )

        # Profit by Category
        if category and profit:
            visuals.append(
                PowerBIVisual(
                    visual_type=PowerBIVisualType.CLUSTERED_BAR_CHART,
                    name="profit_by_category",
                    title="Profit by Category",
                    layout=VisualLayout(0, 0, 600, 260),
                    data_roles={"Category": [category], "Y": [profit]},
                )
            )

        # Sales trend by date
        if date and sales:
            visuals.append(
                PowerBIVisual(
                    visual_type=PowerBIVisualType.LINE_CHART,
                    name="sales_trend",
                    title="Sales Trend",
                    layout=VisualLayout(0, 0, 600, 260),
                    data_roles={"Category": [date], "Y": [sales]},
                )
            )

        # Product/customer distribution if no enough numeric visuals.
        if category and product:
            visuals.append(
                PowerBIVisual(
                    visual_type=PowerBIVisualType.DONUT_CHART,
                    name="product_distribution_by_category",
                    title="Product Distribution by Category",
                    layout=VisualLayout(0, 0, 390, 245),
                    data_roles={"Category": [category], "Values": [PowerBIFieldRef(product.table, product.column, aggregation="count")]},
                )
            )
        elif region and customer:
            visuals.append(
                PowerBIVisual(
                    visual_type=PowerBIVisualType.DONUT_CHART,
                    name="customer_distribution_by_region",
                    title="Customer Distribution by Region",
                    layout=VisualLayout(0, 0, 390, 245),
                    data_roles={"Category": [region], "Values": [PowerBIFieldRef(customer.table, customer.column, aggregation="count")]},
                )
            )

        # Data preview table.
        table_values = self._dedupe_refs(
            [r for r in [date, region, category, customer, product, sales, profit, quantity, discount] if r]
        )[:8]
        if table_values:
            visuals.append(
                PowerBIVisual(
                    visual_type=PowerBIVisualType.TABLE,
                    name="business_data_preview",
                    title="Business Data Preview",
                    layout=VisualLayout(0, 0, 820, 250),
                    data_roles={"Values": table_values},
                )
            )

        # Last fallback if metadata is too limited.
        if not visuals:
            fallback_fields = (dimension_refs[:4] + numeric_refs[:4] + measure_refs[:4])[:8]
            if fallback_fields:
                visuals.append(
                    PowerBIVisual(
                        visual_type=PowerBIVisualType.TABLE,
                        name="data_preview",
                        title="Data Preview",
                        layout=VisualLayout(0, 0, 820, 250),
                        data_roles={"Values": fallback_fields},
                    )
                )

        return visuals[:8]

    def generate_visual_json(self, visual: PowerBIVisual) -> Dict[str, Any]:
        visual_name = self._safe_visual_name(visual.name)[:60] or f"visual_{uuid.uuid4().hex[:8]}"
        prototype_query, projections = self._build_prototype_query(visual.data_roles)

        position = {
            "x": visual.layout.x,
            "y": visual.layout.y,
            "z": visual.layout.z_index,
            "width": visual.layout.width,
            "height": visual.layout.height,
            "tabOrder": visual.layout.z_index,
        }

        single_visual = {
            "visualType": visual.visual_type.value,
            "projections": projections,
            "prototypeQuery": prototype_query,
            "drillFilterOtherVisuals": True,
            "objects": self._visual_objects(visual),
        }

        config = {
            "name": visual_name,
            "layouts": [{"id": 0, "position": position}],
            "singleVisual": single_visual,
        }

        return {
            "name": visual_name,
            "layouts": [{"id": 0, "position": position}],
            "singleVisual": single_visual,
            "filters": "[]",
            "config": json.dumps(config, ensure_ascii=False),
        }

    def generate_page_json(self, page_name: str, visuals: List[PowerBIVisual]) -> Dict[str, Any]:
        page_id = self._safe_visual_name(page_name or "ReportSection")
        containers = [self.generate_visual_json(v) for v in visuals]
        return {
            "name": page_id,
            "displayName": page_name or "Executive Dashboard",
            "displayOption": 1,
            "height": self.CANVAS_HEIGHT,
            "width": self.CANVAS_WIDTH,
            "ordinal": 0,
            "filters": "[]",
            "visualContainers": containers,
            "config": json.dumps({"objects": {}}, ensure_ascii=False),
        }

    def generate_report_json(self, pages: List[Dict[str, Any]] | List[PowerBIVisual]) -> Dict[str, Any]:
        if pages and isinstance(pages[0], PowerBIVisual):
            pages = [self.generate_page_json("Executive Dashboard", pages)]  # type: ignore[list-item]

        sections = pages or [self.generate_page_json("Executive Dashboard", self.build_default_visuals())]
        active = sections[0].get("name", "Executive_Dashboard") if sections else "Executive_Dashboard"

        return {
            "version": "5.54",
            "themeCollection": {
                "baseTheme": {
                    "name": "CY23SU10",
                    "version": "5.54",
                    "type": 2,
                }
            },
            "activeSectionName": active,
            "sections": sections,
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

    def generate_visual_conversion_report(
        self,
        worksheets: List[Dict[str, Any]],
        visuals: List[PowerBIVisual],
    ) -> str:
        lines = [
            "# Visual Conversion Report",
            "",
            f"**Source Worksheets:** {len(worksheets or [])}",
            f"**Power BI Visuals:** {len(visuals or [])}",
            "",
            "| Worksheet/Title | Power BI Visual | Fields |",
            "|---|---|---|",
        ]
        for v in visuals or []:
            fields = []
            for role, refs in v.data_roles.items():
                fields.append(f"{role}: {', '.join([r.column for r in refs])}")
            lines.append(f"| {v.title} | {v.visual_type.value} | {'; '.join(fields)} |")
        return "\n".join(lines)

    # -------------------------- conversion internals ------------------------

    def _convert_single_worksheet(self, worksheet: Dict[str, Any]) -> Optional[PowerBIVisual]:
        title = str(worksheet.get("name") or worksheet.get("title") or "Sheet")
        visual_type = self._map_visual_type(worksheet)

        rows = self._extract_field_names(
            worksheet.get("rows")
            or worksheet.get("row_fields")
            or worksheet.get("dimensions")
            or worksheet.get("group_by")
            or []
        )
        cols = self._extract_field_names(
            worksheet.get("cols")
            or worksheet.get("columns")
            or worksheet.get("column_fields")
            or []
        )
        marks = self._extract_marks(worksheet)

        roles = self._map_fields_to_data_roles(visual_type, rows, cols, marks)
        if not any(roles.values()):
            fallback = self._fallback_fields()
            roles = {"Values": fallback[:6]} if fallback else {}
            visual_type = PowerBIVisualType.TABLE

        if self._is_generic_title(title):
            title = self._title_from_roles(visual_type, roles)

        return PowerBIVisual(
            visual_type=visual_type,
            name=title,
            title=title,
            layout=VisualLayout(0, 0, self.DEFAULT_VISUAL_WIDTH, self.DEFAULT_VISUAL_HEIGHT),
            data_roles=roles,
            filters=[],
        )

    def _map_visual_type(self, worksheet: Dict[str, Any]) -> PowerBIVisualType:
        raw = str(
            worksheet.get("visual_type")
            or worksheet.get("chart_type")
            or worksheet.get("type")
            or worksheet.get("mark_type")
            or ""
        ).lower()

        if "slicer" in raw or "filter" in raw:
            return PowerBIVisualType.SLICER
        if "card" in raw or "kpi" in raw:
            return PowerBIVisualType.CARD
        if "matrix" in raw or "pivot" in raw:
            return PowerBIVisualType.MATRIX
        if "table" in raw or "text" in raw:
            return PowerBIVisualType.TABLE
        if "line" in raw or "trend" in raw:
            return PowerBIVisualType.LINE_CHART
        if "area" in raw:
            return PowerBIVisualType.AREA_CHART
        if "pie" in raw:
            return PowerBIVisualType.PIE_CHART
        if "donut" in raw or "doughnut" in raw:
            return PowerBIVisualType.DONUT_CHART
        if "scatter" in raw or "circle" in raw:
            return PowerBIVisualType.SCATTER_CHART
        if "map" in raw:
            return PowerBIVisualType.MAP
        if "bar" in raw:
            return PowerBIVisualType.CLUSTERED_BAR_CHART
        if "column" in raw:
            return PowerBIVisualType.CLUSTERED_COLUMN_CHART
        return PowerBIVisualType.CLUSTERED_COLUMN_CHART

    def _map_fields_to_data_roles(
        self,
        visual_type: PowerBIVisualType,
        rows: List[str],
        columns: List[str],
        marks: List[str],
    ) -> Dict[str, List[PowerBIFieldRef]]:
        """
        Map ThoughtSpot fields to Power BI visual roles in a safe, clean way.

        Important fixes:
        - Numeric fields are always placed as aggregated Values/Y fields.
        - Only real dimension/text/date fields are used for Category/Legend/Series.
        - Avoid putting several numeric columns into one chart, because that creates
          crowded 100% stacked visuals and unreadable legends.
        - Prefer SUM for Sales/Profit/Quantity/Cost and AVG for Discount/Margin/Rate.
        """
        row_refs = self._dedupe_refs([r for r in [self._field_ref(f) for f in rows] if r])
        col_refs = self._dedupe_refs([r for r in [self._field_ref(f) for f in columns] if r])
        mark_refs = self._dedupe_refs([r for r in [self._field_ref(f, prefer_measure=True) for f in marks] if r])

        all_refs = self._dedupe_refs(row_refs + col_refs + mark_refs)

        def is_measure_ref(ref: PowerBIFieldRef) -> bool:
            return bool(ref.aggregation or ref.is_measure)

        def is_dim_ref(ref: PowerBIFieldRef) -> bool:
            return not is_measure_ref(ref)

        dims = self._dedupe_refs([r for r in all_refs if is_dim_ref(r)])
        measures = self._dedupe_refs([r for r in all_refs if is_measure_ref(r)])

        # Prefer business-friendly numeric fields when multiple numeric columns are present.
        def measure_score(ref: PowerBIFieldRef) -> int:
            name = f"{ref.table} {ref.column}".lower()
            if "sales" in name or "revenue" in name or "amount" in name:
                return 100
            if "profit" in name:
                return 90
            if "quantity" in name or "qty" in name:
                return 80
            if "cost" in name:
                return 70
            if "discount" in name or "margin" in name or "rate" in name:
                return 60
            if ref.is_measure:
                return 50
            return 10

        measures = sorted(measures, key=measure_score, reverse=True)

        # If no explicit measure exists, use the best numeric column from the model.
        if not measures:
            numeric = [n for n in self._numeric_refs_from_tables() if n not in dims]
            measures = sorted(numeric, key=measure_score, reverse=True)

        # Select a secondary dimension only when it is not the main category.
        category = dims[:1]
        series = [d for d in dims[1:] if not category or not (d.table == category[0].table and d.column == category[0].column)][:1]

        if visual_type == PowerBIVisualType.CARD:
            return {"Values": (measures or mark_refs or row_refs or col_refs)[:1]}

        if visual_type == PowerBIVisualType.TABLE:
            # Tables may show multiple fields, but keep them clean and non-duplicated.
            return {"Values": self._dedupe_refs(row_refs + col_refs + mark_refs)[:10]}

        if visual_type == PowerBIVisualType.MATRIX:
            return {"Rows": dims[:2], "Columns": dims[2:3], "Values": measures[:3]}

        if visual_type in {PowerBIVisualType.PIE_CHART, PowerBIVisualType.DONUT_CHART}:
            return {"Category": category, "Values": measures[:1]}

        if visual_type == PowerBIVisualType.SCATTER_CHART:
            return {
                "X": measures[:1],
                "Y": measures[1:2] or measures[:1],
                "Details": category,
            }

        if visual_type == PowerBIVisualType.SLICER:
            return {"Values": category or row_refs[:1] or col_refs[:1]}

        if visual_type == PowerBIVisualType.LINE_CHART:
            # One clean metric by one date/category; optional series only if dimension exists.
            roles = {"Category": category, "Y": measures[:1]}
            if series:
                roles["Series"] = series
            return roles

        # Bar/column charts: one category, one numeric value.
        # This prevents Power BI from creating Count of Sales / 100% stacked messy visuals.
        roles = {"Category": category, "Y": measures[:1]}
        if series:
            roles["Series"] = series
        return roles

    # -------------------------- Power BI query JSON -------------------------

    def _build_prototype_query(
        self,
        roles: Dict[str, List[PowerBIFieldRef]],
    ) -> Tuple[Dict[str, Any], Dict[str, List[Dict[str, Any]]]]:
        selects = []
        projections: Dict[str, List[Dict[str, Any]]] = {}
        table_aliases: Dict[str, str] = {}

        def alias_for(table: str) -> str:
            if table not in table_aliases:
                table_aliases[table] = f"t{len(table_aliases)}"
            return table_aliases[table]

        for role, refs in roles.items():
            projections[role] = []
            for ref in refs or []:
                # IMPORTANT FIX:
                # Power BI defaults plain numeric columns to Count in many imported PBIP reports.
                # For chart/KPI value buckets we must write an explicit Aggregation expression
                # (SUM/AVG) so the visual shows Total Profit/Sales instead of Count of Profit/Sales.
                ref = self._normalize_ref_for_role(role, ref)

                table = self._clean_table_name(ref.table)
                col = self._clean_column_name(ref.column)
                alias = alias_for(table)
                query_ref = f"{table}.{col}"

                if ref.is_measure:
                    expr = {
                        "Measure": {
                            "Expression": {"SourceRef": {"Source": alias}},
                            "Property": col,
                        }
                    }
                elif ref.aggregation:
                    expr = {
                        "Aggregation": {
                            "Expression": {
                                "Column": {
                                    "Expression": {"SourceRef": {"Source": alias}},
                                    "Property": col,
                                }
                            },
                            "Function": self._agg_code(ref.aggregation),
                        }
                    }
                    # QueryRef for aggregated fields should remain stable and readable.
                    query_ref = f"{self._agg_query_name(ref.aggregation)}({table}.{col})"
                else:
                    expr = {
                        "Column": {
                            "Expression": {"SourceRef": {"Source": alias}},
                            "Property": col,
                        }
                    }

                selects.append(
                    {
                        "Name": query_ref,
                        "NativeReferenceName": self._native_reference_name(ref),
                        "Expression": expr,
                    }
                )
                projections[role].append({"queryRef": query_ref, "active": True})

        if not selects:
            fallback_table = self._clean_table_name(self.tables[0].get("name", "orders_fact")) if self.tables else "orders_fact"
            alias = alias_for(fallback_table)
            selects.append(
                {
                    "Name": f"{fallback_table}.Dummy",
                    "NativeReferenceName": "Dummy",
                    "Expression": {
                        "Column": {
                            "Expression": {"SourceRef": {"Source": alias}},
                            "Property": "Dummy",
                        }
                    },
                }
            )
            projections = {"Values": [{"queryRef": f"{fallback_table}.Dummy", "active": True}]}

        from_list = [{"Name": alias, "Entity": table, "Type": 0} for table, alias in table_aliases.items()]
        return {"Version": 2, "From": from_list, "Select": selects}, projections

    # ------------------------------- styling --------------------------------

    def _visual_objects(self, visual: PowerBIVisual) -> Dict[str, Any]:
        # Keep this conservative; invalid Power BI object JSON can break rendering.
        title = self._clean_title(visual.title)
        objects: Dict[str, Any] = {
            "title": [
                {
                    "properties": {
                        "show": {"expr": {"Literal": {"Value": "true"}}},
                        "text": {"expr": {"Literal": {"Value": json.dumps(title)}}},
                        "fontSize": {"expr": {"Literal": {"Value": "11D"}}},
                        "fontColor": {"solid": {"color": "#252423"}},
                        "alignment": {"expr": {"Literal": {"Value": "'left'"}}},
                    },
                    "selector": None,
                }
            ],
            "background": [
                {
                    "properties": {
                        "show": {"expr": {"Literal": {"Value": "true"}}},
                        "color": {"solid": {"color": "#FFFFFF"}},
                        "transparency": {"expr": {"Literal": {"Value": "0D"}}},
                    },
                    "selector": None,
                }
            ],
            "border": [
                {
                    "properties": {
                        "show": {"expr": {"Literal": {"Value": "true"}}},
                        "color": {"solid": {"color": "#E1E5EA"}},
                    },
                    "selector": None,
                }
            ],
        }

        if visual.visual_type in {
            PowerBIVisualType.CLUSTERED_BAR_CHART,
            PowerBIVisualType.CLUSTERED_COLUMN_CHART,
            PowerBIVisualType.STACKED_BAR_CHART,
            PowerBIVisualType.STACKED_COLUMN_CHART,
            PowerBIVisualType.LINE_CHART,
            PowerBIVisualType.AREA_CHART,
            PowerBIVisualType.PIE_CHART,
            PowerBIVisualType.DONUT_CHART,
        }:
            objects.update(
                {
                    "dataPoint": [
                        {
                            "properties": {
                                "defaultColor": {"solid": {"color": self.THEME_COLORS[0]}},
                            },
                            "selector": None,
                        }
                    ],
                    "labels": [
                        {
                            "properties": {
                                "show": {"expr": {"Literal": {"Value": "true"}}},
                                "fontSize": {"expr": {"Literal": {"Value": "9D"}}},
                            },
                            "selector": None,
                        }
                    ],
                    "legend": [
                        {
                            "properties": {
                                "show": {"expr": {"Literal": {"Value": "true"}}},
                                "position": {"expr": {"Literal": {"Value": "'Top'"}}},
                            },
                            "selector": None,
                        }
                    ],
                    "categoryAxis": [
                        {
                            "properties": {
                                "show": {"expr": {"Literal": {"Value": "true"}}},
                                "fontSize": {"expr": {"Literal": {"Value": "9D"}}},
                            },
                            "selector": None,
                        }
                    ],
                    "valueAxis": [
                        {
                            "properties": {
                                "show": {"expr": {"Literal": {"Value": "true"}}},
                                "fontSize": {"expr": {"Literal": {"Value": "9D"}}},
                            },
                            "selector": None,
                        }
                    ],
                }
            )

        if visual.visual_type == PowerBIVisualType.CARD:
            objects.update(
                {
                    "labels": [
                        {
                            "properties": {
                                "fontSize": {"expr": {"Literal": {"Value": "22D"}}},
                                "color": {"solid": {"color": self.THEME_COLORS[0]}},
                            },
                            "selector": None,
                        }
                    ]
                }
            )

        return objects

    # ------------------------------- helpers --------------------------------

    def _measure_refs_from_calculated_fields(self) -> List[PowerBIFieldRef]:
        refs: List[PowerBIFieldRef] = []

        # 1) Measures converted from ThoughtSpot calculated fields.
        for item in self.calculated_fields or []:
            name = (
                item.get("name")
                or item.get("measure")
                or item.get("calculated_field")
                or item.get("field_name")
            )
            if name:
                refs.append(
                    PowerBIFieldRef(
                        table=self._measure_table_name,
                        column=self._clean_column_name(str(name)),
                        is_measure=True,
                    )
                )

        # 2) Auto measures created by pbip_tmdl_injector.py.
        # This is the hard fix for Power BI showing only Count/Count Distinct.
        # Visuals should use measures like [Total Sales], not raw numeric columns.
        refs.extend(self._auto_measure_refs_from_numeric_tables())

        return self._dedupe_refs(refs)

    def _auto_measure_refs_from_numeric_tables(self) -> List[PowerBIFieldRef]:
        refs: List[PowerBIFieldRef] = []
        for table in self.tables or []:
            for col in table.get("columns") or table.get("fields") or []:
                cname = str(col.get("name") if isinstance(col, dict) else col)
                dtype = str(col.get("dataType") or col.get("data_type") or col.get("type") or "").lower() if isinstance(col, dict) else ""
                if dtype in {"int64", "integer", "int", "double", "decimal", "number", "float", "currency"} or self._looks_numeric_field_name(cname):
                    measure_name = self._auto_measure_name_for_column(cname)
                    refs.append(PowerBIFieldRef(self._measure_table_name, measure_name, is_measure=True))
        return self._dedupe_refs(refs)

    def _auto_measure_name_for_column(self, column: str) -> str:
        base = self._pretty_name(column)
        agg = self._default_aggregation_for_column(column)
        if agg in {"avg", "average"}:
            return f"Average {base}"
        if agg == "count_distinct":
            return f"Distinct Count {base}"
        if agg == "count":
            return f"Count of {base}"
        return f"Total {base}"

    def _dimension_refs_from_tables(self) -> List[PowerBIFieldRef]:
        refs: List[PowerBIFieldRef] = []
        for table in self.tables or []:
            tname = self._clean_table_name(str(table.get("name") or table.get("table_name") or "Table"))
            for col in table.get("columns") or table.get("fields") or []:
                cname = str(col.get("name") if isinstance(col, dict) else col)
                dtype = str(col.get("dataType") or col.get("data_type") or col.get("type") or "").lower() if isinstance(col, dict) else ""

                # When upstream metadata is weak, Sales/Profit/Quantity may arrive without
                # a dataType. Do NOT treat those fields as dimensions, otherwise Power BI
                # creates Count of Sales instead of Sum of Sales.
                if self._looks_numeric_field_name(cname):
                    continue

                if dtype in {"string", "text", "date", "datetime", "datetimezone", "boolean"} or not dtype:
                    refs.append(PowerBIFieldRef(tname, self._clean_column_name(cname)))
        return self._dedupe_refs(refs)

    def _numeric_refs_from_tables(self) -> List[PowerBIFieldRef]:
        # Return the auto DAX measures for numeric fields. This prevents Power BI Desktop
        # from treating Sales/Profit as text-like fields and showing only Count.
        return self._auto_measure_refs_from_numeric_tables()

    def _extract_marks(self, worksheet: Dict[str, Any]) -> List[str]:
        values: List[str] = []
        for key in ("marks", "measures", "values", "metrics", "y", "y_axis", "size", "color"):
            values.extend(self._extract_field_names(worksheet.get(key) or []))
        for pane in worksheet.get("pane_encodings") or []:
            enc = pane.get("encodings") or {}
            if isinstance(enc, dict):
                values.extend(self._extract_field_names(list(enc.values())))
        return self._dedupe_strings(values)

    def _extract_field_names(self, obj: Any) -> List[str]:
        out: List[str] = []
        if obj is None:
            return out
        if isinstance(obj, str):
            return [self._clean_field_token(obj)] if obj.strip() else []
        if isinstance(obj, dict):
            for key in ("name", "field", "field_name", "column", "column_name", "caption"):
                if obj.get(key):
                    return [self._clean_field_token(str(obj[key]))]
            for value in obj.values():
                out.extend(self._extract_field_names(value))
            return out
        if isinstance(obj, Iterable):
            for item in obj:
                out.extend(self._extract_field_names(item))
        return self._dedupe_strings([x for x in out if x])

    def _field_ref(self, field_name: str, prefer_measure: bool = False) -> Optional[PowerBIFieldRef]:
        if not field_name:
            return None

        clean = self._clean_field_token(field_name)
        agg = None

        m = re.match(r"(?i)^\s*(sum|avg|average|count|countd|count_distinct|min|max)\s*\((.*?)\)\s*$", clean)
        if m:
            agg = m.group(1).lower().replace("average", "avg").replace("countd", "count_distinct")
            clean = self._clean_field_token(m.group(2))

        table = None
        mt = re.search(r"'([^']+)'\s*\[([^\]]+)\]", clean)
        if mt:
            table, clean = mt.group(1), mt.group(2)
        else:
            mt = re.search(r"\[([^\]]+)\]", clean)
            if mt:
                clean = mt.group(1)

        table = table or self._field_to_table.get(clean.lower()) or self._guess_default_table(clean, prefer_measure)
        is_measure = prefer_measure and not agg and table == self._measure_table_name

        # HARD FIX: metric buckets must use DAX measures. A raw column can still appear
        # as Count in Desktop even when its datatype is numeric in the semantic model.
        if prefer_measure and not agg and (self._is_known_numeric_column(table, clean) or self._looks_numeric_field_name(clean)):
            return PowerBIFieldRef(
                table=self._measure_table_name,
                column=self._auto_measure_name_for_column(clean),
                aggregation=None,
                is_measure=True,
            )

        # If not a measure and the column is known numeric, keep an explicit aggregation.
        if not is_measure and not agg and self._is_known_numeric_column(table, clean):
            agg = self._default_aggregation_for_column(clean)

        return PowerBIFieldRef(
            table=self._clean_table_name(table),
            column=self._clean_column_name(clean),
            aggregation=agg,
            is_measure=is_measure,
        )

    def _guess_default_table(self, column: str, prefer_measure: bool = False) -> str:
        if prefer_measure and self.calculated_fields:
            return self._measure_table_name

        if self.tables:
            facts = [
                t
                for t in self.tables
                if "fact" in str(t.get("name") or t.get("table_name") or "").lower()
                or "sales" in str(t.get("name") or t.get("table_name") or "").lower()
                or "orders" in str(t.get("name") or t.get("table_name") or "").lower()
            ]
            chosen = facts[0] if facts else self.tables[0]
            return str(chosen.get("name") or chosen.get("table_name") or "orders_fact")

        return "orders_fact"

    def _build_field_to_table_index(self, tables: List[Dict[str, Any]]) -> Dict[str, str]:
        index: Dict[str, str] = {}
        for table in tables or []:
            tname = self._clean_table_name(str(table.get("name") or table.get("table_name") or "orders_fact"))
            cols = table.get("columns") or table.get("fields") or []
            for col in cols:
                cname = str(col.get("name") or col.get("column") or col.get("column_name") or "") if isinstance(col, dict) else str(col)
                if cname:
                    index[self._clean_column_name(cname).lower()] = tname
        return index

    def _fallback_fields(self) -> List[PowerBIFieldRef]:
        refs = []
        measure_refs = self._measure_refs_from_calculated_fields()
        if measure_refs:
            return measure_refs

        for table in self.tables[:1]:
            tname = self._clean_table_name(str(table.get("name") or table.get("table_name") or "orders_fact"))
            for col in (table.get("columns") or table.get("fields") or [])[:6]:
                cname = col.get("name") if isinstance(col, dict) else str(col)
                refs.append(PowerBIFieldRef(tname, self._clean_column_name(cname)))
        return refs

    def _calculate_auto_layout(self, index: int, total: int) -> VisualLayout:
        padding = 24

        # KPI row for first 4 card visuals.
        if total > 3 and index < 4:
            card_width = (self.CANVAS_WIDTH - padding * 5) / 4
            return VisualLayout(
                x=padding + index * (card_width + padding),
                y=padding,
                width=card_width,
                height=120,
                z_index=index,
            )

        adjusted_index = index - 4 if total > 3 else index
        y_offset = 160 if total > 3 else padding
        cols = 2
        width = (self.CANVAS_WIDTH - padding * (cols + 1)) / cols
        height = 250
        row, col = divmod(adjusted_index, cols)
        return VisualLayout(
            x=padding + col * (width + padding),
            y=y_offset + row * (height + padding),
            width=width,
            height=height,
            z_index=index,
        )

    def _find_ref(self, names: List[str], refs: List[PowerBIFieldRef], default_agg: Optional[str] = None) -> Optional[PowerBIFieldRef]:
        normalized = [self._normalize_name(x) for x in names]

        def match_score(ref: PowerBIFieldRef) -> int:
            col_norm = self._normalize_name(ref.column)
            # Auto measures are named Total Sales / Average Discount. For matching,
            # also compare against the underlying business column name.
            col_base = re.sub(r"^(total|average|avg|count of|distinct count)\s+", "", col_norm).strip()
            table_norm = self._normalize_name(ref.table)
            best = 0
            for idx, wanted in enumerate(normalized):
                priority = len(normalized) - idx
                if col_norm == wanted or col_base == wanted:
                    best = max(best, 1000 + priority)
                elif wanted in col_base.split() or col_base in wanted.split() or wanted in col_norm.split() or col_norm in wanted.split():
                    best = max(best, 700 + priority)
                elif wanted in col_base or col_base in wanted or wanted in col_norm or col_norm in wanted:
                    best = max(best, 500 + priority)
                elif wanted in table_norm:
                    best = max(best, 100 + priority)
            return best

        ranked = sorted(((match_score(ref), i, ref) for i, ref in enumerate(refs)), key=lambda x: (-x[0], x[1]))
        for score, _, ref in ranked:
            if score > 0:
                if default_agg and not ref.is_measure:
                    return PowerBIFieldRef(ref.table, ref.column, aggregation=default_agg, is_measure=ref.is_measure)
                return ref
        return refs[0] if refs else None

    def _is_known_numeric_column(self, table: str, column: str) -> bool:
        table_clean = self._clean_table_name(table).lower()
        col_clean = self._clean_column_name(column).lower()
        for tbl in self.tables or []:
            tname = self._clean_table_name(str(tbl.get("name") or tbl.get("table_name") or "Table")).lower()
            if tname != table_clean:
                continue
            for col in tbl.get("columns") or tbl.get("fields") or []:
                cname = self._clean_column_name(str(col.get("name") if isinstance(col, dict) else col)).lower()
                dtype = str(col.get("dataType") or col.get("data_type") or col.get("type") or "").lower() if isinstance(col, dict) else ""
                if cname == col_clean and (dtype in {"int64", "integer", "int", "double", "decimal", "number", "float", "currency"} or self._looks_numeric_field_name(cname)):
                    return True
        return False

    def _looks_like_weak_visual_set(self, visuals: List[PowerBIVisual]) -> bool:
        if not visuals:
            return True
        if all(v.visual_type == PowerBIVisualType.TABLE for v in visuals):
            return True
        if len(visuals) <= 2 and self.tables:
            return True
        return False

    def _title_from_roles(self, visual_type: PowerBIVisualType, roles: Dict[str, List[PowerBIFieldRef]]) -> str:
        category = (roles.get("Category") or roles.get("Rows") or roles.get("Values") or [])[:1]
        values = (roles.get("Y") or roles.get("Values") or [])[:1]
        if visual_type == PowerBIVisualType.CARD and values:
            return self._friendly_metric_title(values[0])
        if category and values and category[0].column != values[0].column:
            return f"{self._friendly_metric_title(values[0])} by {self._pretty_name(category[0].column)}"
        if values:
            return self._friendly_metric_title(values[0])
        if category:
            return self._pretty_name(category[0].column)
        return "Business Summary"

    def _friendly_metric_title(self, ref: PowerBIFieldRef) -> str:
        name = self._pretty_name(ref.column)
        if ref.aggregation:
            agg = self._agg_display(ref.aggregation)
            if agg == "SUM":
                return f"Total {name}"
            if agg == "AVG":
                return f"Average {name}"
            if agg == "COUNT":
                return f"Count of {name}"
        return name


    def _is_numeric_column(self, table: str, column: str) -> bool:
        """Return True when a column should be aggregated in visuals.

        This deliberately uses both model metadata and column-name heuristics.
        Some ThoughtSpot exports do not pass Power BI dataType metadata to the
        visual converter, but business fields like Sales and Profit still need
        explicit SUM aggregation in PBIP JSON.
        """
        return self._is_known_numeric_column(table, column) or self._looks_numeric_field_name(column)

    def _normalize_ref_for_role(self, role: str, ref: PowerBIFieldRef) -> PowerBIFieldRef:
        """Return a safe field ref for the visual role.

        Fixes the issue seen in Power BI where Profit/Sales/Quantity appears as
        "Count of Profit" instead of total value. If a numeric business column is
        used in a value role, we force an explicit aggregation in the PBIP JSON.
        Dimension buckets like Category, Series, Legend, Rows, Columns stay as columns.
        """
        if not ref or ref.is_measure or ref.aggregation:
            return ref

        numeric_value_roles = {
            "values", "value", "y", "x", "size", "tooltips",
            "data", "measure", "measures",
        }
        dimension_roles = {
            "category", "categories", "series", "legend", "rows", "columns",
            "axis", "details", "group", "groups", "breakdown", "values" if False else "__never__",
        }
        r = str(role or "").lower().strip()
        if r in dimension_roles:
            return ref

        if r in numeric_value_roles and self._is_numeric_column(ref.table, ref.column):
            return PowerBIFieldRef(
                table=ref.table,
                column=ref.column,
                aggregation=self._default_aggregation_for_column(ref.column),
                is_measure=False,
            )
        return ref


    @staticmethod
    def _looks_numeric_field_name(column: str) -> bool:
        """Heuristic for business metric columns when source dataType is missing."""
        name = str(column or "").lower().strip()
        if not name:
            return False

        # IDs, keys, codes and text/date dimensions must stay categorical.
        non_metric_tokens = [
            " id", "_id", "id ", "key", "code", "name", "category",
            "segment", "region", "country", "state", "city", "date",
            "product", "customer", "order date",
        ]
        if any(token in name for token in non_metric_tokens):
            # Keep explicit count metrics numeric, e.g. Order Count, but do not
            # mistake Country for Count.
            if not re.search(r"\b(count|quantity|qty|units)\b", name):
                return False

        metric_tokens = [
            "sales", "revenue", "profit", "margin", "cost", "amount",
            "price", "quantity", "qty", "discount", "rate", "target",
            "score", "value", "total", "units",
        ]
        return any(token in name for token in metric_tokens) or bool(re.search(r"\bcount\b", name))

    @staticmethod
    def _default_aggregation_for_column(column: str) -> str:
        name = str(column or "").lower()
        if any(k in name for k in ["discount", "margin", "rate", "ratio", "percent", "percentage", "%"]):
            return "avg"
        if any(k in name for k in ["id", "code", "number", "no"]):
            return "count_distinct"
        return "sum"

    @staticmethod
    def _agg_query_name(agg: Optional[str]) -> str:
        # Power BI generated PBIR usually uses Sum(...) / Average(...) style names.
        return {
            "sum": "Sum",
            "avg": "Average",
            "average": "Average",
            "min": "Min",
            "max": "Max",
            "count": "Count",
            "count_distinct": "DistinctCount",
        }.get((agg or "sum").lower(), "Sum")

    @staticmethod
    def _agg_code(agg: str) -> int:
        return {
            "sum": 0,
            "avg": 1,
            "average": 1,
            "min": 2,
            "max": 3,
            "count": 4,
            "count_distinct": 5,
        }.get((agg or "sum").lower(), 0)

    @staticmethod
    def _agg_display(agg: Optional[str]) -> str:
        return {
            "sum": "SUM",
            "avg": "AVG",
            "average": "AVG",
            "min": "MIN",
            "max": "MAX",
            "count": "COUNT",
            "count_distinct": "DISTINCTCOUNT",
        }.get((agg or "").lower(), (agg or "").upper())

    def _native_reference_name(self, ref: PowerBIFieldRef) -> str:
        if ref.is_measure:
            return self._pretty_name(ref.column)
        if ref.aggregation:
            agg = (ref.aggregation or "sum").lower()
            base = self._pretty_name(ref.column)
            if agg in {"avg", "average"}:
                return f"Average {base}"
            if agg == "count_distinct":
                return f"Distinct Count {base}"
            if agg == "count":
                return f"Count {base}"
            if agg == "min":
                return f"Minimum {base}"
            if agg == "max":
                return f"Maximum {base}"
            return f"Total {base}"
        return self._pretty_name(ref.column)

    @staticmethod
    def _clean_field_token(value: str) -> str:
        value = str(value or "").strip()
        value = re.sub(r"^ATTR\((.*)\)$", r"\1", value, flags=re.I)
        value = value.replace("[", "").replace("]", "")
        return value.strip().strip("'\"")

    @staticmethod
    def _clean_column_name(value: str) -> str:
        # IMPORTANT: Do not replace underscores with spaces. Must match TMDL column names exactly.
        value = str(value or "Field").strip().strip("'\"")
        return re.sub(r"\s+", " ", value).strip() or "Field"

    @staticmethod
    def _clean_table_name(value: str) -> str:
        # IMPORTANT: Do not replace spaces with underscores. Must match TMDL table names exactly.
        value = str(value or "Table").strip().strip("'\"")
        return re.sub(r"\s+", " ", value).strip() or "Table"

    @staticmethod
    def _safe_visual_name(value: str) -> str:
        value = str(value or "visual").strip()
        value = re.sub(r"[^A-Za-z0-9_]+", "_", value)
        value = re.sub(r"_+", "_", value).strip("_")
        return value or f"visual_{uuid.uuid4().hex[:8]}"

    @staticmethod
    def _clean_title(value: str) -> str:
        value = str(value or "Business Summary").strip()
        value = value.replace("_", " ")
        value = re.sub(r"\s+", " ", value)
        # Preserve common BI words.
        return " ".join([w.upper() if w.lower() in {"kpi", "dax"} else w.capitalize() for w in value.split()])

    @staticmethod
    def _pretty_name(value: str) -> str:
        value = str(value or "").strip()
        value = re.sub(r"^cf[_\s]+", "", value, flags=re.I)
        value = value.replace("_", " ")
        value = re.sub(r"\s+", " ", value).strip()
        return " ".join(w.capitalize() for w in value.split()) or "Value"

    @staticmethod
    def _normalize_name(value: str) -> str:
        return re.sub(r"[^a-z0-9]+", " ", str(value or "").lower()).strip()

    @staticmethod
    def _is_generic_title(value: str) -> bool:
        v = str(value or "").strip().lower()
        return v in {"sheet", "worksheet", "table", "chart", "visual", "summary"} or v.startswith("sheet ")

    @staticmethod
    def _dedupe_strings(values: List[str]) -> List[str]:
        seen, out = set(), []
        for v in values:
            key = str(v).lower().strip()
            if key and key not in seen:
                seen.add(key)
                out.append(v)
        return out

    @staticmethod
    def _dedupe_refs(values: List[PowerBIFieldRef]) -> List[PowerBIFieldRef]:
        seen, out = set(), []
        for r in values:
            key = (r.table.lower(), r.column.lower(), r.aggregation, r.is_measure)
            if key not in seen:
                seen.add(key)
                out.append(r)
        return out

    @staticmethod
    def _unique_name(base: str, seen: set[str]) -> str:
        name = base or "visual"
        if name not in seen:
            seen.add(name)
            return name
        i = 2
        while f"{name}_{i}" in seen:
            i += 1
        final = f"{name}_{i}"
        seen.add(final)
        return final


# -------------------------- backward compatible helpers --------------------------


def create_visual_converter(
    tables: Optional[List[Dict[str, Any]]] = None,
    calculated_fields: Optional[List[Dict[str, Any]]] = None,
) -> VisualConverter:
    return VisualConverter(tables=tables, calculated_fields=calculated_fields)


def convert_worksheets_to_visuals(
    worksheets: List[Dict[str, Any]],
    **kwargs: Any,
) -> List[PowerBIVisual]:
    tables = kwargs.pop("tables", None)
    calculated_fields = kwargs.pop("calculated_fields", None)
    return VisualConverter(tables=tables, calculated_fields=calculated_fields).convert_worksheets_to_visuals(
        worksheets,
        **kwargs,
    )


def generate_report_json(visuals: List[PowerBIVisual], page_name: str = "Executive Dashboard") -> Dict[str, Any]:
    converter = VisualConverter()
    return converter.generate_report_json([converter.generate_page_json(page_name, visuals)])
