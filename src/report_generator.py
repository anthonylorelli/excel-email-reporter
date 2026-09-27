"""Report generator using Jinja2 templates and consolidated spreadsheet data."""

from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional
import pandas as pd
from jinja2 import Environment, FileSystemLoader, Template


class ReportGenerator:
    """Generates formatted HTML and plain text email reports from consolidated data."""

    def __init__(self, config: Any):
        self.config = config

    def calculate_metrics(self, df: pd.DataFrame) -> List[Dict[str, Any]]:
        """Computes summary metrics defined in configuration."""
        metric_configs = getattr(self.config, "summary_metrics", [])
        results: List[Dict[str, Any]] = []

        if df.empty:
            return results

        for m in metric_configs:
            col = m.get("column")
            agg_type = m.get("type", "sum")
            fmt = m.get("format", "{}")
            label = m.get("label", col)

            if col not in df.columns:
                continue

            try:
                series = pd.to_numeric(df[col], errors="coerce").dropna()
                val = 0
                if agg_type == "sum":
                    val = series.sum()
                elif agg_type == "mean":
                    val = series.mean()
                elif agg_type == "count":
                    val = len(series)
                elif agg_type == "min":
                    val = series.min()
                elif agg_type == "max":
                    val = series.max()

                formatted_val = fmt.format(val)
                results.append({
                    "id": m.get("id"),
                    "label": label,
                    "raw_value": val,
                    "formatted_value": formatted_val,
                })
            except Exception:
                pass

        return results

    def calculate_breakdown(self, df: pd.DataFrame) -> Dict[str, Any]:
        """Calculates grouped breakdown table based on configuration."""
        breakdown_config = getattr(self.config, "breakdown", {})
        group_by = breakdown_config.get("group_by")

        if not group_by or df.empty or group_by not in df.columns:
            return {}

        metric_defs = breakdown_config.get("metrics", [])
        headers = [m.get("header", m.get("column")) for m in metric_defs]

        # Group data
        rows: List[Dict[str, Any]] = []
        for group_val, group_df in df.groupby(group_by):
            row_vals: List[str] = []
            for m in metric_defs:
                col = m.get("column")
                agg = m.get("agg", "sum")
                fmt = m.get("format", "{}")
                if col in group_df.columns:
                    series = pd.to_numeric(group_df[col], errors="coerce").dropna()
                    val = 0
                    if agg == "sum":
                        val = series.sum()
                    elif agg == "mean":
                        val = series.mean()
                    elif agg == "count":
                        val = len(series)
                    row_vals.append(fmt.format(val))
                else:
                    row_vals.append("-")

            rows.append({
                "group": group_val,
                "cells": row_vals,
            })

        return {
            "group_by": group_by,
            "headers": headers,
            "rows": rows,
        }

    def prepare_detail_preview(self, df: pd.DataFrame, max_rows: int = 15) -> List[Dict[str, Any]]:
        """Prepares a preview subset of records for table display."""
        if df.empty:
            return []
        # Filter out internal columns
        display_cols = [c for c in df.columns if not c.startswith("_")]
        sample_df = df[display_cols].head(max_rows)
        return sample_df.to_dict(orient="records")

    def render(
        self,
        df: pd.DataFrame,
        ingest_summaries: List[Any],
        title: Optional[str] = None,
    ) -> Dict[str, str]:
        """Renders both HTML and plain text email bodies."""
        title = title or getattr(self.config, "project", {}).get("name", "Executive Excel Report")
        metrics = self.calculate_metrics(df)
        breakdown = self.calculate_breakdown(df)
        detail_rows = self.prepare_detail_preview(df)
        detail_headers = [c for c in df.columns if not c.startswith("_")] if not df.empty else []

        numeric_columns = [
            c for c in detail_headers
            if not df.empty and pd.api.types.is_numeric_dtype(df[c])
        ]

        files_data = [
            {
                "filename": s.filename,
                "row_count": s.included_count,
            }
            for s in ingest_summaries
        ]

        context = {
            "title": title,
            "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "metrics": metrics,
            "breakdown": breakdown,
            "detail_rows": detail_rows,
            "detail_headers": detail_headers,
            "numeric_columns": numeric_columns,
            "files": files_data,
        }

        # Render HTML
        html_template_path = getattr(self.config, "template_html", "templates/report_template.html")
        html_body = self._render_template(html_template_path, context)

        # Render Plain Text
        text_template_path = getattr(self.config, "template_text", "templates/report_template.txt")
        text_body = self._render_template(text_template_path, context)

        return {
            "html": html_body,
            "text": text_body,
        }

    def _render_template(self, template_path: str, context: Dict[str, Any]) -> str:
        """Helper to render a template file using Jinja2."""
        path = Path(template_path)
        if path.exists():
            env = Environment(loader=FileSystemLoader(str(path.parent)))
            env.filters["ljust"] = lambda s, width: str(s).ljust(width)
            env.filters["rjust"] = lambda s, width: str(s).rjust(width)
            template = env.get_template(path.name)
            return template.render(**context)

        # Basic fallback template if file does not exist
        fallback_str = """
        <h2>{{ title }}</h2>
        <p>Generated at: {{ generated_at }}</p>
        {% for m in metrics %}
        <p><strong>{{ m.label }}:</strong> {{ m.formatted_value }}</p>
        {% endfor %}
        """
        return Template(fallback_str).render(**context)
