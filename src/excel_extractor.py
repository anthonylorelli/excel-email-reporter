"""Excel data extractor and aggregator."""

import glob
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple
import pandas as pd


@dataclass
class FileIngestSummary:
    filename: str
    filepath: str
    row_count: int
    included_count: int


class ExcelExtractor:
    """Extracts, cleans, and consolidates tabular data from Excel spreadsheets."""

    def __init__(self, data_sources_config: Dict[str, Any]):
        self.config = data_sources_config
        self.sheet_name = self.config.get("sheet_name", 0)
        self.header_row = max(0, self.config.get("header_row", 1) - 1)  # 0-indexed for pandas
        self.column_mapping = self.config.get("columns", {})
        self.filters = self.config.get("filters", {})

    def discover_files(self, explicit_files: Optional[List[str]] = None) -> List[Path]:
        """Finds all target Excel files to process."""
        if explicit_files:
            return [Path(p) for p in explicit_files if Path(p).exists()]

        # Configured explicit files
        config_files = self.config.get("files")
        if config_files:
            return [Path(p) for p in config_files if Path(p).exists()]

        # Directory scanning
        directory = self.config.get("directory", "data")
        pattern = self.config.get("file_pattern", "*.xlsx")
        search_path = Path(directory) / pattern

        found = [Path(p) for p in glob.glob(str(search_path))]
        # Filter out temporary Excel lock files (starting with ~$)
        return [p for p in found if not p.name.startswith("~$")]

    def extract_from_file(self, filepath: Path) -> Tuple[pd.DataFrame, int]:
        """Reads a single Excel file and returns cleaned DataFrame and raw row count."""
        try:
            # Read sheet
            df = pd.read_excel(filepath, sheet_name=self.sheet_name, header=self.header_row)
        except Exception as e:
            # Fallback to first sheet if named sheet fails
            df = pd.read_excel(filepath, sheet_name=0, header=self.header_row)

        raw_count = len(df)

        # Invert column mapping: {excel_header: internal_name}
        rename_map = {v: k for k, v in self.column_mapping.items()}
        # Rename columns that exist in the dataframe
        df = df.rename(columns={col: rename_map[col] for col in df.columns if col in rename_map})

        # Apply filtering
        for filter_col, allowed_values in self.filters.items():
            if filter_col in df.columns:
                if isinstance(allowed_values, list):
                    df = df[df[filter_col].isin(allowed_values)]
                else:
                    df = df[df[filter_col] == allowed_values]

        # Drop columns that are completely empty / unnamed
        unnamed_empty = [c for c in df.columns if str(c).startswith("Unnamed:") and df[c].isna().all()]
        if unnamed_empty:
            df = df.drop(columns=unnamed_empty)

        # Add origin metadata columns
        df["_source_file"] = filepath.name
        df["_course_name"] = self.extract_course_name(filepath)

        return df, raw_count

    def extract_course_name(self, filepath: Path) -> str:
        """Extracts course title from the header row or filename."""
        try:
            raw_top = pd.read_excel(filepath, sheet_name=self.sheet_name, nrows=1, header=None)
            for val in raw_top.iloc[0]:
                if pd.notna(val) and "assignment data" in str(val).lower():
                    return str(val).split(" - ")[0].strip()
        except Exception:
            pass

        # Fallback to filename parsing
        stem = filepath.stem
        if " grades" in stem:
            return stem.split(" grades")[0].strip()
        return stem

    def process_all(self, explicit_files: Optional[List[str]] = None) -> Tuple[pd.DataFrame, List[FileIngestSummary]]:
        """Processes all discovered spreadsheets and aggregates results."""
        files = self.discover_files(explicit_files)
        if not files:
            return pd.DataFrame(), []

        dfs: List[pd.DataFrame] = []
        summaries: List[FileIngestSummary] = []

        for f in files:
            try:
                df, raw_count = self.extract_from_file(f)
                dfs.append(df)
                summaries.append(
                    FileIngestSummary(
                        filename=f.name,
                        filepath=str(f),
                        row_count=raw_count,
                        included_count=len(df),
                    )
                )
            except Exception as ex:
                summaries.append(
                    FileIngestSummary(
                        filename=f.name,
                        filepath=str(f),
                        row_count=0,
                        included_count=0,
                    )
                )

        if not dfs:
            return pd.DataFrame(), summaries

        consolidated_df = pd.concat(dfs, ignore_index=True)
        return consolidated_df, summaries

    def export_consolidated_excel(self, df: pd.DataFrame, output_path: str | Path) -> Path:
        """Exports the consolidated DataFrame to an Excel file with basic formatting."""
        out = Path(output_path)
        out.parent.mkdir(parents=True, exist_ok=True)

        with pd.ExcelWriter(out, engine="openpyxl") as writer:
            # Exclude internal metadata columns for the exported user-facing spreadsheet
            export_cols = [c for c in df.columns if not c.startswith("_")]
            df[export_cols].to_excel(writer, sheet_name="Consolidated_Report", index=False)

        return out
