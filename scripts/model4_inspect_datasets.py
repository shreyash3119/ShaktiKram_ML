"""
Model 4 - Body Composition: Read-only dataset inspector
========================================================
Scans supported files under datasets/body_composition/raw, excluding WEEE.
It writes inspection reports to datasets/body_composition/_inspection_reports.
It NEVER edits, renames, moves, deletes, or saves over source dataset files.

Recommended environment:
    pip install pandas openpyxl xlrd pyreadstat pyarrow

Run from project root:
    python scripts/model4_inspect_datasets.py

Optional:
    python scripts/model4_inspect_datasets.py --root "F:\\Projects\\ShaktiKram_ML\\datasets\\body_composition\\raw"
    python scripts/model4_inspect_datasets.py --output "F:\\Projects\\ShaktiKram_ML\\datasets\\body_composition\\_inspection_reports"
"""

from __future__ import annotations

import argparse
import csv
import json
import math
import sys
import traceback
from collections import Counter
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

try:
    import pandas as pd
except ImportError:
    print("ERROR: pandas is required. Install it with: pip install pandas")
    raise


SUPPORTED_EXTENSIONS = {
    ".csv", ".tsv", ".txt",
    ".xlsx", ".xls", ".xlsm",
    ".xpt", ".sas7bdat",
    ".json", ".jsonl", ".ndjson",
    ".parquet",
    ".sav", ".dta",
}

EXCLUDED_DIRECTORY_NAMES = {
    "weee",
    "_inspection_reports",
    "_registry",
}

MAX_SAMPLE_VALUES = 5
MAX_UNIQUE_EXACT = 100_000
CSV_ENCODING_CANDIDATES = ("utf-8", "utf-8-sig", "cp1252", "latin1")


def json_safe(value: Any) -> Any:
    """Convert pandas/numpy/date values into JSON-safe Python values."""
    if value is None:
        return None
    try:
        if pd.isna(value):
            return None
    except (TypeError, ValueError):
        pass
    if hasattr(value, "item"):
        try:
            value = value.item()
        except Exception:
            pass
    if isinstance(value, (pd.Timestamp,)):
        return value.isoformat()
    if isinstance(value, (str, int, float, bool)):
        if isinstance(value, float) and not math.isfinite(value):
            return str(value)
        return value
    return str(value)


def infer_dataset_group(relative_path: Path) -> str:
    parts = [p for p in relative_path.parts if p not in ("", ".")]
    if not parts:
        return "UNCLASSIFIED"
    first = parts[0]
    if first.startswith("NHANES_"):
        return f"NHANES / {first}"
    return first


def read_csv_safely(path: Path, delimiter: str | None = None) -> pd.DataFrame:
    last_error = None
    for encoding in CSV_ENCODING_CANDIDATES:
        try:
            return pd.read_csv(
                path,
                sep=delimiter,
                encoding=encoding,
                low_memory=False,
                on_bad_lines="error",
            )
        except UnicodeDecodeError as exc:
            last_error = exc
        except Exception:
            raise
    raise last_error or RuntimeError("Could not determine CSV encoding")


def dataframe_from_json(path: Path) -> tuple[pd.DataFrame | None, dict]:
    """Read JSON/JSONL while retaining a structural description."""
    if path.suffix.lower() in {".jsonl", ".ndjson"}:
        df = pd.read_json(path, lines=True)
        return df, {"json_structure": "newline-delimited JSON"}
    with path.open("r", encoding="utf-8-sig") as handle:
        obj = json.load(handle)

    structure = {
        "json_structure": type(obj).__name__,
        "top_level_keys": list(obj.keys())[:200] if isinstance(obj, dict) else [],
    }
    if isinstance(obj, list):
        if all(isinstance(item, dict) for item in obj):
            return pd.json_normalize(obj), structure
        return pd.DataFrame({"value": obj}), structure
    if isinstance(obj, dict):
        # A dict of equal-length lists/arrays is naturally tabular.
        try:
            df = pd.DataFrame(obj)
            if len(df.columns) > 0:
                return df, structure
        except Exception:
            pass
        # Common API/data-export structures.
        for key in ("data", "records", "results", "rows", "items"):
            value = obj.get(key)
            if isinstance(value, list):
                if all(isinstance(item, dict) for item in value):
                    return pd.json_normalize(value), structure
                return pd.DataFrame({"value": value}), structure
        return pd.json_normalize(obj), structure
    return pd.DataFrame({"value": [obj]}), structure


def load_single_table(path: Path) -> tuple[pd.DataFrame, dict]:
    ext = path.suffix.lower()
    metadata: dict[str, Any] = {}

    if ext == ".csv":
        return read_csv_safely(path), metadata
    if ext == ".tsv":
        return read_csv_safely(path, delimiter="\t"), metadata
    if ext == ".txt":
        # Treat TXT as tab-delimited first; fallback to comma-delimited.
        try:
            return read_csv_safely(path, delimiter="\t"), {"text_delimiter_attempt": "tab"}
        except Exception:
            return read_csv_safely(path, delimiter=","), {"text_delimiter_attempt": "comma"}
    if ext in {".xlsx", ".xls", ".xlsm"}:
        raise RuntimeError("Workbook must be inspected sheet-by-sheet")
    if ext == ".xpt":
        return pd.read_sas(path, format="xport", encoding="utf-8"), metadata
    if ext == ".sas7bdat":
        return pd.read_sas(path, format="sas7bdat", encoding="utf-8"), metadata
    if ext in {".json", ".jsonl", ".ndjson"}:
        df, metadata = dataframe_from_json(path)
        if df is None:
            raise RuntimeError("JSON did not produce a tabular view")
        return df, metadata
    if ext == ".parquet":
        return pd.read_parquet(path), metadata
    if ext == ".sav":
        try:
            import pyreadstat
        except ImportError as exc:
            raise RuntimeError("Install pyreadstat to inspect .sav files: pip install pyreadstat") from exc
        df, meta = pyreadstat.read_sav(str(path), apply_value_formats=False)
        metadata["source_format"] = "SPSS SAV"
        metadata["file_label"] = getattr(meta, "file_label", None)
        metadata["variable_labels"] = getattr(meta, "column_names_to_labels", {})
        return df, metadata
    if ext == ".dta":
        return pd.read_stata(path, convert_categoricals=False), metadata

    raise RuntimeError(f"Unsupported extension: {ext}")


def profile_dataframe(
    df: pd.DataFrame,
    file_path: Path,
    relative_path: Path,
    table_name: str,
    metadata: dict,
) -> tuple[dict, list[dict]]:
    n_rows, n_cols = df.shape
    table_record = {
        "dataset_group": infer_dataset_group(relative_path),
        "relative_file": str(relative_path),
        "table_or_sheet": table_name,
        "rows": int(n_rows),
        "columns": int(n_cols),
        "duplicate_rows": int(df.duplicated().sum()) if n_rows else 0,
        "metadata": metadata,
    }

    profiles: list[dict] = []
    for position, col in enumerate(df.columns):
        series = df.iloc[:, position]
        non_null = series.dropna()
        missing_count = int(series.isna().sum())
        unique_count: int | str
        if len(non_null) <= MAX_UNIQUE_EXACT:
            unique_count = int(non_null.nunique(dropna=True))
        else:
            unique_count = f">{MAX_UNIQUE_EXACT} (exact count skipped)"

        samples = []
        for value in non_null.head(MAX_SAMPLE_VALUES).tolist():
            safe = json_safe(value)
            samples.append(safe)

        record = {
            "dataset_group": infer_dataset_group(relative_path),
            "relative_file": str(relative_path),
            "table_or_sheet": table_name,
            "column_position_1_based": position + 1,
            "column_name": str(col),
            "pandas_dtype": str(series.dtype),
            "row_count": int(n_rows),
            "non_null_count": int(n_rows - missing_count),
            "missing_count": missing_count,
            "missing_percent": round((missing_count / n_rows) * 100, 4) if n_rows else None,
            "unique_non_null_count": unique_count,
            "sample_values_first_5": json.dumps(samples, ensure_ascii=False),
        }

        if pd.api.types.is_numeric_dtype(series):
            numeric = pd.to_numeric(series, errors="coerce").dropna()
            if len(numeric):
                record.update({
                    "numeric_min": json_safe(numeric.min()),
                    "numeric_max": json_safe(numeric.max()),
                    "numeric_mean": json_safe(numeric.mean()),
                    "numeric_median": json_safe(numeric.median()),
                    "numeric_std": json_safe(numeric.std()),
                    "zero_count": int((numeric == 0).sum()),
                    "negative_count": int((numeric < 0).sum()),
                })
        elif pd.api.types.is_datetime64_any_dtype(series):
            dates = pd.to_datetime(series, errors="coerce").dropna()
            if len(dates):
                record["datetime_min"] = json_safe(dates.min())
                record["datetime_max"] = json_safe(dates.max())

        profiles.append(record)

    return table_record, profiles


def inspect_workbook(path: Path, relative_path: Path):
    workbook_tables = []
    all_profiles = []
    sheet_inventory = []

    # ExcelFile opens workbook metadata without altering the source.
    with pd.ExcelFile(path) as workbook:
        for sheet_name in workbook.sheet_names:
            df = pd.read_excel(workbook, sheet_name=sheet_name)
            table, profiles = profile_dataframe(
                df, path, relative_path, str(sheet_name), {}
            )
            workbook_tables.append(table)
            all_profiles.extend(profiles)
            sheet_inventory.append({
                "relative_file": str(relative_path),
                "sheet_name": str(sheet_name),
                "rows": int(df.shape[0]),
                "columns": int(df.shape[1]),
                "hidden_status": "not inspected by pandas",
            })
    return workbook_tables, all_profiles, sheet_inventory


def inspect_file(path: Path, root: Path):
    relative_path = path.relative_to(root)
    ext = path.suffix.lower()
    size_bytes = path.stat().st_size
    modified = datetime.fromtimestamp(path.stat().st_mtime, timezone.utc).isoformat()

    file_record = {
        "dataset_group": infer_dataset_group(relative_path),
        "relative_file": str(relative_path),
        "extension": ext,
        "size_bytes": size_bytes,
        "size_mib": round(size_bytes / (1024 * 1024), 4),
        "modified_utc": modified,
        "inspection_status": "PENDING",
        "table_count": 0,
        "total_rows_across_tables": 0,
        "total_columns_across_tables": 0,
        "error": "",
    }

    if ext in {".xlsx", ".xls", ".xlsm"}:
        tables, profiles, sheets = inspect_workbook(path, relative_path)
    else:
        df, metadata = load_single_table(path)
        table, profiles = profile_dataframe(
            df, path, relative_path, "(single table)", metadata
        )
        tables, sheets = [table], []

    file_record["inspection_status"] = "SUCCESS"
    file_record["table_count"] = len(tables)
    file_record["total_rows_across_tables"] = sum(t["rows"] for t in tables)
    file_record["total_columns_across_tables"] = sum(t["columns"] for t in tables)
    return file_record, tables, profiles, sheets


def write_csv(path: Path, rows: list[dict]):
    if not rows:
        path.write_text("", encoding="utf-8")
        return
    # Stable union of keys across heterogeneous records.
    headers = []
    seen = set()
    for row in rows:
        for key in row:
            if key not in seen:
                headers.append(key)
                seen.add(key)
    with path.open("w", newline="", encoding="utf-8-sig") as handle:
        writer = csv.DictWriter(handle, fieldnames=headers, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)


def main():
    parser = argparse.ArgumentParser(description="Read-only Model 4 dataset inspector")
    parser.add_argument(
        "--root",
        type=Path,
        default=Path("datasets/body_composition/raw"),
        help="Model 4 raw-data directory (default: datasets/body_composition/raw)",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=Path("datasets/body_composition/_inspection_reports"),
        help="Report output directory (default: datasets/body_composition/_inspection_reports)",
    )
    args = parser.parse_args()

    root = args.root.resolve()
    output = args.output.resolve()

    if not root.exists() or not root.is_dir():
        print(f"ERROR: Raw dataset directory not found: {root}")
        print("Run this script from the project root or pass --root explicitly.")
        sys.exit(2)

    # Prevent reports from being written inside the raw directory.
    if output == root or root in output.parents:
        print("ERROR: Report output must be outside the raw dataset directory.")
        sys.exit(2)

    output.mkdir(parents=True, exist_ok=True)

    files = []
    for path in sorted(root.rglob("*")):
        if not path.is_file():
            continue
        relative = path.relative_to(root)
        if any(part.lower() in EXCLUDED_DIRECTORY_NAMES for part in relative.parts[:-1]):
            continue
        if path.suffix.lower() not in SUPPORTED_EXTENSIONS:
            continue
        files.append(path)

    inventory = []
    tables = []
    profiles = []
    sheets = []
    errors = []

    print(f"Raw root: {root}")
    print(f"Report output: {output}")
    print(f"Supported files discovered (WEEE excluded): {len(files)}")
    print("Inspection is read-only. Source files will not be modified.")

    for index, path in enumerate(files, start=1):
        relative = path.relative_to(root)
        print(f"[{index}/{len(files)}] Inspecting: {relative}")
        try:
            file_record, table_records, column_records, sheet_records = inspect_file(path, root)
            inventory.append(file_record)
            tables.extend(table_records)
            profiles.extend(column_records)
            sheets.extend(sheet_records)
        except Exception as exc:
            error_text = f"{type(exc).__name__}: {exc}"
            inventory.append({
                "dataset_group": infer_dataset_group(relative),
                "relative_file": str(relative),
                "extension": path.suffix.lower(),
                "size_bytes": path.stat().st_size,
                "size_mib": round(path.stat().st_size / (1024 * 1024), 4),
                "modified_utc": datetime.fromtimestamp(path.stat().st_mtime, timezone.utc).isoformat(),
                "inspection_status": "ERROR",
                "table_count": 0,
                "total_rows_across_tables": 0,
                "total_columns_across_tables": 0,
                "error": error_text,
            })
            errors.append({
                "relative_file": str(relative),
                "error": error_text,
                "traceback": traceback.format_exc(),
            })
            print(f"    ERROR: {error_text}")

    write_csv(output / "file_inventory.csv", inventory)
    write_csv(output / "table_and_sheet_inventory.csv", tables)
    write_csv(output / "column_profiles.csv", profiles)
    write_csv(output / "workbook_sheet_inventory.csv", sheets)
    write_csv(output / "inspection_errors.csv", errors)

    summary = {
        "project": "ShaktiKram_ML",
        "model": "Model 4 - Body Composition",
        "inspection_type": "read-only structural and descriptive inspection",
        "generated_at_utc": datetime.now(timezone.utc).isoformat(),
        "raw_root": str(root),
        "report_output": str(output),
        "excluded_directory_names": sorted(EXCLUDED_DIRECTORY_NAMES),
        "supported_extensions": sorted(SUPPORTED_EXTENSIONS),
        "files_discovered": len(files),
        "files_successfully_inspected": sum(1 for x in inventory if x["inspection_status"] == "SUCCESS"),
        "files_with_errors": sum(1 for x in inventory if x["inspection_status"] == "ERROR"),
        "tables_or_sheets_inspected": len(tables),
        "columns_profiled": len(profiles),
        "important_limitations": [
            "This is an initial structural/descriptive inspection, not a preprocessing step.",
            "No source dataset is modified.",
            "No records are merged across files or studies.",
            "No missing values are imputed or recoded.",
            "No outliers are removed and no values are normalized.",
            "Pandas may infer data types while reading; source storage formats remain unchanged.",
            "Workbook hidden-sheet status and Excel formulas are not audited in this first pass.",
            "Value labels and detailed metadata for every statistical file format may require a format-specific follow-up.",
            "A successful read does not establish scientific validity or cross-study compatibility.",
        ],
        "dataset_group_counts": dict(Counter(x["dataset_group"] for x in inventory)),
    }
    (output / "inspection_summary.json").write_text(
        json.dumps(summary, indent=2, ensure_ascii=False), encoding="utf-8"
    )

    print("\nInspection finished.")
    print(f"Successful files: {summary['files_successfully_inspected']}")
    print(f"Files with errors: {summary['files_with_errors']}")
    print(f"Tables/sheets: {summary['tables_or_sheets_inspected']}")
    print(f"Columns profiled: {summary['columns_profiled']}")
    print(f"Reports saved to: {output}")
    print("No raw dataset files were changed.")


if __name__ == "__main__":
    main()
