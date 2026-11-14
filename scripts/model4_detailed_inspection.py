#!/usr/bin/env python3
"""
ShaktiKram_ML — Model 4 Detailed Dataset Inspection
===================================================

READ-ONLY INSPECTION ONLY.
- Never edits, renames, deletes, or overwrites source datasets.
- Does not impute, recode, normalize, merge, or clean source data.
- Any conversion needed for legacy Excel files is performed on a TEMPORARY COPY.
- Reports are written outside the raw-data directory.

Expected project layout:
    F:\\Projects\\ShaktiKram_ML\\
        datasets\\body_composition\\raw\\
        datasets\\body_composition\\_inspection_reports\\first_pass\\
        datasets\\body_composition\\_inspection_reports\\detailed\\

Run from project root:
    python scripts\\model4_detailed_inspection.py

Optional:
    python scripts\\model4_detailed_inspection.py --root "F:\\Projects\\ShaktiKram_ML\\datasets\\body_composition\\raw"
    python scripts\\model4_detailed_inspection.py --output "F:\\Projects\\ShaktiKram_ML\\datasets\\body_composition\\_inspection_reports\\detailed"
    python scripts\\model4_detailed_inspection.py --soffice "C:\\Program Files\\LibreOffice\\program\\soffice.exe"

Dependencies:
    pip install pandas openpyxl xlrd pyreadstat pyarrow

Legacy BIFF2 .xls:
    LibreOffice (soffice) is required if pandas/xlrd cannot read the file.
    The script converts only a temporary copy to XLSX, never the raw file.
"""

from __future__ import annotations

import argparse
import csv
import hashlib
import json
import math
import re
import shutil
import subprocess
import sys
import tempfile
import traceback
from collections import Counter
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

try:
    import pandas as pd
except ImportError:
    print("ERROR: pandas is required. Install with: pip install pandas")
    raise

SUPPORTED_EXTENSIONS = {
    ".csv", ".tsv", ".txt", ".xlsx", ".xls", ".xlsm",
    ".xpt", ".sas7bdat", ".json", ".jsonl", ".ndjson",
    ".parquet", ".sav", ".dta",
}
EXCLUDED_DIRS = {"weee", "_registry", "_inspection_reports"}
MAX_TOP_VALUES = 15
MAX_EXACT_UNIQUE = 250_000
ENCODINGS = ("utf-8", "utf-8-sig", "cp1252", "latin1")


def safe(value: Any) -> Any:
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
    if isinstance(value, pd.Timestamp):
        return value.isoformat()
    if isinstance(value, float) and not math.isfinite(value):
        return str(value)
    if isinstance(value, (str, int, float, bool)):
        return value
    return str(value)


def relpath(path: Path, root: Path) -> str:
    return str(path.relative_to(root)).replace("\\", "/")


def dataset_group(relative: str) -> str:
    parts = Path(relative).parts
    if not parts:
        return "UNCLASSIFIED"
    first = parts[0]
    if first.upper().startswith("NHANES"):
        return first
    return first


def stable_hash(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def read_delimited(path: Path, sep: str | None = None) -> pd.DataFrame:
    last_error = None
    for encoding in ENCODINGS:
        try:
            return pd.read_csv(path, sep=sep, encoding=encoding, low_memory=False)
        except UnicodeDecodeError as exc:
            last_error = exc
    if last_error:
        raise last_error
    return pd.read_csv(path, sep=sep, low_memory=False)


def load_json_table(path: Path) -> tuple[pd.DataFrame, dict]:
    if path.suffix.lower() in {".jsonl", ".ndjson"}:
        return pd.read_json(path, lines=True), {"json_structure": "JSON Lines"}
    with path.open("r", encoding="utf-8-sig") as f:
        obj = json.load(f)
    meta = {
        "json_structure": type(obj).__name__,
        "top_level_keys": list(obj.keys())[:300] if isinstance(obj, dict) else [],
    }
    if isinstance(obj, list):
        if all(isinstance(x, dict) for x in obj):
            return pd.json_normalize(obj), meta
        return pd.DataFrame({"value": obj}), meta
    if isinstance(obj, dict):
        for key in ("data", "records", "results", "rows", "items"):
            value = obj.get(key)
            if isinstance(value, list):
                if value and all(isinstance(x, dict) for x in value):
                    return pd.json_normalize(value), meta
                return pd.DataFrame({"value": value}), meta
        try:
            return pd.DataFrame(obj), meta
        except Exception:
            return pd.json_normalize(obj), meta
    return pd.DataFrame({"value": [obj]}), meta


def convert_legacy_excel_temp(path: Path, temp_root: Path) -> tuple[Path, dict]:
    """
    Convert a COPY to XLSX with LibreOffice. Source path is read only.
    Returns converted path and conversion metadata.
    """
    soffice = SOFFICE_OVERRIDE or shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        # Common Windows installation paths, in case LibreOffice is not on PATH.
        windows_candidates = [
            Path(r"C:\\Program Files\\LibreOffice\\program\\soffice.exe"),
            Path(r"C:\\Program Files (x86)\\LibreOffice\\program\\soffice.exe"),
        ]
        soffice = next((str(candidate) for candidate in windows_candidates if candidate.exists()), None)
    if not soffice:
        raise RuntimeError(
            "Legacy Excel file could not be read directly and LibreOffice/soffice "
            "was not found on PATH or in common Windows installation folders. "
            "Install LibreOffice, add its program folder to PATH, or rerun with --soffice."
        )
    outdir = temp_root / "converted"
    outdir.mkdir(parents=True, exist_ok=True)
    temp_input_dir = temp_root / "input"
    temp_input_dir.mkdir(parents=True, exist_ok=True)
    copied_input = temp_input_dir / path.name
    shutil.copy2(path, copied_input)
    command = [
        soffice, "--headless", "--convert-to", "xlsx",
        "--outdir", str(outdir), str(copied_input),
    ]
    result = subprocess.run(command, capture_output=True, text=True, timeout=180)
    converted = outdir / (path.stem + ".xlsx")
    if result.returncode != 0 or not converted.exists():
        raise RuntimeError(
            "LibreOffice temporary conversion failed. "
            f"returncode={result.returncode}; stdout={result.stdout}; stderr={result.stderr}"
        )
    return converted, {
        "conversion_used": True,
        "conversion_tool": Path(soffice).name,
        "conversion_source": "temporary copy; original untouched",
        "conversion_stdout": result.stdout.strip(),
        "conversion_stderr": result.stderr.strip(),
    }


def load_table(path: Path, temp_root: Path) -> tuple[pd.DataFrame, dict]:
    ext = path.suffix.lower()
    meta: dict[str, Any] = {"reader": None}
    if ext == ".csv":
        df = read_delimited(path)
        meta["reader"] = "pandas.read_csv"
    elif ext == ".tsv":
        df = read_delimited(path, "\t")
        meta["reader"] = "pandas.read_csv(tab)"
    elif ext == ".txt":
        try:
            df = read_delimited(path, "\t")
            meta["reader"] = "pandas.read_csv(tab)"
        except Exception:
            df = read_delimited(path, ",")
            meta["reader"] = "pandas.read_csv(comma)"
    elif ext == ".xpt":
        df = pd.read_sas(path, format="xport", encoding="utf-8")
        meta["reader"] = "pandas.read_sas(xport)"
    elif ext == ".sas7bdat":
        df = pd.read_sas(path, format="sas7bdat", encoding="utf-8")
        meta["reader"] = "pandas.read_sas(sas7bdat)"
    elif ext in {".json", ".jsonl", ".ndjson"}:
        df, json_meta = load_json_table(path)
        meta.update(json_meta)
        meta["reader"] = "pandas.read_json/json_normalize"
    elif ext == ".parquet":
        df = pd.read_parquet(path)
        meta["reader"] = "pandas.read_parquet"
    elif ext == ".sav":
        try:
            import pyreadstat
        except ImportError as exc:
            raise RuntimeError("Install pyreadstat: pip install pyreadstat") from exc
        df, stat_meta = pyreadstat.read_sav(str(path), apply_value_formats=False)
        meta.update({
            "reader": "pyreadstat.read_sav",
            "file_label": getattr(stat_meta, "file_label", None),
            "variable_labels": getattr(stat_meta, "column_names_to_labels", {}),
            "value_labels": getattr(stat_meta, "variable_value_labels", {}),
        })
    elif ext == ".dta":
        df = pd.read_stata(path, convert_categoricals=False)
        meta["reader"] = "pandas.read_stata"
    else:
        raise RuntimeError(f"Not a single-table format: {ext}")
    return df, meta


def is_excel(path: Path) -> bool:
    return path.suffix.lower() in {".xls", ".xlsx", ".xlsm"}


def inspect_excel(path: Path, temp_root: Path) -> tuple[list, list, list, dict]:
    actual_path = path
    conversion_meta = {"conversion_used": False}
    try:
        # Try native reader first. Legacy BIFF2 may not be recognized by pandas.
        workbook = pd.ExcelFile(path)
    except Exception as first_error:
        if path.suffix.lower() != ".xls":
            raise
        actual_path, conversion_meta = convert_legacy_excel_temp(path, temp_root)
        try:
            workbook = pd.ExcelFile(actual_path, engine="openpyxl")
        except Exception as second_error:
            raise RuntimeError(
                f"Native Excel read failed ({first_error}); temporary conversion "
                f"also failed ({second_error})"
            ) from second_error

    table_rows, column_rows, sheet_rows = [], [], []
    try:
        for sheet in workbook.sheet_names:
            df = pd.read_excel(workbook, sheet_name=sheet)
            table, cols = profile_frame(df, sheet, {})
            table_rows.append(table)
            column_rows.extend(cols)
            sheet_rows.append({
                "sheet_name": str(sheet),
                "rows": int(df.shape[0]),
                "columns": int(df.shape[1]),
                "empty_sheet": bool(df.empty),
            })
    finally:
        workbook.close()

    # Inspect workbook metadata/formulas when openpyxl can read the file.
    excel_meta = {}
    try:
        import openpyxl
        wb = openpyxl.load_workbook(actual_path, read_only=False, data_only=False)
        formula_count = 0
        formula_examples = []
        hidden_sheets = []
        merged_ranges = []
        for ws in wb.worksheets:
            if ws.sheet_state != "visible":
                hidden_sheets.append(ws.title)
            if ws.merged_cells:
                merged_ranges.extend([f"{ws.title}!{str(r)}" for r in list(ws.merged_cells.ranges)[:100]])
            for row in ws.iter_rows():
                for cell in row:
                    if isinstance(cell.value, str) and cell.value.startswith("="):
                        formula_count += 1
                        if len(formula_examples) < 20:
                            formula_examples.append({
                                "sheet": ws.title,
                                "cell": cell.coordinate,
                                "formula": cell.value,
                            })
        excel_meta = {
            "formula_count": formula_count,
            "formula_examples": formula_examples,
            "hidden_sheets": hidden_sheets,
            "merged_cell_ranges_first_100": merged_ranges[:100],
        }
        wb.close()
    except Exception as exc:
        excel_meta = {"excel_metadata_error": f"{type(exc).__name__}: {exc}"}

    return table_rows, column_rows, sheet_rows, {**conversion_meta, **excel_meta}


def profile_frame(df: pd.DataFrame, table_name: str, file_meta: dict) -> tuple[dict, list]:
    nrows, ncols = df.shape
    table = {
        "table_or_sheet": table_name,
        "rows": int(nrows),
        "columns": int(ncols),
        "duplicate_full_rows": int(df.duplicated().sum()) if nrows else 0,
        "empty_table": bool(df.empty),
    }
    profiles = []
    for i, col in enumerate(df.columns):
        s = df.iloc[:, i]
        nonnull = s.dropna()
        missing = int(s.isna().sum())
        try:
            nunique = int(nonnull.nunique()) if len(nonnull) <= MAX_EXACT_UNIQUE else f">{MAX_EXACT_UNIQUE}; skipped"
        except Exception:
            nunique = "not computed"
        samples = [safe(v) for v in nonnull.head(5).tolist()]
        rec = {
            "table_or_sheet": table_name,
            "column_position": i + 1,
            "column_name": str(col),
            "dtype": str(s.dtype),
            "rows": nrows,
            "non_missing": int(len(nonnull)),
            "missing": missing,
            "missing_pct": round(missing / nrows * 100, 4) if nrows else None,
            "unique_non_missing": nunique,
            "sample_values": json.dumps(samples, ensure_ascii=False),
            "constant_non_missing": bool(nonnull.nunique(dropna=True) <= 1) if len(nonnull) else False,
        }
        if pd.api.types.is_numeric_dtype(s):
            x = pd.to_numeric(s, errors="coerce").dropna()
            if len(x):
                quantiles = x.quantile([0.01, 0.05, 0.25, 0.5, 0.75, 0.95, 0.99])
                rec.update({
                    "min": safe(x.min()), "max": safe(x.max()),
                    "mean": safe(x.mean()), "std": safe(x.std()),
                    "q01": safe(quantiles.loc[0.01]), "q05": safe(quantiles.loc[0.05]),
                    "q25": safe(quantiles.loc[0.25]), "median": safe(quantiles.loc[0.5]),
                    "q75": safe(quantiles.loc[0.75]), "q95": safe(quantiles.loc[0.95]),
                    "q99": safe(quantiles.loc[0.99]),
                    "zero_count": int((x == 0).sum()),
                    "negative_count": int((x < 0).sum()),
                })
        is_numeric = pd.api.types.is_numeric_dtype(s)
        if len(nonnull) and (not is_numeric or (isinstance(nunique, int) and nunique <= 50)):
            try:
                counts = nonnull.astype(str).value_counts(dropna=False).head(MAX_TOP_VALUES)
                rec["top_value_counts"] = json.dumps(
                    [{"value": str(k), "count": int(v)} for k, v in counts.items()],
                    ensure_ascii=False,
                )
            except Exception:
                rec["top_value_counts"] = ""
        profiles.append(rec)
    return table, profiles


def candidate_id_columns(df: pd.DataFrame) -> list[dict]:
    results = []
    for col in df.columns:
        name = str(col).strip().lower()
        if name in {
            "seqn", "id", "id_sample", "subjno", "subject", "subject_id",
            "participant_id", "participant", "study_id", "record_id",
        } or name.endswith("_id") or name.endswith("id"):
            s = df[col]
            nonnull = s.dropna()
            results.append({
                "column": str(col),
                "non_missing": int(len(nonnull)),
                "unique_non_missing": int(nonnull.nunique()) if len(nonnull) else 0,
                "duplicate_non_missing_values": int(nonnull.duplicated().sum()) if len(nonnull) else 0,
                "missing": int(s.isna().sum()),
                "unique_ratio_non_missing": round(nonnull.nunique() / len(nonnull), 6) if len(nonnull) else None,
            })
    return results


def duplicate_examples(df: pd.DataFrame, limit: int = 25) -> list[dict]:
    if df.empty:
        return []
    mask = df.duplicated(keep=False)
    if not mask.any():
        return []
    dup = df.loc[mask].head(limit)
    return json.loads(dup.to_json(orient="records", date_format="iso"))


def write_csv(path: Path, rows: list[dict]):
    if not rows:
        path.write_text("", encoding="utf-8")
        return
    fields, seen = [], set()
    for row in rows:
        for key in row:
            if key not in seen:
                fields.append(key)
                seen.add(key)
    with path.open("w", newline="", encoding="utf-8-sig") as f:
        writer = csv.DictWriter(f, fieldnames=fields, extrasaction="ignore")
        writer.writeheader()
        writer.writerows(rows)


def classify_file(relative: str) -> tuple[str, str | None]:
    name = Path(relative).stem.lower()
    group = dataset_group(relative)
    if group.upper().startswith("NHANES"):
        module = None
        for candidate in ("bmx", "demo", "dxx"):
            if re.search(rf"(^|[^a-z]){candidate}([^a-z]|$)", name):
                module = candidate.upper()
                break
            if name.startswith(candidate):
                module = candidate.upper()
                break
        return group, module
    return group, None


def inspect_nhanes_crossfiles(loaded_frames: list[dict]) -> tuple[list, list]:
    """Compare NHANES BMX/DEMO/DXX IDs within cycles and module schemas across cycles."""
    id_rows, schema_rows = [], []
    by_cycle = {}
    for item in loaded_frames:
        group, module = classify_file(item["relative_file"])
        if not group.upper().startswith("NHANES") or module not in {"BMX", "DEMO", "DXX"}:
            continue
        df = item["dataframe"]
        seqn_col = next((c for c in df.columns if str(c).upper() == "SEQN"), None)
        by_cycle.setdefault(group, {})[module] = (df, item["relative_file"], seqn_col)
        if seqn_col is not None:
            ids = df[seqn_col].dropna()
            id_rows.append({
                "cycle": group, "module": module, "file": item["relative_file"],
                "rows": len(df), "seqn_non_missing": len(ids),
                "seqn_unique": int(ids.nunique()),
                "seqn_duplicate_values": int(ids.duplicated().sum()),
                "seqn_missing": int(df[seqn_col].isna().sum()),
            })
    for cycle, modules in by_cycle.items():
        module_sets = {}
        for module, (df, file, idcol) in modules.items():
            if idcol is not None:
                module_sets[module] = set(df[idcol].dropna().tolist())
        all_modules = [m for m in ("DEMO", "BMX", "DXX") if m in module_sets]
        for i, left in enumerate(all_modules):
            for right in all_modules[i + 1:]:
                a, b = module_sets[left], module_sets[right]
                id_rows.append({
                    "cycle": cycle, "module": f"{left} vs {right}",
                    "file": "", "rows": None, "seqn_non_missing": None,
                    "seqn_unique": None, "seqn_duplicate_values": None,
                    "seqn_missing": None,
                    "intersection_ids": len(a & b),
                    "left_only_ids": len(a - b),
                    "right_only_ids": len(b - a),
                    "union_ids": len(a | b),
                })
    # Schema presence comparison by module across cycles.
    modules_to_cycles = {}
    for item in loaded_frames:
        group, module = classify_file(item["relative_file"])
        if group.upper().startswith("NHANES") and module in {"BMX", "DEMO", "DXX"}:
            modules_to_cycles.setdefault(module, {})[group] = set(map(str, item["dataframe"].columns))
    for module, cycle_map in modules_to_cycles.items():
        all_cols = set().union(*cycle_map.values()) if cycle_map else set()
        for col in sorted(all_cols):
            present = [cycle for cycle, cols in cycle_map.items() if col in cols]
            schema_rows.append({
                "module": module,
                "column_name": col,
                "cycles_present": ";".join(sorted(present)),
                "number_of_cycles_present": len(present),
                "number_of_cycles_observed": len(cycle_map),
                "present_in_all_observed_cycles": len(present) == len(cycle_map),
            })
    return id_rows, schema_rows


def inspect_colombian(loaded_frames: list[dict]) -> tuple[list, list]:
    key_rows, relationship_rows = [], []
    candidates = [x for x in loaded_frames if "colombian" in x["relative_file"].lower()]
    if not candidates:
        return key_rows, relationship_rows
    # Usually one workbook; sheets are stored as separate loaded-frame entries.
    sheet_frames = {}
    for item in candidates:
        sheet_frames[item["table_or_sheet"]] = item["dataframe"]
        for col in ("id", "id_Sample", "id_sample"):
            if col in item["dataframe"].columns:
                s = item["dataframe"][col].dropna()
                key_rows.append({
                    "sheet": item["table_or_sheet"], "key_column": col,
                    "rows": len(item["dataframe"]), "non_missing": len(s),
                    "unique_values": int(s.nunique()),
                    "duplicate_values": int(s.duplicated().sum()),
                })
    keysets = {}
    for sheet, df in sheet_frames.items():
        for col in ("id", "id_Sample", "id_sample"):
            if col in df.columns:
                keysets[(sheet, col)] = set(df[col].dropna().tolist())
    keys = list(keysets)
    for i, left in enumerate(keys):
        for right in keys[i + 1:]:
            if left[1].lower() == right[1].lower() or left[1].lower() == "id" or right[1].lower() == "id":
                a, b = keysets[left], keysets[right]
                relationship_rows.append({
                    "left_sheet": left[0], "left_key": left[1],
                    "right_sheet": right[0], "right_key": right[1],
                    "intersection_count": len(a & b),
                    "left_only_count": len(a - b),
                    "right_only_count": len(b - a),
                    "union_count": len(a | b),
                    "interpretation": "ID-set overlap only; relationship is not asserted",
                })
    return key_rows, relationship_rows


def inspect_rancho(loaded_frames: list[dict]) -> tuple[list, list]:
    rows, visit_rows = [], []
    candidates = [x for x in loaded_frames if "rancho" in x["relative_file"].lower()]
    for item in candidates:
        df = item["dataframe"]
        idcol = next((c for c in df.columns if str(c).lower() in {"subjno", "subject_id", "id"}), None)
        if idcol:
            s = df[idcol].dropna()
            rows.append({
                "file": item["relative_file"], "identifier_column": str(idcol),
                "rows": len(df), "non_missing_ids": len(s),
                "unique_ids": int(s.nunique()),
                "duplicate_id_values": int(s.duplicated().sum()),
            })
        # Identify suffix patterns such as V7, V8, V9, V10 without interpreting variables.
        visit_map = {}
        for col in df.columns:
            match = re.search(r"V(7|8|9|10)$", str(col), flags=re.IGNORECASE)
            if match:
                visit = "V" + match.group(1)
                visit_map.setdefault(visit, []).append(col)
        for visit, cols in sorted(visit_map.items()):
            visit_rows.append({
                "file": item["relative_file"], "visit_suffix": visit,
                "columns_with_suffix": len(cols),
                "columns": ";".join(map(str, cols)),
                "rows_with_any_non_missing_value": int(df[cols].notna().any(axis=1).sum()),
                "rows_with_all_missing_values": int(df[cols].isna().all(axis=1).sum()),
            })
    return rows, visit_rows


def inspect_university(loaded_frames: list[dict]) -> tuple[list, list]:
    categorical_rows, field_rows = [], []
    candidates = [x for x in loaded_frames if "university_club_sports" in x["relative_file"].lower()]
    categorical_names = {"sport", "position", "sex", "ethnicity", "athtype", "filter_$"}
    for item in candidates:
        df = item["dataframe"]
        idcol = next((c for c in df.columns if str(c).lower() == "id"), None)
        if idcol:
            s = df[idcol].dropna()
            field_rows.append({
                "file": item["relative_file"], "identifier_column": str(idcol),
                "rows": len(df), "non_missing_ids": len(s),
                "unique_ids": int(s.nunique()),
                "duplicate_id_values": int(s.duplicated().sum()),
            })
        for col in df.columns:
            if str(col).strip().lower() in categorical_names:
                counts = df[col].value_counts(dropna=False).head(100)
                categorical_rows.append({
                    "file": item["relative_file"], "column": str(col),
                    "value_counts": json.dumps(
                        [{"value": safe(k), "count": int(v)} for k, v in counts.items()],
                        ensure_ascii=False,
                    ),
                    "note": "Code meanings require source documentation; not inferred.",
                })
    return field_rows, categorical_rows


SOFFICE_OVERRIDE = None


def main():
    global SOFFICE_OVERRIDE
    parser = argparse.ArgumentParser(description="Detailed read-only Model 4 inspection")
    parser.add_argument("--root", type=Path, default=Path("datasets/body_composition/raw"))
    parser.add_argument(
        "--output", type=Path,
        default=Path("datasets/body_composition/_inspection_reports/detailed"),
    )
    parser.add_argument(
        "--soffice", type=str, default=None,
        help="Optional full path to LibreOffice soffice.exe for legacy .xls conversion.",
    )
    args = parser.parse_args()
    SOFFICE_OVERRIDE = args.soffice
    root = args.root.resolve()
    output = args.output.resolve()

    if not root.is_dir():
        print(f"ERROR: Raw-data directory not found: {root}")
        sys.exit(2)
    if output == root or root in output.parents:
        print("ERROR: Report output must be outside the raw-data directory.")
        sys.exit(2)

    output.mkdir(parents=True, exist_ok=True)
    report_dirs = {
        "nhanes": output / "nhanes",
        "colombian": output / "colombian",
        "rancho": output / "rancho_bernardo",
        "university": output / "university_club_sports",
        "cross": output / "cross_dataset",
    }
    for folder in report_dirs.values():
        folder.mkdir(parents=True, exist_ok=True)

    files = []
    for p in sorted(root.rglob("*")):
        if not p.is_file():
            continue
        relative = p.relative_to(root)
        if any(part.lower() in EXCLUDED_DIRS for part in relative.parts[:-1]):
            continue
        if p.suffix.lower() in SUPPORTED_EXTENSIONS:
            files.append(p)

    inventory, table_rows, column_rows, errors = [], [], [], []
    loaded_frames = []
    print(f"Raw root: {root}")
    print(f"Output: {output}")
    print(f"Files discovered: {len(files)}")
    print("READ-ONLY: source files will not be modified.")

    with tempfile.TemporaryDirectory(prefix="skx_model4_inspection_") as temp_name:
        temp_root = Path(temp_name)
        for index, path in enumerate(files, 1):
            relative = relpath(path, root)
            group = dataset_group(relative)
            print(f"[{index}/{len(files)}] {relative}")
            base = {
                "dataset_group": group,
                "relative_file": relative,
                "extension": path.suffix.lower(),
                "size_bytes": path.stat().st_size,
                "sha256": stable_hash(path),
                "status": "PENDING",
                "error": "",
            }
            try:
                if is_excel(path):
                    tables, cols, sheets, excel_meta = inspect_excel(path, temp_root / f"file_{index}")
                    base.update(excel_meta)
                    base["table_count"] = len(tables)
                    for table in tables:
                        table["relative_file"] = relative
                        table["dataset_group"] = group
                    for col in cols:
                        col["relative_file"] = relative
                        col["dataset_group"] = group
                    for sheet in sheets:
                        sheet["relative_file"] = relative
                        sheet["dataset_group"] = group
                    # Reload each sheet as a DataFrame for targeted relationship checks.
                    read_path = path
                    if base.get("conversion_used"):
                        # Recreate conversion in this temp session for the dataframes.
                        read_path, _ = convert_legacy_excel_temp(path, temp_root / f"reload_{index}")
                    book = pd.ExcelFile(read_path)
                    for sheet_name in book.sheet_names:
                        frame = pd.read_excel(book, sheet_name=sheet_name)
                        loaded_frames.append({
                            "relative_file": relative, "table_or_sheet": str(sheet_name),
                            "dataframe": frame,
                        })
                    book.close()
                else:
                    df, meta = load_table(path, temp_root / f"file_{index}")
                    table, cols = profile_frame(df, "(single table)", meta)
                    table.update({"relative_file": relative, "dataset_group": group})
                    for col in cols:
                        col["relative_file"] = relative
                        col["dataset_group"] = group
                    tables, sheets = [table], []
                    base["reader_metadata"] = meta
                    base["table_count"] = 1
                    loaded_frames.append({
                        "relative_file": relative, "table_or_sheet": "(single table)",
                        "dataframe": df,
                    })
                table_rows.extend(tables)
                column_rows.extend(cols)
                base["status"] = "SUCCESS"
                base["rows_across_tables"] = sum(x["rows"] for x in tables)
                base["columns_across_tables"] = sum(x["columns"] for x in tables)
                base["candidate_identifiers"] = json.dumps([
                    {"table_or_sheet": item["table_or_sheet"],
                     "candidates": candidate_id_columns(item["dataframe"])}
                    for item in loaded_frames if item["relative_file"] == relative
                ], ensure_ascii=False)
            except Exception as exc:
                error = f"{type(exc).__name__}: {exc}"
                base.update({"status": "ERROR", "error": error})
                errors.append({
                    "relative_file": relative, "error": error,
                    "traceback": traceback.format_exc(),
                })
                print(f"    ERROR: {error}")
            inventory.append(base)

    # Dataset-specific analyses. These compare observed structures only.
    nhanes_ids, nhanes_schema = inspect_nhanes_crossfiles(loaded_frames)
    colombian_keys, colombian_relationships = inspect_colombian(loaded_frames)
    rancho_ids, rancho_visits = inspect_rancho(loaded_frames)
    university_ids, university_categories = inspect_university(loaded_frames)

    # Exact duplicate examples are reported separately, never removed.
    dup_rows = []
    for item in loaded_frames:
        df = item["dataframe"]
        examples = duplicate_examples(df)
        dup_rows.append({
            "relative_file": item["relative_file"],
            "table_or_sheet": item["table_or_sheet"],
            "duplicate_full_row_count": int(df.duplicated().sum()),
            "duplicate_examples_first_25": json.dumps(examples, ensure_ascii=False),
        })

    write_csv(output / "file_inventory_detailed.csv", inventory)
    write_csv(output / "table_sheet_inventory_detailed.csv", table_rows)
    write_csv(output / "column_profiles_detailed.csv", column_rows)
    write_csv(output / "inspection_errors.csv", errors)
    write_csv(output / "duplicate_row_audit.csv", dup_rows)

    write_csv(report_dirs["nhanes"] / "seqn_integrity_and_overlap.csv", nhanes_ids)
    write_csv(report_dirs["nhanes"] / "cross_cycle_variable_presence.csv", nhanes_schema)
    write_csv(report_dirs["colombian"] / "identifier_profiles.csv", colombian_keys)
    write_csv(report_dirs["colombian"] / "cross_sheet_id_overlap.csv", colombian_relationships)
    write_csv(report_dirs["rancho"] / "participant_id_profile.csv", rancho_ids)
    write_csv(report_dirs["rancho"] / "visit_suffix_profile.csv", rancho_visits)
    write_csv(report_dirs["university"] / "participant_id_profile.csv", university_ids)
    write_csv(report_dirs["university"] / "coded_category_frequencies.csv", university_categories)

    successful = sum(x["status"] == "SUCCESS" for x in inventory)
    failed = sum(x["status"] == "ERROR" for x in inventory)
    summary = {
        "project": "ShaktiKram_ML",
        "model": "Model 4 - Body Composition",
        "inspection_stage": "detailed",
        "inspection_type": "read-only structural, descriptive, and dataset-specific diagnostics",
        "generated_at_utc": datetime.now(timezone.utc).isoformat(),
        "raw_root": str(root),
        "output_directory": str(output),
        "excluded_directory_names": sorted(EXCLUDED_DIRS),
        "files_discovered": len(files),
        "files_successfully_inspected": successful,
        "files_with_errors": failed,
        "tables_or_sheets_profiled": len(table_rows),
        "columns_profiled": len(column_rows),
        "sha256_hashes_recorded": True,
        "dataset_specific_checks": [
            "NHANES: SEQN uniqueness, pairwise ID overlap among recognized BMX/DEMO/DXX modules, and variable presence across cycles",
            "Colombian: id/id_Sample key profiles and observed cross-sheet ID-set overlap",
            "Rancho Bernardo: candidate participant ID profile and V7-V10 suffix/availability profile",
            "University Club Sports: ID profile and value frequencies for selected coded/category-like fields",
            "All tables: duplicate full-row counts and up to 25 example duplicate rows",
            "All columns: missingness, unique counts, samples, numeric percentiles, zeros, negatives, and frequent values",
            "Excel: sheet visibility, formula count/examples, and merged-cell ranges when openpyxl can inspect the workbook",
        ],
        "important_interpretation_limits": [
            "Candidate identifiers are not declared valid keys by this script.",
            "ID overlaps show set relationships only; they do not establish valid joins or participant linkage.",
            "NHANES module recognition depends on filenames containing BMX, DEMO, or DXX.",
            "Cycle schema presence does not establish variable equivalence, measurement comparability, or valid pooling.",
            "Missingness is descriptive; the script does not infer why values are missing.",
            "Category frequencies do not decode numeric codes; code meanings require source documentation.",
            "Visit suffixes are detected syntactically and are not interpreted as verified visit dates.",
            "Duplicate records are flagged and sampled but never removed.",
            "Excel formulas are counted from workbook formula cells; formulas are not recalculated.",
            "Legacy workbook conversion is temporary and does not change the raw source.",
            "No preprocessing, imputation, normalization, filtering, merging, or model training is performed.",
        ],
        "dataset_group_counts": dict(Counter(x["dataset_group"] for x in inventory)),
    }
    (output / "detailed_inspection_summary.json").write_text(
        json.dumps(summary, indent=2, ensure_ascii=False), encoding="utf-8"
    )

    print("\nDetailed inspection finished.")
    print(f"Successful files: {successful}")
    print(f"Files with errors: {failed}")
    print(f"Tables/sheets: {len(table_rows)}")
    print(f"Columns profiled: {len(column_rows)}")
    print(f"Reports: {output}")
    print("Raw source datasets were not modified.")


if __name__ == "__main__":
    main()
