from pathlib import Path
import argparse, csv, json, math
from collections import Counter
from datetime import datetime, timezone

SUPPORTED = {".csv", ".tsv", ".txt", ".xlsx", ".xlsm", ".json", ".jsonl", ".parquet"}

def safe_value(v):
    if v is None:
        return None
    s = str(v).replace("\r", " ").replace("\n", " ")
    return s[:180]

def profile_delimited(path, sample_rows=10000):
    import pandas as pd
    sep = "\t" if path.suffix.lower() == ".tsv" else None
    # Read a bounded sample first; do not load large files wholly into memory.
    df = pd.read_csv(path, sep=sep, engine="python", nrows=sample_rows, on_bad_lines="warn")
    return profile_frame(df, path, sampled=True, sample_rows=sample_rows)

def profile_excel(path):
    import pandas as pd
    xls = pd.ExcelFile(path)
    results = []
    for sheet in xls.sheet_names:
        df = pd.read_excel(path, sheet_name=sheet, nrows=10000)
        results.append(profile_frame(df, path, sheet=sheet, sampled=True, sample_rows=10000))
    return results

def profile_frame(df, path, sheet=None, sampled=True, sample_rows=10000):
    cols = []
    for c in df.columns:
        s = df[c]
        examples = []
        for v in s.dropna().head(5).tolist():
            examples.append(safe_value(v))
        cols.append({
            "column": str(c),
            "inferred_dtype": str(s.dtype),
            "non_null_in_sample": int(s.notna().sum()),
            "null_in_sample": int(s.isna().sum()),
            "unique_in_sample": int(s.nunique(dropna=True)),
            "example_values": examples
        })
    return {
        "file": str(path),
        "sheet": sheet,
        "sampled": sampled,
        "sample_rows_limit": sample_rows,
        "rows_in_sample": int(len(df)),
        "columns_count": int(len(df.columns)),
        "columns": cols,
        "duplicate_rows_in_sample": int(df.duplicated().sum()) if len(df) else 0
    }

def profile_json(path, max_bytes=2*1024*1024):
    size = path.stat().st_size
    result = {"file": str(path), "file_size_bytes": size, "mode": "structure_only"}
    if size > max_bytes:
        result["note"] = "File exceeds JSON preview limit; not loaded. Inspect with a dedicated streaming parser after format is confirmed."
        return result
    try:
        text = path.read_text(encoding="utf-8-sig")
        if path.suffix.lower() == ".jsonl":
            items = []
            for line in text.splitlines()[:100]:
                if line.strip():
                    items.append(json.loads(line))
            result.update({"format": "jsonl", "sample_records": len(items),
                           "sample_record_keys": sorted({str(k) for x in items if isinstance(x, dict) for k in x.keys()})})
        else:
            obj = json.loads(text)
            result["top_level_type"] = type(obj).__name__
            if isinstance(obj, dict):
                result["top_level_keys"] = list(obj.keys())[:200]
                result["top_level_key_count"] = len(obj)
            elif isinstance(obj, list):
                result["top_level_length"] = len(obj)
                result["first_item_type"] = type(obj[0]).__name__ if obj else None
                if obj and isinstance(obj[0], dict):
                    result["first_item_keys"] = list(obj[0].keys())[:200]
    except Exception as e:
        result["error"] = f"{type(e).__name__}: {e}"
    return result

def main():
    ap = argparse.ArgumentParser(description="Read-only bounded profiler for Model 3 tabular/JSON files.")
    ap.add_argument("--root", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--max-json-mb", type=float, default=2.0)
    args = ap.parse_args()
    root, out = Path(args.root).resolve(), Path(args.out).resolve()
    out.mkdir(parents=True, exist_ok=True)
    records, errors = [], []
    candidates = [p for p in sorted(root.rglob("*")) if p.is_file() and p.suffix.lower() in SUPPORTED]
    for p in candidates:
        try:
            ext = p.suffix.lower()
            if ext in {".csv", ".tsv", ".txt"}:
                records.append(profile_delimited(p))
            elif ext in {".xlsx", ".xlsm"}:
                records.extend(profile_excel(p))
            elif ext in {".json", ".jsonl"}:
                records.append(profile_json(p, int(args.max_json_mb*1024*1024)))
            elif ext == ".parquet":
                import pandas as pd
                df = pd.read_parquet(p)
                records.append(profile_frame(df, p, sampled=False, sample_rows=len(df)))
        except Exception as e:
            errors.append({"file": str(p), "error": f"{type(e).__name__}: {e}"})
    report = {
        "dataset_root": str(root),
        "generated_utc": datetime.now(timezone.utc).isoformat(),
        "profiled_file_or_sheet_count": len(records),
        "files_considered": len(candidates),
        "profiles": records,
        "errors": errors,
        "important_limitations": [
            "Delimited files are profiled using at most the first 10,000 rows; counts are sample counts, not full-file counts.",
            "Large JSON files are not loaded when they exceed the configured preview limit.",
            "No values are converted, normalized, merged, or written back to source files.",
            "No participant linkage or scientific interpretation is attempted."
        ]
    }
    out_path = out / "model3_table_profiles.json"
    out_path.write_text(json.dumps(report, indent=2, ensure_ascii=False), encoding="utf-8")
    print(f"Profile complete: {len(records)} file/sheet profiles; {len(errors)} errors")
    print(f"Report: {out_path}")

if __name__ == "__main__":
    main()
