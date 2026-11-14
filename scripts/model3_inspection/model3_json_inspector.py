#!/usr/bin/env python3
"""Read-only structural inspector for large JSON files using ijson streaming."""
from __future__ import annotations
import argparse, json
from pathlib import Path

def find_first_array_prefix(path: Path):
    import ijson
    with path.open("rb") as f:
        for prefix, event, value in ijson.parse(f):
            if event == "start_array":
                return prefix
    return None

def inspect_json(path: Path, sample_limit: int):
    import ijson
    array_prefix = find_first_array_prefix(path)
    sample = []
    record_count = 0
    type_counts = {}
    key_counts = {}
    examples = {}
    if array_prefix is None:
        # For a JSON object without arrays, report its top-level structure only.
        with path.open("rb") as f:
            top_type = None
            top_keys = []
            for prefix, event, value in ijson.parse(f):
                if prefix == "" and event in ("start_map", "start_array"):
                    top_type = "object" if event == "start_map" else "array"
                if event == "map_key" and prefix == "":
                    top_keys.append(str(value))
                    if len(top_keys) >= 200:
                        break
        return {
            "file": str(path), "size_bytes": path.stat().st_size,
            "top_level_type": top_type, "top_level_keys_preview": top_keys,
            "array_found": False,
            "note": "No JSON array was found. This is a structural preview, not a full semantic profile."
        }

    item_prefix = f"{array_prefix}.item" if array_prefix else "item"
    with path.open("rb") as f:
        for item in ijson.items(f, item_prefix):
            record_count += 1
            if len(sample) < sample_limit:
                sample.append(item)
            if isinstance(item, dict):
                type_counts["object"] = type_counts.get("object", 0) + 1
                for key, value in item.items():
                    key_counts[key] = key_counts.get(key, 0) + 1
                    if key not in examples and value is not None:
                        examples[key] = {
                            "python_type": type(value).__name__,
                            "example": str(value)[:160]
                        }
            elif isinstance(item, list):
                type_counts["array"] = type_counts.get("array", 0) + 1
            else:
                t = type(item).__name__
                type_counts[t] = type_counts.get(t, 0) + 1
    return {
        "file": str(path), "size_bytes": path.stat().st_size,
        "top_level_array_prefix": array_prefix or "<root>",
        "records_counted": record_count,
        "records_sampled": len(sample),
        "record_type_counts": type_counts,
        "object_key_presence_counts_in_full_array": key_counts,
        "first_non_null_value_examples_by_key": examples,
        "sample_records": sample,
        "note": "Array items are streamed; only the configured number of initial records are retained."
    }

def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--root", required=True, help="Root folder containing raw Model 3 files")
    ap.add_argument("--out", required=True, help="Output folder")
    ap.add_argument("--sample-records", type=int, default=5, help="Number of initial records to retain (default: 5)")
    ap.add_argument("--file", default=None, help="Optional exact JSON file path; otherwise inspect JSON files under root")
    args = ap.parse_args()
    root, out = Path(args.root), Path(args.out)
    if not root.exists():
        raise SystemExit(f"Input root does not exist: {root}")
    out.mkdir(parents=True, exist_ok=True)
    try:
        import ijson  # noqa: F401
    except ImportError:
        raise SystemExit("Missing dependency 'ijson'. Install it with: py -m pip install ijson")
    paths = [Path(args.file)] if args.file else sorted(
        p for p in root.rglob("*.json") if p.is_file()
        and not p.name.startswith("._") and p.name != ".DS_Store"
    )
    reports, errors = [], []
    for path in paths:
        try:
            reports.append(inspect_json(path, max(1, args.sample_records)))
            print(f"[OK] {path}")
        except Exception as exc:
            errors.append({"file": str(path), "error": f"{type(exc).__name__}: {exc}"})
            print(f"[ERROR] {path}: {exc}")
    payload = {
        "tool": "model3_json_inspector",
        "json_files_found": len(paths),
        "json_files_profiled": len(reports),
        "errors": errors,
        "files": reports,
        "limitations": [
            "Read-only; no source JSON is modified.",
            "For files containing multiple arrays, the first array encountered is profiled.",
            "Only the first configured records are retained as examples, although the selected array is counted.",
            "This tool does not interpret exercise taxonomy, muscle mapping, or scientific meaning."
        ]
    }
    target = out / "model3_json_structure_report.json"
    target.write_text(json.dumps(payload, indent=2, ensure_ascii=False, default=str), encoding="utf-8")
    print(f"\nSaved: {target}")

if __name__ == "__main__":
    main()
