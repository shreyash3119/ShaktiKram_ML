#!/usr/bin/env python3
"""Read-only, header-agnostic cell quality inspection for Excel workbooks."""
from __future__ import annotations
import argparse, json
from pathlib import Path
from collections import Counter
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

def inspect_sheet(ws, max_rows: int, max_cols: int, duplicate_preview_limit: int):
    rows = min(ws.max_row or 0, max_rows)
    cols = min(ws.max_column or 0, max_cols)
    nonempty_by_col = Counter()
    blank_by_col = Counter()
    row_signatures = Counter()
    row_nonempty_counts = []
    for row in ws.iter_rows(min_row=1, max_row=rows, min_col=1, max_col=cols, values_only=True):
        values = tuple("" if v is None else str(v).strip() for v in row)
        nonempty = sum(v != "" for v in values)
        row_nonempty_counts.append(nonempty)
        for idx, value in enumerate(values, start=1):
            letter = get_column_letter(idx)
            if value == "":
                blank_by_col[letter] += 1
            else:
                nonempty_by_col[letter] += 1
        if nonempty:
            row_signatures[values] += 1
    duplicate_rows = [
        {"occurrences": count, "row_values_preview": list(signature[:20])}
        for signature, count in row_signatures.items() if count > 1
    ]
    duplicate_rows.sort(key=lambda x: x["occurrences"], reverse=True)
    duplicate_rows = duplicate_rows[:duplicate_preview_limit]
    denom = rows
    columns = []
    for idx in range(1, cols + 1):
        letter = get_column_letter(idx)
        nn = nonempty_by_col[letter]
        columns.append({
            "column_letter": letter,
            "nonempty_cells": nn,
            "blank_cells": blank_by_col[letter],
            "blank_rate": round(blank_by_col[letter] / denom, 4) if denom else None
        })
    return {
        "sheet_name": ws.title,
        "rows_considered_including_blank_rows": rows,
        "columns_considered": cols,
        "nonempty_cells_by_row_min": min(row_nonempty_counts) if row_nonempty_counts else 0,
        "nonempty_cells_by_row_max": max(row_nonempty_counts) if row_nonempty_counts else 0,
        "columns": columns,
        "exact_duplicate_nonempty_row_patterns_preview": duplicate_rows,
        "duplicate_patterns_preview_count": len(duplicate_rows),
        "note": "This is a raw cell-level profile. It does not assume a header row or participant-ID column."
    }

def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--root", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--max-rows", type=int, default=10000)
    ap.add_argument("--max-cols", type=int, default=200)
    ap.add_argument("--duplicate-preview", type=int, default=25)
    args = ap.parse_args()
    root, out = Path(args.root), Path(args.out)
    if not root.exists():
        raise SystemExit(f"Input root does not exist: {root}")
    out.mkdir(parents=True, exist_ok=True)
    paths = sorted(p for p in root.rglob("*")
                   if p.is_file() and p.suffix.lower() in {".xlsx", ".xlsm"}
                   and not p.name.startswith("._") and p.name != ".DS_Store")
    reports, errors = [], []
    for path in paths:
        try:
            wb = load_workbook(path, read_only=True, data_only=True)
            try:
                reports.append({
                    "file": str(path),
                    "sheets": [inspect_sheet(ws, args.max_rows, args.max_cols, args.duplicate_preview)
                               for ws in wb.worksheets]
                })
            finally:
                wb.close()
            print(f"[OK] {path.relative_to(root)}")
        except Exception as exc:
            errors.append({"file": str(path), "error": f"{type(exc).__name__}: {exc}"})
            print(f"[ERROR] {path}: {exc}")
    payload = {
        "tool": "model3_data_quality_inspector",
        "input_root": str(root),
        "workbooks_found": len(paths),
        "workbooks_profiled": len(reports),
        "errors": errors,
        "workbooks": reports,
        "limitations": [
            "Read-only; source workbooks are not modified.",
            "Blank rates are calculated over the inspected rectangular range, including blank rows.",
            "Exact duplicate row patterns are not proof of duplicate participants or invalid records.",
            "No participant IDs, headers, units, timepoints, or scientific meanings are inferred."
        ]
    }
    target = out / "model3_raw_data_quality_report.json"
    target.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")
    print(f"\nSaved: {target}")

if __name__ == "__main__":
    main()
