#!/usr/bin/env python3
"""Read-only Excel workbook layout inspector for ShaktiKram Model 3."""
from __future__ import annotations
import argparse, json, re
from pathlib import Path
from datetime import date, datetime
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

SUPPORTED = {".xlsx", ".xlsm"}

def clean(value, limit=160):
    if value is None:
        return None
    if isinstance(value, (datetime, date)):
        value = value.isoformat()
    text = str(value).replace("\n", " ").replace("\r", " ").strip()
    return text[:limit] + ("…" if len(text) > limit else "")

def inspect_workbook(path: Path, max_rows: int, max_cols: int):
    result = {"file": str(path), "size_bytes": path.stat().st_size, "worksheets": []}
    wb = load_workbook(path, read_only=False, data_only=False)
    try:
        for ws in wb.worksheets:
            row_limit = min(ws.max_row or 0, max_rows)
            col_limit = min(ws.max_column or 0, max_cols)
            nonempty_rows = []
            nonempty_cols = {}
            for row in ws.iter_rows(min_row=1, max_row=row_limit,
                                    min_col=1, max_col=col_limit, values_only=False):
                cells = []
                for cell in row:
                    value = cell.value
                    if value is not None:
                        cells.append({
                            "cell": cell.coordinate,
                            "value": clean(value),
                            "data_type": cell.data_type
                        })
                        letter = get_column_letter(cell.column)
                        nonempty_cols[letter] = nonempty_cols.get(letter, 0) + 1
                if cells:
                    nonempty_rows.append({
                        "row": row[0].row,
                        "nonempty_cell_count": len(cells),
                        "cells": cells[:30],
                        "cells_truncated": len(cells) > 30
                    })
            merged = [str(rng) for rng in list(ws.merged_cells.ranges)[:200]]
            ws_result = {
                "sheet_name": ws.title,
                "sheet_state": ws.sheet_state,
                "max_row_reported": ws.max_row,
                "max_column_reported": ws.max_column,
                "inspected_rows": row_limit,
                "inspected_columns": col_limit,
                "freeze_panes": str(ws.freeze_panes) if ws.freeze_panes else None,
                "merged_ranges_first_200": merged,
                "nonempty_cell_count_by_column": nonempty_cols,
                "nonempty_rows_preview": nonempty_rows[:250],
                "rows_preview_truncated": len(nonempty_rows) > 250
            }
            ws_result["note"] = (
                "Row/cell preview is capped. Coordinates and displayed values are structural evidence only; "
                "this script does not decide which row is the true header."
            )
            result["worksheets"].append(ws_result)
    finally:
        wb.close()
    return result

def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--root", required=True, help="Root folder containing raw Model 3 files")
    ap.add_argument("--out", required=True, help="Output folder for inspection JSON")
    ap.add_argument("--max-rows", type=int, default=250, help="Maximum rows inspected per sheet (default: 250)")
    ap.add_argument("--max-cols", type=int, default=160, help="Maximum columns inspected per sheet (default: 160)")
    args = ap.parse_args()
    root, out = Path(args.root), Path(args.out)
    if not root.exists():
        raise SystemExit(f"Input root does not exist: {root}")
    out.mkdir(parents=True, exist_ok=True)
    workbooks = sorted(p for p in root.rglob("*")
                       if p.is_file() and p.suffix.lower() in SUPPORTED
                       and not p.name.startswith("._") and p.name != ".DS_Store")
    reports, errors = [], []
    for path in workbooks:
        try:
            reports.append(inspect_workbook(path, args.max_rows, args.max_cols))
            print(f"[OK] {path.relative_to(root)}")
        except Exception as exc:
            errors.append({"file": str(path), "error": f"{type(exc).__name__}: {exc}"})
            print(f"[ERROR] {path}: {exc}")
    payload = {
        "tool": "model3_excel_layout_inspector",
        "input_root": str(root),
        "workbooks_found": len(workbooks),
        "workbooks_profiled": len(reports),
        "errors": errors,
        "workbooks": reports,
        "limitations": [
            "Read-only inspection; source workbooks are not changed.",
            "Formula cells are inspected as formulas, not recalculated values.",
            "Only the first configured rows and columns are inspected.",
            "The script does not infer headers, units, timepoints, participant linkage, or scientific meaning."
        ]
    }
    target = out / "model3_excel_layout_report.json"
    target.write_text(json.dumps(payload, indent=2, ensure_ascii=False), encoding="utf-8")
    print(f"\nSaved: {target}")

if __name__ == "__main__":
    main()
