SHAKIKRAM_ML — MODEL 3 DETAILED INSPECTION TOOLKIT
=================================================

Purpose
-------
These are read-only inspection scripts. They do NOT clean, transform, merge,
rename, overwrite, or otherwise modify source datasets.

Place the three .py scripts and this README in:
F:\Projects\ShaktiKram_ML\scripts\model3_inspection\

Open PowerShell in the project root:
cd F:\Projects\ShaktiKram_ML

Install dependencies:
py -m pip install openpyxl ijson

1) Inspect Excel workbook layout
--------------------------------
py scripts\model3_inspection\model3_excel_layout_inspector.py --root "datasets\muscle_development\raw" --out "datasets\_inspection\model_3"

Output:
datasets\_inspection\model_3\model3_excel_layout_report.json

2) Inspect JSON files using streaming
-------------------------------------
py scripts\model3_inspection\model3_json_inspector.py --root "datasets\muscle_development\raw" --out "datasets\_inspection\model_3" --sample-records 5

Output:
datasets\_inspection\model_3\model3_json_structure_report.json

To inspect only the Exercise Muscle Volume JSON:
py scripts\model3_inspection\model3_json_inspector.py --root "datasets\muscle_development\raw" --out "datasets\_inspection\model_3" --file "datasets\muscle_development\raw\Exercise_Muscle_Volume\data\exercises.json" --sample-records 8

3) Inspect raw cell-level data quality
--------------------------------------
py scripts\model3_inspection\model3_data_quality_inspector.py --root "datasets\muscle_development\raw" --out "datasets\_inspection\model_3"

Output:
datasets\_inspection\model_3\model3_raw_data_quality_report.json

Expected reports
----------------
- model3_excel_layout_report.json
- model3_json_structure_report.json
- model3_raw_data_quality_report.json

Important notes
---------------
- Ignore macOS metadata files such as .DS_Store and ._*; the scripts exclude them.
- Excel layout inspector reports coordinates, merged cells, hidden sheets and
  capped row/cell previews. It does not guess the header row.
- JSON inspector uses ijson to stream the first array found in each JSON file.
  If a JSON contains several arrays, only the first encountered array is profiled.
- Data-quality inspector is header-agnostic. Blank rates and exact duplicate
  row patterns are structural clues only, not scientific conclusions.
- Review reports before sharing; sample values may contain personal or
  participant-level information. Remove/redact any sensitive values first.
- No preprocessing should begin until dataset-specific specifications and
  validation rules have been reviewed and approved.
