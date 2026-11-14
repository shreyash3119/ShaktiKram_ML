SHAKTIKRAM ML — MODEL 3 LOCAL DATASET INSPECTION TOOLKIT
Purpose
-------
Create a file inventory and bounded structural profiles locally, so you do not need to upload
a complete dataset archive or individual raw data files to ChatGPT.

Project paths
-------------
Project root: F:\Projects\ShaktiKram_ML
Model 3 raw data: F:\Projects\ShaktiKram_ML\datasets\muscle_development\raw
Inspection outputs: F:\Projects\ShaktiKram_ML\datasets\_inspection\model_3
Scripts location: F:\Projects\ShaktiKram_ML\scripts

Safety
------
- Read-only: these scripts do not edit, rename, move, or delete source files.
- Inventory records paths, extensions, sizes and timestamps. SHA-256 is optional.
- CSV/TSV/TXT profiling samples at most the first 10,000 rows.
- Excel profiling inspects sheets and at most 10,000 rows per sheet.
- JSON files larger than the configured preview limit are not loaded.
- No participant IDs are joined across datasets.
- No scientific feature engineering, unit conversion, scoring or preprocessing occurs.
- Reports may contain column names and sample values. Review them for sensitive information
  before sharing.

Install
-------
Open PowerShell in F:\Projects\ShaktiKram_ML and run:
    py -m pip install pandas openpyxl pyarrow

Run inventory:
    py scripts\model3_inspection\model3_inventory.py --root "datasets\muscle_development\raw" --out "datasets\_inspection\model_3"

Optional SHA-256 inventory (can take longer):
    py scripts\model3_inspection\model3_inventory.py --root "datasets\muscle_development\raw" --out "datasets\_inspection\model_3" --hash

Run structural profiling:
    py scripts\model3_inspection\model3_profile_tables.py --root "datasets\muscle_development\raw" --out "datasets\_inspection\model_3"

Outputs
-------
datasets\_inspection\model_3\model3_file_inventory.csv
datasets\_inspection\model_3\model3_inventory_summary.json
datasets\_inspection\model_3\model3_table_profiles.json

What to send back
-----------------
Share the three generated reports, not the raw dataset files:
- model3_file_inventory.csv
- model3_inventory_summary.json
- model3_table_profiles.json

The profiler intentionally does not interpret PDF/DOCX contents, proprietary formats, or scientific
binary files. We will inspect documentation separately if required, and create format-specific
profilers only when the inventory shows they are needed.
