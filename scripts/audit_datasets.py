#!/usr/bin/env python3
"""
ShaktiKram ML - dataset audit (v2)

Walks  <root>/<model_folder>/raw/<dataset>/**  and writes audit files to <root>/_audit/.
It never modifies anything under raw/.

Fixes over v1
  1. macOS '._*' AppleDouble files (and .DS_Store, Thumbs.db, __MACOSX) are ignored and
     listed once in ignored_files.csv instead of flooding audit_errors.csv.
  2. model_folder is derived from every file path (first path component).
  3. Known-harmless problems (e.g. WEEE empty tags.csv, ragged Zephyr event logs) are written
     to audit_benign_issues.csv. Ragged CSVs are re-read tolerantly, so they still get a
     registry row and column stats. Only real failures go to audit_errors.csv.

Extras
  * folder_check.csv   - does every expected dataset folder exist / contain files?
                         (catches e.g. University_Club_Sports missing or empty)
  * file_inventory.csv - EVERY file, with how it was handled, so nothing is skipped silently
                         (archives, .sav, .parquet, etc. are flagged as 'unsupported')
  * Header / date quality flags: MULTI_ROW_HEADER_SUSPECTED, WIDE_REPEATED_COLUMNS,
    suspicious dates (before 1990 or more than a year in the future)
  * --dump-docs <dataset>  writes the text of README/info/txt files (de-duplicated)

Usage (Windows example)
  python audit_datasets.py --root "F:\\Projects\\ShaktiKram_ML\\datasets"
  python audit_datasets.py --root "F:\\Projects\\ShaktiKram_ML\\datasets" --only WEEE
  python audit_datasets.py --root "F:\\Projects\\ShaktiKram_ML\\datasets" --dump-docs WEEE

Requires: pandas, openpyxl (xlsx). xlrd only if you have legacy .xls files.
"""
import argparse
import codecs
import hashlib
import json
import os
import re
import sys
import time
from collections import Counter, defaultdict
from pathlib import Path

import pandas as pd

# --------------------------------------------------------------------------- config
MODEL_FOLDERS = [
    "progress_tracking",
    "recovery_performance",
    "muscle_development",
    "body_composition",
]

# Expected dataset folders (edit if you rename/move datasets).
EXPECTED = {
    "progress_tracking": ["OpenPowerlifting", "Greig_2023", "Jukic_2023", "3-7RM_Strength_Test"],
    "recovery_performance": ["PMData", "REST", "SoccerMon"],
    "muscle_development": [
        "Northampton_12_Week", "Northampton_Detraining",
        "Mendeley_Resistance_Study", "Exercise_Muscle_Volume",
    ],
    "body_composition": [
        "NHANES_2011_2012", "NHANES_2013_2014", "NHANES_2017_2018",
        "Colombian_Body_Composition", "Rancho_Bernardo",
        "University_Club_Sports", "WEEE",
    ],
}

IGNORED_NAMES = {".DS_Store", "Thumbs.db", "desktop.ini"}
IGNORED_DIRS = {"__MACOSX"}

TABLE_EXT = {".csv", ".tsv", ".xlsx", ".xlsm", ".xls", ".json", ".xpt"}
DOC_EXT = {".txt", ".md", ".pdf", ".docx", ".doc", ".rtf"}
DOC_STEMS = ("readme", "license", "licence", "citation", "notice")
MEDIA_EXT = {".mp4", ".mov", ".avi", ".mkv", ".jpg", ".jpeg", ".png", ".gif", ".webp", ".mp3", ".wav"}
ARCHIVE_EXT = {".zip", ".rar", ".7z", ".gz", ".tar", ".tgz", ".bz2"}
CODE_EXT = {".r", ".py", ".ipynb", ".rmd", ".do", ".m", ".sql"}
RAW_BINARY_EXT = {".dat", ".hed", ".edf", ".bdf", ".bin"}
OTHER_TABLE_EXT = {".sav", ".dta", ".sas7bdat", ".parquet", ".feather", ".h5", ".hdf5", ".mat", ".rdata", ".rds"}

# Problems already reviewed and judged harmless.
KNOWN_BENIGN = [
    dict(dataset=r"^WEEE$", file=r"^tags\.csv$", error="EmptyDataError",
         reason="E4 tags file is empty (no event-marker button presses recorded)"),
    dict(dataset=r"^WEEE$", file=r"Event_Data_\d+\.csv$", error="ParserError",
         reason="Zephyr event log has rows with more fields than the header; "
                "re-read tolerantly by skipping malformed lines"),
]

DATE_NAME_RE = re.compile(r"(date|time|timestamp)", re.I)
TIME_ONLY_RE = re.compile(r"^\s*\d{1,2}:\d{2}(:\d{2}(\.\d+)?)?\s*$")
EARLIEST_OK = pd.Timestamp("1900-01-01")
EPOCH_START = pd.Timestamp("1970-01-01")
EPOCH_END = pd.Timestamp("1970-01-02")
LATEST_OK = pd.Timestamp.now() + pd.Timedelta(days=365)
CHUNK = 250_000
MAX_COLNAMES_CHARS = 5000


# --------------------------------------------------------------------------- helpers
def is_ignored(path: Path) -> str:
    """Return a reason string if the file should be ignored, else ''."""
    if path.name.startswith("._"):
        return "macOS AppleDouble (._*) metadata stub"
    if path.name in IGNORED_NAMES:
        return "OS metadata file"
    if any(p in IGNORED_DIRS for p in path.parts):
        return "inside __MACOSX"
    return ""


def classify_file(path: Path) -> str:
    ext = path.suffix.lower()
    stem = path.stem.lower()
    if ext in TABLE_EXT:
        return "table"
    if ext in DOC_EXT or stem.startswith(DOC_STEMS) or (ext == "" and stem.startswith(DOC_STEMS)):
        return "doc"
    if ext in MEDIA_EXT:
        return "media"
    if ext in CODE_EXT:
        return "code"
    if ext in RAW_BINARY_EXT:
        return "raw_binary_not_read"
    if ext in ARCHIVE_EXT:
        return "archive_not_extracted"
    if ext in OTHER_TABLE_EXT:
        return "unsupported_table_format"
    return "other"


def detect_encoding(path: Path) -> str:
    with open(path, "rb") as f:
        data = f.read(65536)
    for enc in ("utf-8-sig", "cp1252"):
        try:
            codecs.getincrementaldecoder(enc)(errors="strict").decode(data, final=False)
            return enc
        except UnicodeDecodeError:
            continue
    return "latin-1"


def sniff_sep(path: Path, enc: str) -> str:
    try:
        with open(path, "r", encoding=enc, errors="replace") as f:
            first = f.readline()
    except OSError:
        return ","
    counts = {",": first.count(","), ";": first.count(";"), "\t": first.count("\t"), "|": first.count("|")}
    best = max(counts, key=counts.get)
    return best if counts[best] > 0 else ","


def benign_match(dataset: str, file_name: str, error_type: str):
    for rule in KNOWN_BENIGN:
        if (re.search(rule["dataset"], dataset) and re.search(rule["file"], file_name)
                and rule["error"] == error_type):
            return rule["reason"]
    return None


# --------------------------------------------------------------------------- table stats
class TableStats:
    """Accumulates row/column statistics over one or more DataFrame chunks."""

    def __init__(self):
        self.rows = 0
        self.names = None
        self.missing = []
        self.dtype = []
        self.types = []
        self.d_all_min, self.d_all_max = [], []
        self.d_ok_min, self.d_ok_max = [], []
        self.d_susp = []
        self.d_method = []
        self.d_is_date = []

    def _init(self, df):
        n = df.shape[1]
        self.names = [str(c) for c in df.columns]
        self.missing = [0] * n
        self.dtype = [str(df.dtypes.iloc[i]) for i in range(n)]
        self.types = [
            ",".join(sorted({type(v).__name__ for v in df.iloc[:200, i].dropna()})) for i in range(n)
        ]
        self.d_all_min, self.d_all_max = [None] * n, [None] * n
        self.d_ok_min, self.d_ok_max = [None] * n, [None] * n
        self.d_susp = [0] * n
        self.d_method = [""] * n
        self.d_is_date = [False] * n

    def update(self, df: pd.DataFrame):
        if self.names is None:
            self._init(df)
        self.rows += len(df)
        for i in range(df.shape[1]):
            s = df.iloc[:, i]
            self.missing[i] += int(s.isna().sum())
            self._dates(i, s)

    def _dates(self, i, s):
        name = self.names[i]
        is_dt = pd.api.types.is_datetime64_any_dtype(s)
        is_text = pd.api.types.is_object_dtype(s) or pd.api.types.is_string_dtype(s)
        if not (is_dt or (is_text and DATE_NAME_RE.search(name))):
            return
        sample = s.dropna()
        if sample.empty:
            return
        try:
            if is_dt:
                parsed = sample
                if getattr(parsed.dt, "tz", None) is not None:
                    parsed = parsed.dt.tz_convert("UTC").dt.tz_localize(None)
                self.d_method[i] = "datetime dtype"
            else:
                if TIME_ONLY_RE.match(str(sample.iloc[0])):
                    self.d_method[i] = "time-of-day only (skipped)"
                    return
                try:
                    parsed = pd.to_datetime(sample, errors="coerce", format="mixed", utc=True)
                except TypeError:
                    parsed = pd.to_datetime(sample, errors="coerce", utc=True)
                if parsed.notna().mean() < 0.8:
                    return
                parsed = parsed.dropna().dt.tz_localize(None)
                self.d_method[i] = "date-like column name"
            if parsed.empty:
                return
            self.d_is_date[i] = True
            lo, hi = parsed.min(), parsed.max()
            self.d_all_min[i] = lo if self.d_all_min[i] is None else min(self.d_all_min[i], lo)
            self.d_all_max[i] = hi if self.d_all_max[i] is None else max(self.d_all_max[i], hi)
            bad = (parsed < EARLIEST_OK) | (parsed > LATEST_OK) | ((parsed >= EPOCH_START) & (parsed < EPOCH_END))
            self.d_susp[i] += int(bad.sum())
            good = parsed[~bad]
            if not good.empty:
                glo, ghi = good.min(), good.max()
                self.d_ok_min[i] = glo if self.d_ok_min[i] is None else min(self.d_ok_min[i], glo)
                self.d_ok_max[i] = ghi if self.d_ok_max[i] is None else max(self.d_ok_max[i], ghi)
        except Exception:
            return

    # flags for the registry "notes" column
    def quality_flags(self):
        flags = []
        if not self.names:
            return flags
        n = len(self.names)
        unnamed = sum(1 for c in self.names if c.startswith("Unnamed:"))
        if n >= 3 and unnamed / n >= 0.3:
            flags.append(f"MULTI_ROW_HEADER_SUSPECTED ({unnamed}/{n} 'Unnamed' columns)")
        rep = sum(1 for c in self.names if re.search(r"\.\d+$", c))
        if rep >= 5:
            flags.append(f"WIDE_REPEATED_COLUMNS ({rep} columns with .N suffix)")
        def _isnum(x):
            try:
                float(str(x).strip())
                return True
            except ValueError:
                return False
        numeric = sum(1 for c in self.names if _isnum(c))
        if n >= 2 and numeric / n >= 0.5:
            flags.append(f"NUMERIC_HEADER_SUSPECTED ({numeric}/{n} column names are numbers; file likely has no header row)")
        susp = sum(self.d_susp)
        if susp:
            flags.append(f"SUSPICIOUS_DATES ({susp} value(s): epoch-zero 1970-01-01, before 1900, or >1y in future)")
        return flags


# --------------------------------------------------------------------------- readers
def audit_csv(path: Path):
    """Returns (list of (table_name, TableStats), notes list, status)."""
    enc = detect_encoding(path)
    sep = sniff_sep(path, enc)
    notes, status = [], "success"
    stats = TableStats()
    try:
        for chunk in pd.read_csv(path, sep=sep, encoding=enc, chunksize=CHUNK, low_memory=False):
            stats.update(chunk)
        if stats.names is None:  # header only, no rows
            hdr = pd.read_csv(path, sep=sep, encoding=enc, nrows=0)
            stats.update(hdr)
    except pd.errors.ParserError:
        stats = TableStats()
        skipped = [0]

        def bad(_line):
            skipped[0] += 1
            return None

        for chunk in pd.read_csv(path, sep=sep, encoding=enc, engine="python",
                                 on_bad_lines=bad, chunksize=CHUNK):
            stats.update(chunk)
        status = "recovered_skipped_bad_lines"
        notes.append(f"{skipped[0]} malformed line(s) skipped")
    if enc != "utf-8-sig":
        notes.append(f"encoding={enc}")
    if sep != ",":
        notes.append(f"separator={sep!r}")
    return [("CSV", stats)], notes, status


def audit_excel(path: Path):
    out = []
    with open(path, "rb") as f:
        head = f.read(512)
    known = head.startswith(b"PK\x03\x04") or head.startswith(b"\xd0\xcf\x11\xe0")
    if not known:
        return audit_misnamed_excel(path, head)
    with pd.ExcelFile(path) as xl:
        for sheet in xl.sheet_names:
            st = TableStats()
            st.update(xl.parse(sheet))
            out.append((sheet, st))
    return out, [], "success"


def audit_misnamed_excel(path: Path, head: bytes):
    """File has an Excel extension but is not a real xls/xlsx (HTML export, XML sheet, or text)."""
    stripped = head.lstrip(b"\xef\xbb\xbf \t\r\n").lower()
    if stripped.startswith(b"<"):
        kind = "SpreadsheetML/XML" if b"urn:schemas-microsoft-com:office:spreadsheet" in head.lower() else "HTML"
        if kind == "HTML":
            tables = pd.read_html(path)
            out = []
            for i, df in enumerate(tables, 1):
                st = TableStats()
                st.update(df)
                out.append((f"HTML_table_{i}", st))
            return out, [f"MISNAMED_EXCEL: file is actually HTML ({len(tables)} table(s))"], "success"
        raise ValueError("MISNAMED_EXCEL: SpreadsheetML XML - open in Excel and re-save as .xlsx")
    enc = detect_encoding(path)
    sep = sniff_sep(path, enc)
    st = TableStats()
    for chunk in pd.read_csv(path, sep=sep, encoding=enc, chunksize=CHUNK, low_memory=False):
        st.update(chunk)
    return [("TEXT", st)], [f"MISNAMED_EXCEL: file is actually delimited text (sep={sep!r}, encoding={enc})"], "success"


def audit_json(path: Path):
    obj = None
    for enc in ("utf-8-sig", "latin-1"):
        try:
            with open(path, "r", encoding=enc) as f:
                obj = json.load(f)
            break
        except UnicodeDecodeError:
            continue
    name = "JSON"
    if isinstance(obj, list) and obj and all(isinstance(x, dict) for x in obj):
        df = pd.json_normalize(obj)
    elif isinstance(obj, dict):
        lists = [k for k, v in obj.items() if isinstance(v, list) and v and isinstance(v[0], dict)]
        if len(lists) == 1:
            df, name = pd.json_normalize(obj[lists[0]]), f"JSON[{lists[0]}]"
        else:
            df, name = pd.json_normalize(obj), "JSON_object"
    else:
        df = pd.DataFrame({"value": [obj]})
    st = TableStats()
    st.update(df)
    return [(name, st)], [], "success"


def audit_xpt(path: Path):
    df = pd.read_sas(path, format="xport", encoding="latin-1")
    st = TableStats()
    st.update(df)
    return [("XPT", st)], [], "success"


def audit_table(path: Path):
    ext = path.suffix.lower()
    if ext in (".csv", ".tsv"):
        return audit_csv(path)
    if ext in (".xlsx", ".xlsm", ".xls"):
        return audit_excel(path)
    if ext == ".json":
        return audit_json(path)
    if ext == ".xpt":
        return audit_xpt(path)
    raise ValueError(f"unsupported extension {ext}")


# --------------------------------------------------------------------------- output rows
def fmt_ts(t):
    return "" if t is None else pd.Timestamp(t).isoformat()


def colnames_str(names):
    s = " | ".join(names)
    if len(s) > MAX_COLNAMES_CHARS:
        s = s[:MAX_COLNAMES_CHARS] + f" ...(+{len(names)} columns total)"
    return s


def write_csv(df, path: Path):
    df.to_csv(path, index=False, encoding="utf-8-sig")


# --------------------------------------------------------------------------- main audit
def run_audit(root: Path, out: Path, only: str | None):
    out.mkdir(parents=True, exist_ok=True)
    suffix = f"_only-{only}" if only else ""
    registry, columns, docs, errors, benign, ignored, inventory = [], [], [], [], [], [], []
    t0 = time.time()
    n_files = 0

    for model in MODEL_FOLDERS:
        raw = root / model / "raw"
        if not raw.is_dir():
            print(f"[warn] missing folder: {raw}")
            continue
        for path in sorted(raw.rglob("*")):
            if path.is_dir():
                continue
            rel = path.relative_to(root)
            parts = rel.parts
            dataset = parts[2] if len(parts) >= 4 else "(raw root)"
            if only and dataset.lower() != only.lower():
                continue
            base = dict(model_folder=model, dataset=dataset, file_name=path.name,
                        relative_file_path=str(rel))
            reason = is_ignored(path)
            if reason:
                ignored.append({**base, "reason": reason})
                continue
            kind = classify_file(path)
            size = path.stat().st_size
            inventory.append({**base, "extension": path.suffix.lower(), "size_bytes": size,
                              "handled_as": kind})
            if kind == "doc":
                docs.append({**base, "extension": path.suffix.lower(), "size_bytes": size})
                continue
            if kind != "table":
                continue

            n_files += 1
            print(f"[{n_files}] {rel}", flush=True)
            try:
                tables, notes, status = audit_table(path)
            except Exception as exc:  # noqa: BLE001
                etype = type(exc).__name__
                msg = str(exc).replace("\n", " ")[:300]
                why = benign_match(dataset, path.name, etype)
                if etype == "EmptyDataError":
                    category = "benign_known" if why else "empty_file_unlisted"
                    why = why or "file has no content"
                    registry.append({**base, "table_or_sheet": "CSV", "format": path.suffix.lstrip(".").lower(),
                                     "row_count": 0, "column_count": 0, "column_names": "",
                                     "inspection_status": "empty_file", "notes": why})
                    benign.append({**base, "category": category, "error_type": etype,
                                   "message": msg, "reason": why})
                else:
                    errors.append({**base, "table_or_sheet": "", "error_type": etype, "error_message": msg})
                    registry.append({**base, "table_or_sheet": "", "format": path.suffix.lstrip(".").lower(),
                                     "row_count": "", "column_count": "", "column_names": "",
                                     "inspection_status": "error", "notes": f"{etype}: {msg}"})
                continue

            if status == "recovered_skipped_bad_lines":
                why = benign_match(dataset, path.name, "ParserError")
                benign.append({**base, "category": "benign_known" if why else "recovered_parse_unlisted",
                               "error_type": "ParserError", "message": "; ".join(notes),
                               "reason": why or "re-read tolerantly; review before use"})

            for tname, st in tables:
                flags = st.quality_flags()
                all_notes = "; ".join(notes + flags)
                registry.append({**base, "table_or_sheet": tname,
                                 "format": path.suffix.lstrip(".").lower(),
                                 "row_count": st.rows, "column_count": len(st.names or []),
                                 "column_names": colnames_str(st.names or []),
                                 "inspection_status": status, "notes": all_notes})
                for i, cname in enumerate(st.names or []):
                    miss = st.missing[i]
                    columns.append({
                        **base, "table_or_sheet": tname, "column_name": cname,
                        "pandas_dtype": st.dtype[i], "observed_value_types": st.types[i],
                        "row_count": st.rows, "missing_count": miss,
                        "missing_percent": round(100 * miss / st.rows, 3) if st.rows else 0.0,
                        "non_missing_count": st.rows - miss,
                        "is_date_column": st.d_is_date[i],
                        "date_min": fmt_ts(st.d_all_min[i]), "date_max": fmt_ts(st.d_all_max[i]),
                        "date_min_plausible": fmt_ts(st.d_ok_min[i]),
                        "date_max_plausible": fmt_ts(st.d_ok_max[i]),
                        "suspicious_date_count": st.d_susp[i],
                        "date_detection_method": st.d_method[i],
                    })

    # ----- write outputs
    reg_cols = ["model_folder", "dataset", "table_or_sheet", "file_name", "relative_file_path",
                "format", "row_count", "column_count", "column_names", "inspection_status", "notes"]
    write_csv(pd.DataFrame(registry, columns=reg_cols), out / f"dataset_registry{suffix}.csv")
    write_csv(pd.DataFrame(columns), out / f"column_audit{suffix}.csv")
    write_csv(pd.DataFrame(docs), out / f"documentation_files{suffix}.csv")
    write_csv(pd.DataFrame(errors, columns=["model_folder", "dataset", "file_name", "relative_file_path",
                                            "table_or_sheet", "error_type", "error_message"]),
              out / f"audit_errors{suffix}.csv")
    write_csv(pd.DataFrame(benign), out / f"audit_benign_issues{suffix}.csv")
    write_csv(pd.DataFrame(ignored), out / f"ignored_files{suffix}.csv")
    inv = pd.DataFrame(inventory)
    write_csv(inv, out / f"file_inventory{suffix}.csv")

    # ----- folder check
    fc = []
    for model, names in EXPECTED.items():
        for ds in names:
            d = root / model / "raw" / ds
            if only and ds.lower() != only.lower():
                continue
            if not d.is_dir():
                fc.append(dict(model_folder=model, dataset=ds, folder_exists=False, files=0, total_mb=0,
                               table_files=0, unsupported_files=0, status="MISSING", extensions=""))
                continue
            sub = inv[(inv.model_folder == model) & (inv.dataset == ds)] if len(inv) else inv
            nfiles = len(sub)
            unsupported = int(sub.handled_as.isin(["archive_not_extracted", "unsupported_table_format",
                                                  "other", "raw_binary_not_read"]).sum()) if nfiles else 0
            tables = int((sub.handled_as == "table").sum()) if nfiles else 0
            if nfiles == 0:
                status = "EMPTY"
            elif tables == 0 and unsupported > 0:
                status = "NO_READABLE_TABLES (check unsupported files)"
            elif tables == 0:
                status = "NO_TABLES"
            else:
                status = "OK"
            fc.append(dict(model_folder=model, dataset=ds, folder_exists=True, files=nfiles,
                           total_mb=round(sub.size_bytes.sum() / 1e6, 2) if nfiles else 0,
                           table_files=tables, unsupported_files=unsupported, status=status,
                           extensions=", ".join(f"{k}:{v}" for k, v in
                                                Counter(sub.extension).most_common()) if nfiles else ""))
    fcdf = pd.DataFrame(fc)
    write_csv(fcdf, out / f"folder_check{suffix}.csv")

    # fuzzy search for folders that could be a differently-named dataset
    hits = []
    pat = re.compile(r"univ|club|sport", re.I)
    for dirpath, dirnames, _ in os.walk(root):
        dirnames[:] = [d for d in dirnames if d not in IGNORED_DIRS and d != "_audit"]
        for d in dirnames:
            if pat.search(d):
                hits.append(os.path.join(dirpath, d))
    # unexpected dataset folders
    unexpected = []
    for model in MODEL_FOLDERS:
        raw = root / model / "raw"
        if raw.is_dir():
            for d in sorted(p.name for p in raw.iterdir() if p.is_dir()):
                if d not in EXPECTED.get(model, []) and d not in IGNORED_DIRS:
                    unexpected.append(f"{model}/raw/{d}")

    # ----- summary
    reg = pd.DataFrame(registry, columns=reg_cols)
    lines = [f"ShaktiKram audit summary  ({time.strftime('%Y-%m-%d %H:%M')})",
             f"root: {root}", f"elapsed: {time.time() - t0:.0f}s", "",
             f"table files processed : {n_files}",
             f"tables/sheets recorded: {len(reg)}",
             f"ignored (AppleDouble etc.): {len(ignored)}",
             f"benign issues (known + unlisted): {len(benign)}  "
             f"-> unlisted to review: {sum(1 for b in benign if not b['category'].startswith('benign_known'))}",
             f"REAL errors           : {len(errors)}", ""]
    if len(reg):
        lines.append("status counts: " + ", ".join(f"{k}={v}" for k, v in reg.inspection_status.value_counts().items()))
        flagged = reg[reg.notes.fillna("").str.contains("MULTI_ROW_HEADER|WIDE_REPEATED|SUSPICIOUS_DATES")]
        lines += ["", "Quality flags (see 'notes' in registry):"]
        for _, r in flagged.drop_duplicates(["dataset", "file_name", "table_or_sheet"]).iterrows():
            lines.append(f"  - {r.dataset} / {r.file_name} / {r.table_or_sheet}: {r.notes}")
    lines += ["", "Folder check:"]
    for _, r in fcdf.iterrows():
        lines.append(f"  {r.model_folder}/{r.dataset}: {r.status}  (files={r.files}, {r.total_mb} MB)")
    lines += ["", "Possible matches for university/club/sport folders anywhere under root:"]
    lines += [f"  {h}" for h in hits[:50]] or ["  (none)"]
    lines += ["", "Dataset folders present on disk but not in EXPECTED:"]
    lines += [f"  {u}" for u in unexpected] or ["  (none)"]
    unsupported_all = inv[inv.handled_as.isin(["archive_not_extracted", "unsupported_table_format", "other", "raw_binary_not_read"])] \
        if len(inv) else inv
    lines += ["", f"Files NOT read (archives/unsupported/other): {len(unsupported_all)}  -> see file_inventory.csv"]
    (out / f"audit_summary{suffix}.txt").write_text("\n".join(lines), encoding="utf-8")
    print("\n" + "\n".join(lines))
    print(f"\nOutputs written to: {out}")


# --------------------------------------------------------------------------- doc dump
def dump_docs(root: Path, out: Path, dataset: str, max_chars: int = 6000):
    out.mkdir(parents=True, exist_ok=True)
    groups = defaultdict(list)
    skipped_binary = []
    for model in MODEL_FOLDERS:
        base = root / model / "raw"
        if not base.is_dir():
            continue
        for path in sorted(base.rglob("*")):
            if (path.is_dir() or is_ignored(path) or len(path.relative_to(root).parts) < 4
                    or path.relative_to(root).parts[2].lower() != dataset.lower()):
                continue
            if classify_file(path) != "doc":
                continue
            if path.suffix.lower() in (".pdf", ".docx", ".doc", ".rtf"):
                skipped_binary.append(str(path.relative_to(root)))
                continue
            text = path.read_text(encoding=detect_encoding(path), errors="replace")
            groups[hashlib.md5(text.encode("utf-8", "ignore")).hexdigest()].append((path, text))
    lines = [f"Documentation text for dataset: {dataset}", ""]
    for _, items in sorted(groups.items(), key=lambda kv: -len(kv[1])):
        first_path, text = items[0]
        lines.append("=" * 90)
        lines.append(f"{first_path.name}   (identical content in {len(items)} file(s))")
        for p, _t in items[:5]:
            lines.append(f"   {p.relative_to(root)}")
        if len(items) > 5:
            lines.append(f"   ... +{len(items) - 5} more")
        lines.append("-" * 90)
        lines.append(text[:max_chars] + ("\n...[truncated]" if len(text) > max_chars else ""))
        lines.append("")
    if skipped_binary:
        lines += ["=" * 90, "Binary docs (open manually):"] + [f"   {p}" for p in skipped_binary[:60]]
    target = out / f"doc_dump_{dataset}.txt"
    target.write_text("\n".join(lines), encoding="utf-8")
    print(f"Wrote {target}  ({len(groups)} distinct text documents)")


# --------------------------------------------------------------------------- CLI
def main():
    ap = argparse.ArgumentParser(description="ShaktiKram ML dataset audit (v2)")
    ap.add_argument("--root", required=True, help="datasets root, e.g. F:\\Projects\\ShaktiKram_ML\\datasets")
    ap.add_argument("--out", help="output folder (default: <root>/_audit)")
    ap.add_argument("--only", help="audit a single dataset folder name (e.g. WEEE)")
    ap.add_argument("--dump-docs", metavar="DATASET", help="write README/info text of a dataset and exit")
    args = ap.parse_args()
    root = Path(args.root)
    if not root.is_dir():
        sys.exit(f"Root not found: {root}")
    out = Path(args.out) if args.out else root / "_audit"
    if args.dump_docs:
        dump_docs(root, out, args.dump_docs)
        return
    run_audit(root, out, args.only)


if __name__ == "__main__":
    main()
