from pathlib import Path
import argparse, csv, hashlib, json
from datetime import datetime, timezone

def sha256_file(path, chunk_size=1024*1024):
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(chunk_size), b""):
            h.update(chunk)
    return h.hexdigest()

def main():
    ap = argparse.ArgumentParser(description="Read-only recursive inventory for Model 3 datasets.")
    ap.add_argument("--root", required=True, help="Model 3 raw dataset root")
    ap.add_argument("--out", required=True, help="Inspection output directory")
    ap.add_argument("--hash", action="store_true", help="Calculate SHA-256 (can take time for large files)")
    args = ap.parse_args()
    root, out = Path(args.root).resolve(), Path(args.out).resolve()
    if not root.exists() or not root.is_dir():
        raise SystemExit(f"Dataset root does not exist or is not a directory: {root}")
    out.mkdir(parents=True, exist_ok=True)
    rows = []
    for p in sorted(root.rglob("*")):
        if not p.is_file():
            continue
        stat = p.stat()
        row = {
            "relative_path": str(p.relative_to(root)),
            "file_name": p.name,
            "extension": p.suffix.lower() or "[no_extension]",
            "size_bytes": stat.st_size,
            "size_mb": round(stat.st_size / (1024 * 1024), 4),
            "modified_utc": datetime.fromtimestamp(stat.st_mtime, timezone.utc).isoformat(),
            "sha256": sha256_file(p) if args.hash else ""
        }
        rows.append(row)
    csv_path = out / "model3_file_inventory.csv"
    with csv_path.open("w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=list(rows[0].keys()) if rows else [
            "relative_path","file_name","extension","size_bytes","size_mb","modified_utc","sha256"])
        w.writeheader()
        w.writerows(rows)
    summary = {
        "dataset_root": str(root),
        "generated_utc": datetime.now(timezone.utc).isoformat(),
        "file_count": len(rows),
        "total_size_bytes": sum(r["size_bytes"] for r in rows),
        "total_size_mb": round(sum(r["size_bytes"] for r in rows)/(1024*1024), 3),
        "extension_counts": {}
    }
    for r in rows:
        summary["extension_counts"][r["extension"]] = summary["extension_counts"].get(r["extension"], 0) + 1
    (out / "model3_inventory_summary.json").write_text(json.dumps(summary, indent=2), encoding="utf-8")
    print(f"Inventory complete: {len(rows)} files")
    print(f"CSV: {csv_path}")
    print(f"Summary: {out / 'model3_inventory_summary.json'}")

if __name__ == "__main__":
    main()
