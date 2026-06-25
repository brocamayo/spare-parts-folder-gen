"""
Spare Parts Photo Folder Generator
====================================
Reads your Excel/CSV parts list and generates a 3-level folder structure
ready for photo drop-in.

Folder structure created:
  output_root/
    {Manufacturer_Name}/
      {Supplier}/
        {Part_Id} - {Manufacturer_Part_Number}/
          (drop photos here)

Filters applied:
  - Primary_Area  = PC-Manual
  - PM_Status     = ACTIVE, DUPLICATE, or PENDING_REVIEW

Also generates a README.txt in each part folder with key part info.

Usage (defaults pre-set for Alan's machine):
  python spare_parts_folder_gen.py
  python spare_parts_folder_gen.py --input parts.xlsx --output "C:\\Some\\Other\\Path"
  python spare_parts_folder_gen.py --dry-run
"""

import os
import re
import csv
import argparse
import sys

try:
    import openpyxl
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False


# ── CONFIG ──────────────────────────────────────────────────────────────────
DEFAULT_INPUT  = r"C:\Users\<your-username>\Documents\Spare Photos\parts.xlsx"
DEFAULT_OUTPUT = r"C:\Users\<your-username>\Documents\Spare Photos"

REQUIRED_COLUMNS = {
    "Part_Id":                  ["part_id", "partid", "part id"],
    "Manufacturer_Name":        ["manufacturer_name", "manufacturer name", "manufacturer"],
    "Manufacturer_Part_Number": ["manufacturer_part_number", "mfr part number", "mfr_part_number", "part number"],
    "Part_Description":         ["part_description", "description", "part description"],
    "Supplier":                 ["supplier"],
    "PM_Status":                ["pm_status", "status"],
    "Primary_Area":             ["primary_area", "primary area", "area"],
}

# ── FILTERS (edit these if your criteria change) ─────────────────────────────
FILTER_PRIMARY_AREA = "PC-Manual"                              # exact match, case-insensitive
ALLOWED_STATUSES    = {"ACTIVE", "DUPLICATE", "PENDING_REVIEW"}  # include these rows

MAX_FOLDER_NAME_LEN = 60  # keep paths short — Windows MAX_PATH is 260 chars
# ────────────────────────────────────────────────────────────────────────────


def sanitize(name: str, max_len: int = MAX_FOLDER_NAME_LEN) -> str:
    """Remove filesystem-illegal characters and truncate."""
    name = str(name).strip()
    # Replace illegal chars with underscore
    name = re.sub(r'[\\/:*?"<>|]', '_', name)
    # Collapse multiple spaces/underscores
    name = re.sub(r'[\s_]+', ' ', name).strip()
    return name[:max_len]


def normalize_header(h: str) -> str:
    return h.strip().lower().replace(" ", "_").replace("-", "_")


def map_columns(headers: list) -> dict:
    """Map canonical field names → actual column indices."""
    norm_headers = [normalize_header(h) for h in headers]
    mapping = {}
    for canonical, aliases in REQUIRED_COLUMNS.items():
        for i, h in enumerate(norm_headers):
            if h == normalize_header(canonical) or h in aliases:
                mapping[canonical] = i
                break
    return mapping


def read_excel(path: str) -> tuple[list, list]:
    if not OPENPYXL_AVAILABLE:
        raise ImportError("openpyxl not installed. Run: pip install openpyxl")
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    rows = [[str(cell.value) if cell.value is not None else "" for cell in row]
            for row in ws.iter_rows()]
    wb.close()
    if not rows:
        return [], []
    return rows[0], rows[1:]


def read_csv(path: str) -> tuple[list, list]:
    with open(path, newline='', encoding='utf-8-sig') as f:
        reader = csv.reader(f)
        rows = list(reader)
    if not rows:
        return [], []
    return rows[0], rows[1:]


def load_parts(path: str) -> tuple[list, list]:
    ext = os.path.splitext(path)[1].lower()
    if ext in (".xlsx", ".xlsm", ".xls"):
        return read_excel(path)
    elif ext == ".csv":
        return read_csv(path)
    else:
        raise ValueError(f"Unsupported file type: {ext}. Use .xlsx or .csv")


def get_cell(row, mapping, field, default="UNKNOWN"):
    idx = mapping.get(field)
    if idx is None or idx >= len(row):
        return default
    val = row[idx].strip()
    return val if val and val.lower() not in ("none", "null", "") else default


def create_readme(folder_path: str, part_info: dict):
    """Drop a README.txt inside each part folder with key details."""
    readme = os.path.join(folder_path, "README.txt")
    lines = [
        "SPARE PART PHOTO FOLDER",
        "=" * 40,
        f"Part ID:              {part_info['Part_Id']}",
        f"Description:          {part_info['Part_Description']}",
        f"Manufacturer:         {part_info['Manufacturer_Name']}",
        f"Mfr Part Number:      {part_info['Manufacturer_Part_Number']}",
        f"Supplier:             {part_info['Supplier']}",
        f"Status:               {part_info['PM_Status']}",
        f"Primary Area:         {part_info['Primary_Area']}",
        "",
        "DROP PHOTOS IN THIS FOLDER.",
        f"Recommended naming: {part_info['Part_Id']}_{part_info['Manufacturer_Part_Number']}.jpg",
        "",
        "Photo sources to try:",
        "  1. Manufacturer website product page",
        "  2. Supplier catalog / vendor email",
        "  3. Field photo (label the part before shooting)",
    ]
    with open(readme, "w", encoding="utf-8") as f:
        f.write("\n".join(lines))


def generate_folders(input_path: str, output_root: str, dry_run: bool = False):
    print(f"\n{'[DRY RUN] ' if dry_run else ''}Loading: {input_path}")
    headers, data_rows = load_parts(input_path)

    if not headers:
        print("ERROR: No headers found in file.")
        sys.exit(1)

    col_map = map_columns(headers)

    missing = [c for c in REQUIRED_COLUMNS if c not in col_map]
    if missing:
        print(f"\nWARNING: Could not find these columns (will use 'UNKNOWN'):")
        for m in missing:
            print(f"  - {m}")
        print(f"\nDetected columns: {headers}\n")

    stats = {
        "created": 0,
        "skipped_area": 0,
        "skipped_status": 0,
        "skipped_empty": 0,
        "already_exists": 0,
    }
    manufacturer_counts = {}

    for i, row in enumerate(data_rows, start=2):
        if not any(cell.strip() for cell in row):
            stats["skipped_empty"] += 1
            continue

        part_id      = get_cell(row, col_map, "Part_Id")
        mfr_name     = get_cell(row, col_map, "Manufacturer_Name")
        mfr_part_num = get_cell(row, col_map, "Manufacturer_Part_Number")
        description  = get_cell(row, col_map, "Part_Description")
        supplier     = get_cell(row, col_map, "Supplier")
        pm_status    = get_cell(row, col_map, "PM_Status", default="")
        primary_area = get_cell(row, col_map, "Primary_Area", default="")

        # ── Filter: Primary_Area ─────────────────────────────────────────────
        if primary_area.strip().lower() != FILTER_PRIMARY_AREA.lower():
            stats["skipped_area"] += 1
            continue

        # ── Filter: PM_Status ────────────────────────────────────────────────
        if pm_status.strip().upper() not in ALLOWED_STATUSES:
            stats["skipped_status"] += 1
            continue

        # ── Sanitize supplier (may have comma-separated multiple suppliers) ──
        # Use first supplier only to keep folder names clean
        supplier_clean = supplier.split(",")[0].strip()

        # ── Build folder names ───────────────────────────────────────────────
        mfr_folder      = sanitize(mfr_name)
        supplier_folder = sanitize(supplier_clean)
        part_folder     = sanitize(f"{part_id} - {mfr_part_num}")

        # 3-level path: Manufacturer / Supplier / PartId - MfrPartNum
        full_path = os.path.join(output_root, mfr_folder, supplier_folder, part_folder)

        if os.path.exists(full_path):
            stats["already_exists"] += 1
            continue

        if not dry_run:
            os.makedirs(full_path, exist_ok=True)
            create_readme(full_path, {
                "Part_Id": part_id,
                "Part_Description": description,
                "Manufacturer_Name": mfr_name,
                "Manufacturer_Part_Number": mfr_part_num,
                "Supplier": supplier,
                "PM_Status": pm_status,
                "Primary_Area": primary_area,
            })

        stats["created"] += 1
        manufacturer_counts[mfr_folder] = manufacturer_counts.get(mfr_folder, 0) + 1

        if dry_run:
            print(f"  [WOULD CREATE] {mfr_folder} / {supplier_folder} / {part_folder}")

    # ── Summary ──────────────────────────────────────────────────────────────
    print("\n" + "=" * 55)
    print("FOLDER GENERATION COMPLETE" if not dry_run else "DRY RUN COMPLETE (no folders created)")
    print("=" * 55)
    print(f"  Folders created:           {stats['created']}")
    print(f"  Already existed (skipped): {stats['already_exists']}")
    print(f"  Skipped (wrong area):      {stats['skipped_area']}")
    print(f"  Skipped (wrong status):    {stats['skipped_status']}")
    print(f"  Skipped (empty row):       {stats['skipped_empty']}")
    print(f"\n  Manufacturers found:       {len(manufacturer_counts)}")
    print(f"  Output root:               {os.path.abspath(output_root)}")

    if manufacturer_counts:
        print("\n  Parts per Manufacturer:")
        for mfr, count in sorted(manufacturer_counts.items(), key=lambda x: -x[1]):
            print(f"    {mfr:<50} {count:>4} parts")


def main():
    parser = argparse.ArgumentParser(
        description="Generate manufacturer > supplier > part photo folders for spare parts."
    )
    parser.add_argument("--input",  "-i", default=DEFAULT_INPUT,
                        help=f"Path to Excel (.xlsx) or CSV file (default: {DEFAULT_INPUT})")
    parser.add_argument("--output", "-o", default=DEFAULT_OUTPUT,
                        help=f"Root folder to create structure in (default: {DEFAULT_OUTPUT})")
    parser.add_argument("--dry-run", action="store_true",
                        help="Preview what would be created without making any folders")
    args = parser.parse_args()

    if not os.path.exists(args.input):
        print(f"\nERROR: Input file not found:\n  {args.input}")
        print("\nEither:")
        print("  1. Place your Excel file at the path above, OR")
        print("  2. Run with --input flag:  python spare_parts_folder_gen.py --input \"C:\\path\\to\\your\\file.xlsx\"")
        sys.exit(1)

    generate_folders(args.input, args.output, dry_run=args.dry_run)


if __name__ == "__main__":
    main()