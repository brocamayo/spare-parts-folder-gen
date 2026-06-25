# Spare Parts Folder Generator

Generates a 3-level photo folder structure for spare parts from an Excel parts list.

## Folder Structure Created
Manufacturer / Supplier / PartId - MfrPartNumber

## Requirements
- Python 3.9+
- openpyxl (`pip install openpyxl`)

## Usage
python "Spare parts folder gen.py" --input "your_parts_list.xlsx"

## Filters Applied
- Primary_Area = PC-Manual
- PM_Status = ACTIVE, DUPLICATE, or PENDING_REVIEW

## Notes
- Place your Excel file in the same folder as the script
- Run with --dry-run first to preview folders before creating them
