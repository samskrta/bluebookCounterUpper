#!/usr/bin/env python3
"""
Blue Book Counter Upper
-----------------------

Small utility to parse the first sheet (typically named 'ALL') of an Excel
export and count the number of unique quotes created by each technician.

Assumptions based on the report structure:
- Each quote begins on a row where column B contains the creator string in the
  form "TECH NAME (HH:MM AM|PM)".
- Rows belonging to the same quote leave column B empty.
- Column B's header can be something like "Created %sBy (At)"; this is ignored.

Usage:
  python bluebook_count.py --file data/BlueBook_Report_10_01_2025.xlsx

Optional:
  --sheet ALL             Sheet name (defaults to first sheet if not found)
  --csv out.csv           Write results to a CSV file in addition to stdout

This script has no external dependencies beyond openpyxl.
"""

from __future__ import annotations

import argparse
import csv
import os
import re
import subprocess
import sys
from collections import Counter, OrderedDict
from pathlib import Path

from openpyxl import load_workbook


TECH_ENTRY_REGEX = re.compile(
    r"^(?P<name>.+?)\s*\((?P<time>\d{1,2}:\d{2}\s?[AP]M)\)$",
    re.IGNORECASE,
)


def extract_technician_name(cell_value: object) -> str | None:
    """Return technician name if the given cell value marks a new quote.

    A valid creator cell matches "<name> (<time>)" where <time> is like
    5:07 PM or 05:07PM. Returns the <name> portion or None if not a match.
    """

    if cell_value is None:
        return None

    text = str(cell_value).strip()
    if not text:
        return None

    # Skip common header text
    header_like = text.lower().replace(" ", "")
    if header_like.startswith("created%sby(at)".replace(" ", "")):
        return None

    match = TECH_ENTRY_REGEX.match(text)
    if match:
        name = match.group("name").strip()
        # Normalize excessive inner whitespace
        name = re.sub(r"\s+", " ", name)
        return name
    return None


def count_quotes_per_technician(xlsx_path: Path, sheet_name: str | None) -> OrderedDict[str, int]:
    """Parse the workbook and count quotes per technician.

    If sheet_name is provided but not found, the first sheet is used.
    Returns an OrderedDict sorted by descending count then by name.
    """

    if not xlsx_path.exists():
        raise FileNotFoundError(f"File not found: {xlsx_path}")

    wb = load_workbook(filename=str(xlsx_path), data_only=True, read_only=True)
    ws = None
    if sheet_name and sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
    else:
        ws = wb.worksheets[0]

    counts: Counter[str] = Counter()

    # Iterate all rows, examine column B (index 2)
    for row in ws.iter_rows(values_only=True):
        col_b_value = row[1] if len(row) > 1 else None
        name = extract_technician_name(col_b_value)
        if name:
            counts[name] += 1

    # Sort by descending count then ascending name for stable output
    sorted_items = sorted(counts.items(), key=lambda kv: (-kv[1], kv[0].lower()))
    ordered = OrderedDict(sorted_items)
    return ordered


def print_summary(results: OrderedDict[str, int]) -> None:
    total = sum(results.values())
    print("Technician,Quotes")
    for tech, count in results.items():
        print(f"{tech},{count}")
    print(f"TOTAL,{total}")


def write_csv(results: OrderedDict[str, int], out_path: Path) -> None:
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with out_path.open("w", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["Technician", "Quotes"])
        for tech, count in results.items():
            writer.writerow([tech, count])
        writer.writerow(["TOTAL", sum(results.values())])


def open_with_default_app(path: Path) -> None:
    """Best-effort open of the given path with the OS default app."""
    try:
        if sys.platform == "darwin":
            subprocess.run(["open", str(path)], check=False)
        elif os.name == "nt":
            os.startfile(str(path))  # type: ignore[attr-defined]
        else:
            subprocess.run(["xdg-open", str(path)], check=False)
    except Exception:
        # Non-fatal if we fail to open automatically
        pass


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Count Blue Book quotes per technician from Excel export.")
    default_file = Path("data/BlueBook_Report_10_01_2025.xlsx")
    parser.add_argument("--file", type=Path, default=None, help="Path to the Excel file (.xlsx)")
    parser.add_argument("--sheet", type=str, default="ALL", help="Sheet name to parse (default: ALL)")
    parser.add_argument("--csv", type=Path, default=None, help="Optional path to write results CSV")
    parser.add_argument("--gui", action="store_true", help="Use a simple file picker GUI and message box output")

    args = parser.parse_args(argv)

    selected_file: Path | None = args.file

    # GUI fallback: if --gui is set or --file is missing, prompt for a file via Tk
    if args.gui or selected_file is None:
        try:
            # Defer Tk imports so CLI environments without GUI don't pay the cost
            import tkinter as tk
            from tkinter import filedialog, messagebox

            root = tk.Tk()
            root.withdraw()
            root.attributes("-topmost", True)
            initial_dir = str(Path.cwd() / "data")
            selected_path = filedialog.askopenfilename(
                title="Select BlueBook Excel report",
                initialdir=initial_dir,
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            )
            if not selected_path:
                messagebox.showwarning("BlueBook Counter Upper", "No file selected. Exiting.")
                return 1
            selected_file = Path(selected_path)
        except Exception as exc:
            print(f"GUI selection failed: {exc}", file=sys.stderr)
            # If GUI fails and no file was provided, abort
            if selected_file is None:
                return 1

    # Default CSV output path next to source file if not provided
    default_csv: Path | None = None
    if selected_file is not None:
        default_csv = selected_file.with_name(selected_file.stem + "_counts.csv")

    try:
        results = count_quotes_per_technician(selected_file, args.sheet)
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 1

    print_summary(results)

    out_csv_path: Path | None = args.csv or default_csv
    if out_csv_path is not None:
        try:
            write_csv(results, out_csv_path)
        except Exception as exc:
            print(f"Failed to write CSV: {exc}", file=sys.stderr)
            return 1
        # Best-effort auto-open of the generated CSV
        open_with_default_app(out_csv_path)

    # If GUI mode, show a simple completion dialog
    if args.gui:
        try:
            import tkinter as tk
            from tkinter import messagebox

            root = tk.Tk()
            root.withdraw()
            root.attributes("-topmost", True)
            total = sum(results.values())
            message = (
                f"Finished. Found {total} quotes across {len(results)} technicians.\n\n"
                f"CSV saved to:\n{out_csv_path}"
            )
            messagebox.showinfo("BlueBook Counter Upper", message)
        except Exception:
            # Non-fatal if dialog fails
            pass
    return 0


if __name__ == "__main__":
    raise SystemExit(main())



