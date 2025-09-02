#!/usr/bin/env python3
"""
Read the 'IMPORT' worksheet from the Hudson template Excel file into a pandas DataFrame.

Usage:
  python scripts/read_template.py [-i path/to/template.xlsx] [--preview-rows N]

If -i/--input is omitted, the script will attempt to auto-detect a single template
file inside the "template-files" directory located at the project root.
"""
from __future__ import annotations

import argparse
from pathlib import Path
from typing import Optional

import pandas as pd

SHEET_NAME = "IMPORT"
DEFAULT_SEARCH_DIRNAME = "template-files"


def find_default_template(start_dir: Path) -> Optional[Path]:
    """Try to locate a single template Excel file under template-files.

    Strategy:
    1) Prefer files that match *HUDSON* and *TEMPLATE* (case-insensitive) with .xlsx/.xlsm extension
    2) Otherwise, fall back to the first .xlsx/.xlsm file
    3) Return None if none found or if ambiguous
    """
    candidates = []
    if not start_dir.exists():
        return None

    exts = {".xlsx", ".xlsm"}
    for p in sorted(start_dir.iterdir()):
        if p.is_file() and p.suffix.lower() in exts:
            candidates.append(p)

    if not candidates:
        return None

    # Prefer HUDSON + TEMPLATE in name
    def score(path: Path) -> int:
        name = path.name.lower()
        s = 0
        if "hudson" in name:
            s += 1
        if "template" in name:
            s += 1
        return s

    candidates.sort(key=score, reverse=True)

    # If top two have the same score and >1 file exists, treat as ambiguous
    if len(candidates) > 1 and score(candidates[0]) == score(candidates[1]):
        # ambiguous default
        return None

    return candidates[0]


def read_import_sheet(xlsx_path: Path, sheet_name: str = SHEET_NAME) -> pd.DataFrame:
    xlsx_path = Path(xlsx_path)
    if not xlsx_path.exists():
        raise FileNotFoundError(f"Template Excel file not found: {xlsx_path}")

    df = pd.read_excel(xlsx_path, sheet_name=sheet_name, engine="openpyxl")
    return df


def main() -> None:
    parser = argparse.ArgumentParser(description="Read the IMPORT sheet from the Hudson template file into a pandas DataFrame.")
    parser.add_argument(
        "-i",
        "--input",
        help="Path to the Hudson template Excel file (.xlsx or .xlsm). If omitted, auto-detects in template-files/",
        default=None,
    )
    parser.add_argument(
        "--preview-rows",
        type=int,
        default=5,
        help="Number of rows to preview (default: 5)",
    )

    args = parser.parse_args()

    project_root = Path(__file__).resolve().parents[1]
    default_dir = project_root / DEFAULT_SEARCH_DIRNAME

    if args.input:
        template_path = Path(args.input)
    else:
        candidate = find_default_template(default_dir)
        if candidate is None:
            raise SystemExit(
                f"No unambiguous template file found in '{default_dir}'. "
                f"Please specify with -i/--input."
            )
        template_path = candidate

    df = read_import_sheet(template_path, sheet_name=SHEET_NAME)

    print(f"Read DataFrame from sheet '{SHEET_NAME}' in {template_path}")
    with pd.option_context("display.max_columns", 30, "display.width", 140):
        print(df.head(args.preview_rows))
    print(f"DataFrame shape: {df.shape}")


if __name__ == "__main__":
    main()

