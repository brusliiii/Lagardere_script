#!/usr/bin/env python3
"""Merge CSV files and optionally preview the output."""

import argparse
import csv
import glob
import sys
from typing import List


def load_files(input_glob: str) -> List[str]:
    files = sorted(glob.glob(input_glob))
    if not files:
        raise FileNotFoundError(f"No files matched: {input_glob}")
    return files


def merge_csvs(input_glob: str, output_path: str, preview_rows: int) -> None:
    files = load_files(input_glob)

    headers = None
    total_rows = 0

    with open(output_path, "w", newline="", encoding="utf-8") as out_f:
        writer = None

        for path in files:
            with open(path, "r", newline="", encoding="utf-8") as in_f:
                reader = csv.reader(in_f)
                file_headers = next(reader, None)
                if file_headers is None:
                    continue

                if headers is None:
                    headers = file_headers
                    writer = csv.writer(out_f)
                    writer.writerow(headers)
                elif file_headers != headers:
                    raise ValueError(f"Header mismatch in {path}")

                for row in reader:
                    writer.writerow(row)
                    total_rows += 1

    print(f"Merged {len(files)} files, {total_rows} rows -> {output_path}")

    if preview_rows > 0:
        print("\nPreview:")
        with open(output_path, "r", newline="", encoding="utf-8") as f:
            reader = csv.reader(f)
            for i, row in enumerate(reader):
                if i >= preview_rows + 1:  # include header
                    break
                print(row)


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Merge CSV files (same headers) and preview the output."
    )
    parser.add_argument("input_glob", help="Input glob, e.g. './data/*.csv'")
    parser.add_argument("output", help="Output CSV file")
    parser.add_argument(
        "--preview",
        type=int,
        default=5,
        help="Number of data rows to preview (0 to disable)",
    )

    args = parser.parse_args()

    try:
        merge_csvs(args.input_glob, args.output, args.preview)
    except (FileNotFoundError, ValueError) as exc:
        print(str(exc), file=sys.stderr)
        return 1

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
