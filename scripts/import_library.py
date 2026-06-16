"""CLI: масовий імпорт бібліотеки аналогів з файлу (адреса → посилання)."""

from __future__ import annotations

import argparse
import json
from datetime import datetime
from pathlib import Path

from realtify.analog_library import import_library, parse_library_file
from realtify.paths import PROJECT_ROOT


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Імпорт бібліотеки аналогів (CSV/Excel: address;url;[city];[property_type];[complex_name])."
    )
    parser.add_argument("--file", type=Path, required=True, help="Файл бібліотеки аналогів.")
    parser.add_argument("--out", type=Path, default=None, help="Тимчасова папка для збору скриншотів.")
    args = parser.parse_args(argv)

    file_path = args.file if args.file.is_absolute() else PROJECT_ROOT / args.file
    entries = parse_library_file(file_path)
    print(f"Адрес у файлі: {len(entries)}")
    out_dir = args.out or (
        PROJECT_ROOT / "web_runs" / ("library_import_" + datetime.now().strftime("%Y%m%d_%H%M%S"))
    )
    report = import_library(entries, output_dir=out_dir)
    summary = {k: v for k, v in report.items() if k != "results"}
    print(json.dumps(summary, ensure_ascii=False))
    for item in report["results"]:
        print(
            "  ",
            item.get("address"),
            "->",
            item.get("status"),
            item.get("collected", ""),
            item.get("key", item.get("error", "")),
        )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
