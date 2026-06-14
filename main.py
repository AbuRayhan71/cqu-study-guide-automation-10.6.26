from __future__ import annotations

import argparse
from pathlib import Path

from app.docx_pipeline import GenerationMetadata, generate_study_guide


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Generate a CQU-branded study guide DOCX.")
    parser.add_argument("input_docx", type=Path)
    parser.add_argument("template_docx", type=Path)
    parser.add_argument("output_docx", type=Path)
    parser.add_argument("--unit-code", default="")
    parser.add_argument("--unit-name", default="")
    parser.add_argument("--week", dest="week_number", default="")
    parser.add_argument("--week-title", default="")
    parser.add_argument("--version", default="")
    parser.add_argument("--year", default="")
    parser.add_argument("--term", default="")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    metadata = GenerationMetadata(
        unit_code=args.unit_code,
        unit_name=args.unit_name,
        week_number=args.week_number,
        week_title=args.week_title,
        version=args.version,
        year=args.year,
        term=args.term,
    )
    result = generate_study_guide(args.input_docx, args.template_docx, args.output_docx, metadata)
    print(f"Generated: {result.output_path}")
    print(f"Blocks: {result.block_count}, images: {result.image_count}, tables: {result.table_count}")


if __name__ == "__main__":
    main()
