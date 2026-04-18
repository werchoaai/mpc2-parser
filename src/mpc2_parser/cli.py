"""Command-line interface for the MPC2 parser.

Usage:
    python -m mpc2_parser.cli variant1 <asc_folder> <output.xlsx> [--split-method vertex|midpoint]
    python -m mpc2_parser.cli variant2 <asc_folder> <master.xlsx> [--output <out.xlsx>]
    python -m mpc2_parser.cli variant4 <asc_folder> <output.xlsx>
    python -m mpc2_parser.cli json     <asc_file> [--split-method vertex|midpoint]
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from .core import process_measurement
from .parser import to_serializable
from .outputs import (
    write_auswertung_workbook,
    append_to_messuebersicht,
    write_combined_workbook,
)


def _gather_asc_files(folder: Path) -> list[Path]:
    return sorted([*folder.glob("*.ASC"), *folder.glob("*.asc")])


def _process_folder(folder: Path, split_method: str):
    asc_files = _gather_asc_files(folder)
    if not asc_files:
        print(f"No .ASC files found in {folder}", file=sys.stderr)
        sys.exit(1)
    print(f"Processing {len(asc_files)} ASC files from {folder}...")
    measurements = []
    for asc in asc_files:
        try:
            m = process_measurement(asc, split_method=split_method)
            measurements.append(m)
            an = m.analysis
            print(
                f"  ✓ {asc.name}  "
                f"Ja={an.ja_ma_cm2:.3f}  Jr={an.jr_ma_cm2:.3f}  "
                f"Jr/Ja={an.jr_ja:.4f}  Qr/Qa={an.qr_qa:.4f}"
            )
        except Exception as e:
            print(f"  ✗ {asc.name}: {e}", file=sys.stderr)
    return measurements


def cmd_variant1(args):
    measurements = _process_folder(Path(args.folder), args.split_method)
    out = write_auswertung_workbook(
        measurements, Path(args.output),
        project_name=Path(args.folder).name,
    )
    print(f"\n→ Wrote Auswertung workbook: {out}")


def cmd_variant2(args):
    measurements = _process_folder(Path(args.folder), args.split_method)
    out = append_to_messuebersicht(
        measurements,
        master_xlsx=Path(args.master),
        output_xlsx=Path(args.output) if args.output else None,
        overwrite_existing_ids=args.overwrite,
    )
    print(f"\n→ Wrote Messübersicht: {out}")


def cmd_variant4(args):
    measurements = _process_folder(Path(args.folder), args.split_method)
    out = write_combined_workbook(
        measurements, Path(args.output),
        project_name=Path(args.folder).name,
    )
    print(f"\n→ Wrote combined workbook: {out}")


def cmd_json(args):
    m = process_measurement(
        Path(args.asc),
        split_method=args.split_method,
        split_override=args.split_override,
    )
    data = m.to_json_dict() if args.full else to_serializable(m.to_summary_dict())
    print(json.dumps(data, indent=2, default=str, ensure_ascii=False))


def main():
    parser = argparse.ArgumentParser(prog="mpc2_parser")
    sub = parser.add_subparsers(dest="cmd", required=True)

    def add_common(p):
        p.add_argument(
            "--split-method", choices=["vertex", "midpoint"], default="vertex",
            help="Split detection: vertex (physics-based) or midpoint (heuristic)",
        )

    p1 = sub.add_parser("variant1", help="Per-project Auswertung.xlsx (drop-in for Manuel)")
    p1.add_argument("folder", help="Folder of .ASC files")
    p1.add_argument("output", help="Output .xlsx path")
    add_common(p1)
    p1.set_defaults(func=cmd_variant1)

    p2 = sub.add_parser("variant2", help="Append rows to Messübersicht master")
    p2.add_argument("folder", help="Folder of .ASC files")
    p2.add_argument("master", help="Path to Messübersicht_*.xlsx")
    p2.add_argument("--output", help="Output .xlsx (default: <master>_updated.xlsx)")
    p2.add_argument("--overwrite", action="store_true", help="Overwrite existing rows by Messung ID")
    add_common(p2)
    p2.set_defaults(func=cmd_variant2)

    p4 = sub.add_parser("variant4", help="Combined workbook: summary + per-measurement sheets")
    p4.add_argument("folder", help="Folder of .ASC files")
    p4.add_argument("output", help="Output .xlsx path")
    add_common(p4)
    p4.set_defaults(func=cmd_variant4)

    pj = sub.add_parser("json", help="Parse a single ASC and emit JSON")
    pj.add_argument("asc", help="Path to .ASC file")
    pj.add_argument("--full", action="store_true", help="Include full curve data")
    pj.add_argument("--split-override", type=int, help="Manual split index")
    add_common(pj)
    pj.set_defaults(func=cmd_json)

    args = parser.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
