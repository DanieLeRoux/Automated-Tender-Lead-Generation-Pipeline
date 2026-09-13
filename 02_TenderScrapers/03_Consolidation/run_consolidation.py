"""Run central TenderScrapers consolidation against the restructured framework.

This version reads 00_Config/sources_registry.xlsx directly and writes final
consolidated workbooks to 04_FinalOutputs by default.
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path


HERE = Path(__file__).resolve().parent
DEFAULT_PROJECT_ROOT = HERE.parent if HERE.name == "03_Consolidation" else HERE


def _add_core_to_path(project_root: Path) -> None:
    core_parent = project_root / "01_Core"
    if str(core_parent) not in sys.path:
        sys.path.insert(0, str(core_parent))


def main() -> int:
    parser = argparse.ArgumentParser(description="Consolidate and centrally score TenderScrapers outputs.")
    parser.add_argument("--project-root", default=str(DEFAULT_PROJECT_ROOT), help="02_TenderScrapers project root.")
    parser.add_argument("--sources-registry", "--registry", dest="sources_registry", help="Path to 00_Config/sources_registry.xlsx.")
    parser.add_argument("--scoring-config", help="Path to 00_Config/scoring_config.json.")
    parser.add_argument("--output-root", help="Optional final output folder override. Defaults to 04_FinalOutputs.")
    parser.add_argument("--sources-config", help="Legacy sources_config.json path. Prefer --sources-registry.")
    parser.add_argument("--data-root", help="Legacy alias for --project-root; retained for compatibility.")
    args = parser.parse_args()

    project_root = Path(args.project_root or args.data_root).resolve()
    if args.data_root and not args.project_root:
        project_root = Path(args.data_root).resolve()

    _add_core_to_path(project_root)
    from core.consolidation import consolidate  # pylint: disable=import-error,import-outside-toplevel

    registry_path = Path(args.sources_registry or (project_root / "00_Config" / "sources_registry.xlsx")).resolve()
    scoring_path = Path(args.scoring_config or (project_root / "00_Config" / "scoring_config.json")).resolve()
    output_root = Path(args.output_root).resolve() if args.output_root else project_root / "04_FinalOutputs"

    if args.sources_config and not registry_path.exists():
        paths = consolidate(
            sources_config_path=args.sources_config,
            scoring_config_path=scoring_path,
            data_root=project_root,
            output_root=output_root,
        )
    else:
        paths = consolidate(
            sources_registry_path=registry_path,
            scoring_config_path=scoring_path,
            project_root=project_root,
            data_root=project_root,
            output_root=output_root,
        )

    print("Consolidation complete. Files created:")
    for name, path in paths.items():
        print(f"- {name}: {path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
