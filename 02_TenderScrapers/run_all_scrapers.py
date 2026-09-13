"""Run all active TenderScrapers sources from the central source registry.

This runner is aligned to the restructured framework:

00_Config
01_Core
02_Sources
03_Consolidation
04_FinalOutputs
05_SystemLogs
06_SystemState
07_SystemMetadata
08_Hotfixes
09_Tests
99_Archive

Key behaviours:
- reads 00_Config/sources_registry.xlsx;
- skips inactive sources;
- runs each scraper from its own 00_Scripts folder;
- writes central run summaries to 05_SystemLogs/RunLogs;
- does not fail the whole run because one source is inactive or missing;
- can optionally run World Bank translation;
- attempts consolidation after the source runs, if a consolidation script is found.
"""
from __future__ import annotations

import argparse
import shlex
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path
from typing import Any

import pandas as pd


HERE = Path(__file__).resolve().parent
DEFAULT_PROJECT_ROOT = HERE
DEFAULT_REGISTRY = HERE / "00_Config" / "sources_registry.xlsx"

REGISTRY_SHEET = "Sources_Registry"

STANDARD_SOURCE_SUBFOLDERS = [
    "00_Scripts",
    "01_RawData",
    "02_Outputs",
    "03_Logs",
    "04_State",
    "05_Metadata",
]


def _now_stamp() -> str:
    return datetime.now().strftime("%Y%m%d_%H%M%S")


def _normalise_bool(value: Any) -> bool:
    if isinstance(value, bool):
        return value
    if value is None:
        return False
    try:
        if pd.isna(value):
            return False
    except TypeError:
        pass
    text = str(value).strip().casefold()
    return text in {"true", "yes", "y", "1", "active", "enabled"}


def _safe_text(value: Any) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except TypeError:
        pass
    return str(value).strip()


def _resolve(project_root: Path, path_value: Any) -> Path:
    text = _safe_text(path_value)
    if not text:
        return Path()
    path = Path(text.replace("\\", "/"))
    if path.is_absolute():
        return path
    return project_root / path


def _split_args(value: Any) -> list[str]:
    text = _safe_text(value)
    if not text:
        return []
    # Windows-friendly enough for simple default argument strings.
    return shlex.split(text, posix=False)


def load_registry(registry_path: Path) -> pd.DataFrame:
    if not registry_path.exists():
        raise FileNotFoundError(f"Source registry not found: {registry_path}")

    df = pd.read_excel(registry_path, sheet_name=REGISTRY_SHEET, engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]
    required = {"source_name", "source_key", "active", "folder_name", "script_path"}
    missing = sorted(required - set(df.columns))
    if missing:
        raise ValueError(f"Registry is missing required column(s): {', '.join(missing)}")

    df = df[df["source_name"].notna()].copy()
    df["active_bool"] = df["active"].apply(_normalise_bool)
    return df


def ensure_source_folders(project_root: Path, folder_name: str) -> None:
    source_root = project_root / "02_Sources" / folder_name
    for subfolder in STANDARD_SOURCE_SUBFOLDERS:
        (source_root / subfolder).mkdir(parents=True, exist_ok=True)


def write_summary(rows: list[dict[str, Any]], output_root: Path) -> tuple[Path, Path | None]:
    output_root.mkdir(parents=True, exist_ok=True)
    ts = _now_stamp()
    df = pd.DataFrame(rows)

    csv_path = output_root / f"run_summary_{ts}.csv"
    xlsx_path = output_root / f"run_summary_{ts}.xlsx"

    df.to_csv(csv_path, index=False)

    try:
        with pd.ExcelWriter(
            xlsx_path,
            engine="xlsxwriter",
            engine_kwargs={"options": {"strings_to_urls": False}},
            datetime_format="yyyy-mm-dd hh:mm:ss",
        ) as writer:
            df.to_excel(writer, index=False, sheet_name="Run_Summary")
            worksheet = writer.sheets["Run_Summary"]
            worksheet.freeze_panes(1, 0)
            for idx, col in enumerate(df.columns):
                width = min(max(len(str(col)) + 2, 14), 70)
                worksheet.set_column(idx, idx, width)
    except Exception as exc:  # pragma: no cover - summary XLSX is helpful but not critical
        print(f"WARNING: Could not write Excel run summary: {exc}")
        return csv_path, None

    return csv_path, xlsx_path


def _write_source_process_log(
    log_dir: Path,
    source_key: str,
    cmd: list[str],
    return_code: int,
    elapsed_seconds: float,
    stdout: str,
    stderr: str,
) -> Path:
    log_dir.mkdir(parents=True, exist_ok=True)
    safe_key = "".join(ch if ch.isalnum() or ch in {"-", "_"} else "_" for ch in source_key) or "source"
    log_path = log_dir / f"{safe_key}_{_now_stamp()}.log"

    with open(log_path, "w", encoding="utf-8") as fh:
        fh.write(f"command: {' '.join(cmd)}\n")
        fh.write(f"return_code: {return_code}\n")
        fh.write(f"elapsed_seconds: {elapsed_seconds:.2f}\n")
        fh.write("\n--- STDOUT ---\n")
        fh.write(stdout or "")
        fh.write("\n--- STDERR ---\n")
        fh.write(stderr or "")

    return log_path


def run_source(
    *,
    project_root: Path,
    source: pd.Series,
    python_executable: str,
    run_log_root: Path,
    translate_worldbank: bool,
    worldbank_auto_install_models: bool,
) -> dict[str, Any]:
    source_name = _safe_text(source.get("source_name"))
    source_key = _safe_text(source.get("source_key")) or source_name
    folder_name = _safe_text(source.get("folder_name"))
    active = bool(source.get("active_bool", False))

    row: dict[str, Any] = {
        "run_timestamp": datetime.now().isoformat(timespec="seconds"),
        "source_name": source_name,
        "source_key": source_key,
        "active": active,
        "status": "",
        "return_code": "",
        "elapsed_seconds": "",
        "script_path": "",
        "process_log_path": "",
        "message": "",
    }

    if not active:
        row.update(status="SKIPPED_INACTIVE", message="Source inactive in sources_registry.xlsx.")
        print(f"SKIP inactive: {source_name}")
        return row

    if not folder_name:
        row.update(status="INVALID_REGISTRY_ROW", message="Missing folder_name.")
        print(f"SKIP invalid registry row: {source_name} has no folder_name")
        return row

    ensure_source_folders(project_root, folder_name)

    script_path = _resolve(project_root, source.get("script_path"))
    row["script_path"] = str(script_path)

    if not script_path.exists():
        row.update(status="MISSING_SCRIPT", message=f"Script not found: {script_path}")
        print(f"SKIP missing script: {source_name} -> {script_path}")
        return row

    cmd = [python_executable, str(script_path)]

    # Optional registry extension: a default_args column can be added later without changing the runner.
    cmd.extend(_split_args(source.get("default_args")))

    if source_key.casefold() == "worldbank" and translate_worldbank:
        if "--translate-non-english" not in cmd:
            cmd.append("--translate-non-english")
        if worldbank_auto_install_models and "--argos-auto-install-models" not in cmd:
            cmd.append("--argos-auto-install-models")

    print(f"\nRUN {source_name}")
    print(" ".join(cmd))

    started = time.perf_counter()
    result = subprocess.run(
        cmd,
        cwd=str(script_path.parent),
        text=True,
        capture_output=True,
        encoding="utf-8",
        errors="replace",
    )
    elapsed = time.perf_counter() - started

    process_log = _write_source_process_log(
        run_log_root,
        source_key,
        cmd,
        result.returncode,
        elapsed,
        result.stdout,
        result.stderr,
    )

    row["return_code"] = result.returncode
    row["elapsed_seconds"] = round(elapsed, 2)
    row["process_log_path"] = str(process_log)

    # Keep the console useful without dumping very large scraper logs.
    if result.stdout:
        tail = "\n".join(result.stdout.splitlines()[-12:])
        print(tail)
    if result.stderr:
        err_tail = "\n".join(result.stderr.splitlines()[-8:])
        print(err_tail)

    if result.returncode == 0:
        row.update(status="OK", message="Completed.")
    else:
        row.update(status="FAILED", message=f"Return code {result.returncode}. See process log.")
        print(f"FAILED {source_name} with return code {result.returncode}")

    return row


def _candidate_consolidation_scripts(project_root: Path) -> list[Path]:
    return [
        project_root / "03_Consolidation" / "run_consolidation.py",
        project_root / "03_Consolidation" / "consolidate_outputs.py",
        project_root / "03_Consolidation" / "consolidation.py",
        project_root / "03_Consolidation" / "build_consolidated_outputs.py",
        project_root / "01_Core" / "consolidation.py",
    ]


def find_consolidation_script(project_root: Path) -> Path | None:
    for path in _candidate_consolidation_scripts(project_root):
        if path.exists():
            return path
    return None


def build_consolidation_command(
    *,
    python_executable: str,
    script_path: Path,
    project_root: Path,
    registry_path: Path,
) -> list[str]:
    cmd = [python_executable, str(script_path)]
    try:
        text = script_path.read_text(encoding="utf-8", errors="ignore")
    except OSError:
        text = ""

    # Add arguments only if the consolidation script appears to support them.
    if "--project-root" in text:
        cmd.extend(["--project-root", str(project_root)])
    elif "--data-root" in text:
        cmd.extend(["--data-root", str(project_root)])

    if "--sources-registry" in text:
        cmd.extend(["--sources-registry", str(registry_path)])
    elif "--registry" in text:
        cmd.extend(["--registry", str(registry_path)])

    return cmd


def run_consolidation(
    *,
    project_root: Path,
    registry_path: Path,
    python_executable: str,
    run_log_root: Path,
) -> dict[str, Any]:
    row: dict[str, Any] = {
        "run_timestamp": datetime.now().isoformat(timespec="seconds"),
        "source_name": "CONSOLIDATION",
        "source_key": "consolidation",
        "active": True,
        "status": "",
        "return_code": "",
        "elapsed_seconds": "",
        "script_path": "",
        "process_log_path": "",
        "message": "",
    }

    script = find_consolidation_script(project_root)
    if not script:
        row.update(
            status="SKIPPED_NOT_FOUND",
            message=(
                "No consolidation script found in 03_Consolidation or 01_Core. "
                "Upload/update the consolidation script next."
            ),
        )
        print("\nSKIP consolidation: no consolidation script found.")
        return row

    cmd = build_consolidation_command(
        python_executable=python_executable,
        script_path=script,
        project_root=project_root,
        registry_path=registry_path,
    )
    row["script_path"] = str(script)

    print("\nRUN consolidation")
    print(" ".join(cmd))

    started = time.perf_counter()
    result = subprocess.run(
        cmd,
        cwd=str(script.parent),
        text=True,
        capture_output=True,
        encoding="utf-8",
        errors="replace",
    )
    elapsed = time.perf_counter() - started

    process_log = _write_source_process_log(
        run_log_root,
        "consolidation",
        cmd,
        result.returncode,
        elapsed,
        result.stdout,
        result.stderr,
    )

    row["return_code"] = result.returncode
    row["elapsed_seconds"] = round(elapsed, 2)
    row["process_log_path"] = str(process_log)

    if result.stdout:
        print("\n".join(result.stdout.splitlines()[-20:]))
    if result.stderr:
        print("\n".join(result.stderr.splitlines()[-10:]))

    if result.returncode == 0:
        row.update(status="OK", message="Consolidation completed.")
    else:
        row.update(status="FAILED", message=f"Return code {result.returncode}. See process log.")
        print(f"CONSOLIDATION FAILED with return code {result.returncode}")

    return row


def main() -> int:
    parser = argparse.ArgumentParser(description="Run all active tender scrapers from sources_registry.xlsx.")
    parser.add_argument("--project-root", default=str(DEFAULT_PROJECT_ROOT), help="02_TenderScrapers project root.")
    parser.add_argument("--registry", default=str(DEFAULT_REGISTRY), help="Path to 00_Config/sources_registry.xlsx.")
    parser.add_argument("--python", default=sys.executable, help="Python executable used for child scraper runs.")
    parser.add_argument("--skip-consolidation", action="store_true")
    parser.add_argument("--stop-on-failure", action="store_true", help="Stop immediately if any active source scraper fails.")
    parser.add_argument("--translate-worldbank", action="store_true", help="Run World Bank with --translate-non-english.")
    parser.add_argument(
        "--worldbank-auto-install-models",
        action="store_true",
        help="When translating World Bank, also pass --argos-auto-install-models.",
    )
    args = parser.parse_args()

    project_root = Path(args.project_root).resolve()
    registry_path = Path(args.registry).resolve()
    run_log_root = project_root / "05_SystemLogs" / "RunLogs"
    run_log_root.mkdir(parents=True, exist_ok=True)

    print(f"Project root: {project_root}")
    print(f"Registry: {registry_path}")
    print(f"Run logs: {run_log_root}")

    registry = load_registry(registry_path)

    summary_rows: list[dict[str, Any]] = []
    failures: list[tuple[str, int | str]] = []

    for _, source in registry.iterrows():
        row = run_source(
            project_root=project_root,
            source=source,
            python_executable=args.python,
            run_log_root=run_log_root,
            translate_worldbank=args.translate_worldbank,
            worldbank_auto_install_models=args.worldbank_auto_install_models,
        )
        summary_rows.append(row)

        if row["status"] == "FAILED":
            failures.append((row["source_name"], row["return_code"]))
            if args.stop_on_failure:
                csv_path, xlsx_path = write_summary(summary_rows, run_log_root)
                print(f"\nRun summary: {csv_path}")
                if xlsx_path:
                    print(f"Run summary Excel: {xlsx_path}")
                return int(row["return_code"] or 1)

    if not args.skip_consolidation:
        consolidation_row = run_consolidation(
            project_root=project_root,
            registry_path=registry_path,
            python_executable=args.python,
            run_log_root=run_log_root,
        )
        summary_rows.append(consolidation_row)
        if consolidation_row["status"] == "FAILED":
            failures.append(("CONSOLIDATION", consolidation_row["return_code"]))

    csv_path, xlsx_path = write_summary(summary_rows, run_log_root)
    print(f"\nRun summary: {csv_path}")
    if xlsx_path:
        print(f"Run summary Excel: {xlsx_path}")

    if failures:
        print("\nSome active items failed:")
        for name, code in failures:
            print(f"- {name}: return code {code}")
        return 1

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
