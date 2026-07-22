from __future__ import annotations

import argparse
import os
import subprocess
import sys
from pathlib import Path

import build_tmall_report
import build_pdd_report
from handoff_protection import find_header_row, is_excluded_order_path


SCRIPT_DIR = Path(__file__).resolve().parent
PLATFORM_SCRIPTS = {
    "tmall": SCRIPT_DIR / "build_tmall_report.py",
    "pdd": SCRIPT_DIR / "build_pdd_report.py",
}

DEFAULT_DATABASE_ROOT = Path(os.environ.get("USERPROFILE", str(Path.home()))) / "Desktop" / "运营数据库"
ORDER_FILE_EXTENSIONS = {".xls", ".xlsx", ".csv"}


def option_value(arguments: list[str], option: str, default: str) -> str:
    for index, value in enumerate(arguments):
        if value == option and index + 1 < len(arguments):
            return arguments[index + 1]
        if value.startswith(option + "="):
            return value.split("=", 1)[1]
    return default


def detect_platforms(database_root: Path) -> set[str]:
    orders_dir = database_root / "订单数据"
    if not orders_dir.exists():
        return set()

    detected: set[str] = set()
    for path in orders_dir.rglob("*"):
        if (
            not path.is_file()
            or path.suffix.lower() not in ORDER_FILE_EXTENSIONS
            or path.name.startswith("~$")
            or is_excluded_order_path(path)
        ):
            continue
        try:
            sheets = build_tmall_report.read_workbook_sheets(
                path, max_rows=20
            )
        except Exception:
            continue
        for sheet in sheets:
            if find_header_row(sheet.rows, build_tmall_report.is_tmall_order_headers) is not None:
                detected.add("tmall")
            if find_header_row(sheet.rows, build_pdd_report.is_pdd_order_headers) is not None:
                detected.add("pdd")
    return detected


def run_platform(platform: str, passthrough: list[str]) -> int:
    command = [sys.executable, str(PLATFORM_SCRIPTS[platform]), *passthrough]
    print(f"开始生成{platform}报表")
    return subprocess.run(command, check=False).returncode


def main() -> int:
    parser = argparse.ArgumentParser(description="Generate a Tmall or Pinduoduo operating report.")
    parser.add_argument("--platform", choices=[*sorted(PLATFORM_SCRIPTS), "both"])
    args, passthrough = parser.parse_known_args()
    if args.platform == "both":
        selected = ["tmall", "pdd"]
    elif args.platform:
        selected = [args.platform]
    else:
        database_root = Path(option_value(passthrough, "--database-root", str(DEFAULT_DATABASE_ROOT)))
        selected = [platform for platform in ("tmall", "pdd") if platform in detect_platforms(database_root)]
        if not selected:
            print("未识别到天猫或拼多多订单表，请检查订单数据目录和表头。", file=sys.stderr)
            return 2

    results = [run_platform(platform, passthrough) for platform in selected]
    return max(results, default=0)


if __name__ == "__main__":
    raise SystemExit(main())
