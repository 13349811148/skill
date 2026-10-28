from __future__ import annotations

import argparse
import json
import os
import subprocess
import sys
from pathlib import Path
import time
from wps_upload import prepare as prepare_wps_upload
from report_archive import requested_dates

import build_tmall_report
import build_pdd_report
from portable_reports import read_workbook_sheets
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
        orders_dir = database_root / "2.订单数据库"
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
            sheets = read_workbook_sheets(
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
    if os.environ.get('DAILY_REPORT_ARCHIVE')=='1' and not option_value(passthrough,'--output',''):
        name='商品日报_拼多多.xlsx' if platform=='pdd' else '商品日报_天猫.xlsx'
        passthrough=[*passthrough,'--output',str(Path(os.environ['DAILY_REPORT_ARCHIVE_DIR'])/name)]
    command = [sys.executable, str(PLATFORM_SCRIPTS[platform]), *passthrough]
    started_ns = time.time_ns()
    print(f"开始生成{platform}报表")
    result = subprocess.run(command, check=False).returncode
    if result == 0:
        module = build_tmall_report if platform == 'tmall' else build_pdd_report
        report = Path(option_value(passthrough, '--output', str(module.DEFAULT_OUTPUT)))
        try:
            task = prepare_wps_upload(report, platform, started_ns)
            print('WPS_UPLOAD_HANDOFF: ' + json.dumps(task, ensure_ascii=False), flush=True)
            print('下一步必须读取 references/wps_upload.md，通过浏览器上传商品日报并验证在线链接；本地导出不等于上传成功。', flush=True)
        except Exception as exc:
            print(f'WPS 上传交接失败；本地商品日报保留: {exc}', file=sys.stderr)
            return 1
    return result


def main() -> int:
    parser = argparse.ArgumentParser(description="Generate a Tmall or Pinduoduo operating report.")
    parser.add_argument("--platform", choices=[*sorted(PLATFORM_SCRIPTS), "both"])
    parser.add_argument("--store-input", action="append", default=[], help="明确店铺日指标 CSV/XLSX，可重复")
    parser.add_argument("--store-output", help="新增店铺维度日报路径")
    parser.add_argument("--store-only", action="store_true", help="仅生成店铺维度，正常请求默认两种均输出")
    parser.add_argument('--archive-dir',type=Path,help='固定报表目录，默认运营数据库/数据报表输出')
    args, passthrough = parser.parse_known_args()
    archive_root=args.archive_dir or Path(option_value(passthrough,'--database-root',str(DEFAULT_DATABASE_ROOT)))/'数据报表输出'
    os.environ['DAILY_REPORT_ARCHIVE']='1'
    os.environ['DAILY_REPORT_ARCHIVE_DIR']=str(archive_root)
    os.environ['DAILY_REPORT_DATES']=json.dumps(sorted(requested_dates(passthrough)))
    store_arguments = []
    store_arguments.extend(['--archive-dir',str(archive_root)])
    for option in ("--database-root", "--start-date", "--end-date", "--date"):
        for index, value in enumerate(passthrough):
            if value == option and index + 1 < len(passthrough):
                store_arguments.extend([option, passthrough[index + 1]])
            elif value.startswith(option + "="):
                store_arguments.extend([option, value.split("=", 1)[1]])
    if args.store_output:
        store_arguments.extend(["--output", args.store_output])
    for source in args.store_input:
        store_arguments.extend(["--input", source])
    store_result = subprocess.run([sys.executable, str(SCRIPT_DIR / "build_store_report.py"), *store_arguments], check=False).returncode
    if args.store_only:
        return store_result

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

    if len(selected)>1 and option_value(passthrough,'--output',''):
        parser.error('多平台须分别保存，不能共用一个--output；请用--archive-dir指定同一目录')
    results = [run_platform(platform, passthrough) for platform in selected]
    return max([store_result, *results], default=0)


if __name__ == "__main__":
    raise SystemExit(main())
