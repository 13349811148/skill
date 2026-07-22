from __future__ import annotations

import argparse
import csv
import os
import re
import sys
from collections import defaultdict
from dataclasses import dataclass
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any

from promotion_protection import (
    PromotionProtectionError,
    PromotionSnapshotRow,
    date_bounds_from_filename,
    deduplicate_exact_files,
    infer_path_dimensions,
    reconcile_promotion_costs,
    require_single_day_value,
    select_latest_snapshots,
)

try:
    import win32com.client as win32
except Exception as exc:  # pragma: no cover - environment guard
    win32 = None
    WIN32_ERROR = exc
else:
    WIN32_ERROR = None


DEFAULT_USER_PROFILE = Path(os.environ.get("USERPROFILE", str(Path.home())))
DEFAULT_DESKTOP = DEFAULT_USER_PROFILE / "Desktop"
DEFAULT_DATABASE_ROOT = DEFAULT_DESKTOP / "运营数据库"
DEFAULT_TEMPLATE = DEFAULT_DESKTOP / "数据表（群内格式）.xls"
DEFAULT_MARKETING_DIR = DEFAULT_DATABASE_ROOT / "营销活动监控"
DEFAULT_MARKETING = DEFAULT_MARKETING_DIR / "营销活动.xls"
DEFAULT_COST_TABLE = DEFAULT_MARKETING_DIR / "成本表.xlsx"
DEFAULT_OUTPUT_DIR = DEFAULT_DATABASE_ROOT / "数据报表输出"
DEFAULT_OUTPUT = DEFAULT_OUTPUT_DIR / "数据报表拼多多.xlsx"
REFERENCE_TABLE_EXTENSIONS = {".xls", ".xlsx"}
PROMOTION_MECHANISM_COLUMN = "促销机制        (天猫参加活动填万人团\n淘宝参加活动填百亿补贴)"
TEMPLATE_SHEET = "每日销售数据"
ORDER_MAX_ROWS = 100_000
SMALL_RECEIPT_UPPER_BOUND = 1.0
COST_LOW_DEVIATION_THRESHOLD = 0.10
SMALL_RECEIPT_ACTIONS = ("confirm", "include", "exclude")
EXPECTED_UNIT_RECEIPT_COLUMN = "单件预估实收金额"
ARRIVAL_PRICE_MATCH_TOLERANCE = 0.005
ARRIVAL_PRICE_LOW_DEVIATION_LIMIT = 0.05
ARRIVAL_PRICE_WARNING_PREFIX = "到手价提醒："
PROMOTION_SHOP_CANDIDATES = ("店铺名称", "店铺", "汇总_店铺名称", "店铺名")
PROMOTION_PLAN_CANDIDATES = (
    "计划ID",
    "推广计划ID",
    "计划编号",
    "计划名称",
    "推广计划名称",
    "单元ID",
    "推广单元ID",
    "单元名称",
    "推广单元名称",
)

SKILL_ROOT = Path(__file__).resolve().parents[1]
COLUMNS_CSV = SKILL_ROOT / "references" / "template_columns.csv"


@dataclass
class WorkbookData:
    path: Path
    sheet_name: str
    rows: list[list[Any]]
    total_rows: int = 0
    truncated: bool = False


class ReportError(RuntimeError):
    pass


def text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def order_output_rows(
    normal_rows: list[dict[str, Any]],
    empty_burn_rows: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    """Keep all empty-burn promotion rows at the absolute bottom."""
    sort_key = lambda row: (
        text(row.get("日期")),
        text(row.get("店铺名称")),
        text(row.get("商品ID")),
        text(row.get("商品SKU")),
    )
    return sorted(normal_rows, key=sort_key) + sorted(empty_burn_rows, key=sort_key)


def number(value: Any) -> float:
    if value is None or value == "":
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    cleaned = str(value).replace(",", "").replace("￥", "").replace("元", "").strip()
    if cleaned in {"", "-", "--"}:
        return 0.0
    try:
        return float(cleaned)
    except ValueError:
        return 0.0


def parse_date(value: Any) -> str:
    if value is None or value == "":
        return ""
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    raw = str(value).strip()
    if not raw:
        return ""
    raw = raw.replace("/", "-").replace(".", "-")
    match = re.search(r"(20\d{2})[-年](\d{1,2})[-月](\d{1,2})", raw)
    if match:
        return f"{int(match.group(1)):04d}-{int(match.group(2)):02d}-{int(match.group(3)):02d}"
    match = re.search(r"(20\d{2})(\d{2})(\d{2})", raw)
    if match:
        return f"{match.group(1)}-{match.group(2)}-{match.group(3)}"
    try:
        return datetime.fromisoformat(raw[:10]).strftime("%Y-%m-%d")
    except ValueError:
        return raw[:10]


def parse_requested_date(value: str) -> str:
    parsed = parse_date(value)
    if not re.fullmatch(r"20\d{2}-\d{2}-\d{2}", parsed):
        raise ReportError(f"日期格式不正确: {value}。请使用 YYYY-MM-DD，例如 2026-06-25。")
    return parsed


def requested_dates_from_args(args: argparse.Namespace) -> set[str]:
    dates: set[str] = set()
    for item in args.date or []:
        for part in re.split(r"[,，\s]+", item):
            if part.strip():
                dates.add(parse_requested_date(part.strip()))

    if args.start_date or args.end_date:
        if not args.start_date or not args.end_date:
            raise ReportError("--start-date 和 --end-date 必须同时填写。")
        start = datetime.strptime(parse_requested_date(args.start_date), "%Y-%m-%d")
        end = datetime.strptime(parse_requested_date(args.end_date), "%Y-%m-%d")
        if end < start:
            raise ReportError("--end-date 不能早于 --start-date。")
        current = start
        while current <= end:
            dates.add(current.strftime("%Y-%m-%d"))
            current += timedelta(days=1)

    if not dates:
        raise ReportError(
            "请先指定要统计的日期。示例: --date 2026-06-25；多天可重复写 --date，或使用 --start-date 2026-06-25 --end-date 2026-06-27。"
        )
    return dates


def normalize_header(value: Any) -> str:
    return re.sub(r"\s+", "", text(value)).lower()


def find_col(headers: list[str], candidates: list[str]) -> str | None:
    normalized = {normalize_header(header): header for header in headers}
    for candidate in candidates:
        key = normalize_header(candidate)
        if key in normalized:
            return normalized[key]
    for header in headers:
        h = normalize_header(header)
        if any(normalize_header(candidate) in h for candidate in candidates):
            return header
    return None


def is_pdd_order_headers(headers: list[str]) -> bool:
    normalized = {normalize_header(header) for header in headers}
    required = {"样式id", "订单成交时间"}
    distinct = {"多多支付立减金额(元)", "团id"}
    return required <= normalized and bool(normalized & distinct)


def read_template_columns() -> list[str]:
    columns: list[str] = []
    with COLUMNS_CSV.open("r", encoding="utf-8-sig", newline="") as fh:
        reader = csv.DictReader(fh)
        for row in reader:
            columns.append(row["column"])
    return columns


def ensure_excel() -> Any:
    if win32 is None:
        raise ReportError(f"Cannot use Excel COM: {WIN32_ERROR}")
    excel = win32.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    return excel


def read_workbook(
    path: Path, sheet_name: str | None = None, max_rows: int | None = None
) -> WorkbookData:
    excel = ensure_excel()
    try:
        wb = excel.Workbooks.Open(str(path), ReadOnly=True)
        try:
            ws = wb.Worksheets(sheet_name) if sheet_name else wb.Worksheets(1)
            used = ws.UsedRange
            first_row = used.Row
            first_col = used.Column
            total_rows = int(used.Rows.Count)
            row_count = min(total_rows, max_rows) if max_rows is not None else total_rows
            col_count = int(used.Columns.Count)
            end_row = first_row + row_count - 1
            end_col = first_col + col_count - 1
            values = ws.Range(ws.Cells(first_row, first_col), ws.Cells(end_row, end_col)).Value
            if values is None:
                rows: list[list[Any]] = []
            elif not isinstance(values, tuple):
                rows = [[values]]
            else:
                rows = [list(row if isinstance(row, tuple) else (row,)) for row in values]
            return WorkbookData(
                path=path,
                sheet_name=ws.Name,
                rows=rows,
                total_rows=total_rows,
                truncated=max_rows is not None and total_rows > max_rows,
            )
        finally:
            wb.Close(False)
    finally:
        excel.Quit()


def read_all_workbook_sheets(path: Path) -> list[WorkbookData]:
    excel = ensure_excel()
    try:
        wb = excel.Workbooks.Open(str(path), ReadOnly=True)
        try:
            result: list[WorkbookData] = []
            for ws in wb.Worksheets:
                used = ws.UsedRange
                first_row = used.Row
                first_col = used.Column
                row_count = min(int(used.Rows.Count), 5000)
                col_count = min(int(used.Columns.Count), 100)
                end_row = first_row + row_count - 1
                end_col = first_col + col_count - 1
                values = ws.Range(ws.Cells(first_row, first_col), ws.Cells(end_row, end_col)).Value
                if values is None:
                    raw_rows: list[list[Any]] = []
                elif not isinstance(values, tuple):
                    raw_rows = [[values]]
                else:
                    raw_rows = [list(row if isinstance(row, tuple) else (row,)) for row in values]
                rows: list[list[Any]] = [row for row in raw_rows if any(text(value) for value in row)]
                result.append(WorkbookData(path=path, sheet_name=ws.Name, rows=rows))
            return result
        finally:
            wb.Close(False)
    finally:
        excel.Quit()


def rows_to_dicts(rows: list[list[Any]]) -> list[dict[str, Any]]:
    if not rows:
        return []
    headers = [text(value) for value in rows[0]]
    result: list[dict[str, Any]] = []
    for raw in rows[1:]:
        if not any(text(v) for v in raw):
            continue
        item = {headers[i]: raw[i] if i < len(raw) else "" for i in range(len(headers)) if headers[i]}
        result.append(item)
    return result


def discover_files(root: Path, subdir: str) -> list[Path]:
    folder = root / subdir
    if not folder.exists():
        return []
    suffixes = {".xls", ".xlsx", ".csv"}
    files = [
        path
        for path in folder.rglob("*")
        if path.is_file()
        and path.suffix.lower() in suffixes
        and "_合并前备份" not in str(path)
        and not path.name.startswith("~$")
    ]
    return sorted(files)


def reference_workbooks(folder: Path) -> list[Path]:
    if not folder.exists():
        return []
    return sorted(
        [
            path
            for path in folder.iterdir()
            if path.is_file()
            and path.suffix.lower() in REFERENCE_TABLE_EXTENSIONS
            and not path.name.startswith("~$")
        ],
        key=lambda path: path.name.lower(),
    )


def resolve_marketing_path(argument: str | None, database_root: Path) -> Path:
    if argument:
        return Path(argument)
    folder = database_root / "营销活动监控"
    candidates = reference_workbooks(folder)
    updated_candidates = [path for path in candidates if "营销活动_更新" in path.stem]
    if updated_candidates:
        return max(updated_candidates, key=lambda path: path.stat().st_mtime)
    for name in ("营销活动.xls", "营销活动.xlsx", "营销活动表.xlsx", "营销活动模板.xlsx"):
        preferred = folder / name
        if preferred in candidates:
            return preferred
    keyword_candidates = [path for path in candidates if "营销活动" in path.stem]
    if len(keyword_candidates) == 1:
        return keyword_candidates[0]
    if not keyword_candidates:
        raise ReportError(f"营销活动表不存在: {folder}；请放入营销活动表，或使用 --marketing 指定文件。")
    names = "、".join(path.name for path in keyword_candidates)
    raise ReportError(f"检测到多个营销活动表候选: {names}；请使用 --marketing 指定本次使用的文件。")


def resolve_cost_table_path(argument: str | None, database_root: Path, marketing_path: Path) -> Path:
    if argument:
        return Path(argument)
    folder = database_root / "营销活动监控"
    candidates = [path for path in reference_workbooks(folder) if path != marketing_path]
    non_template = [path for path in candidates if "模板" not in path.stem and "营销活动" not in path.stem]
    named_costs = [path for path in non_template if "成本" in path.stem]
    if len(named_costs) == 1:
        return named_costs[0]
    if len(named_costs) > 1:
        names = "、".join(path.name for path in named_costs)
        raise ReportError(f"检测到多个成本表候选: {names}；请使用 --cost-table 指定本次使用的文件。")
    if len(non_template) == 1:
        return non_template[0]
    if len(non_template) > 1:
        names = "、".join(path.name for path in non_template)
        raise ReportError(f"无法自动判断成本表，候选文件为: {names}；请使用 --cost-table 指定文件。")
    template_candidates = [path for path in candidates if "成本" in path.stem and "模板" in path.stem]
    if len(template_candidates) == 1:
        return template_candidates[0]
    if not template_candidates:
        raise ReportError(f"成本表不存在: {folder}；请放入实际成本表，或使用 --cost-table 指定文件。")
    names = "、".join(path.name for path in template_candidates)
    raise ReportError(f"检测到多个成本表模板: {names}；请使用 --cost-table 指定本次使用的文件。")


def marketing_promotion_mechanism(row: dict[str, Any]) -> Any:
    return (
        row.get("活动名称")
        or row.get(PROMOTION_MECHANISM_COLUMN)
        or row.get("促销机制")
        or row.get("促销机制（官补）")
        or ""
    )


def is_subsidy_activity(row: dict[str, Any]) -> bool:
    activity_price = text(row.get("活动价"))
    registration_price = text(row.get("报名价"))
    return bool(activity_price and registration_price and number(activity_price) != number(registration_price))


def resolve_arrival_price(
    expected_unit_receipt: Any,
    actual_amount: float,
    quantity: float,
    date: str,
    style_id: str,
    product_id: str,
) -> tuple[float | str, str | None]:
    """Return the PDD report arrival price and a user-facing exception warning."""
    actual_arrival_price = actual_amount / quantity if quantity else ""
    if actual_arrival_price == "":
        return "", None
    if text(expected_unit_receipt) == "":
        detail = f"日期={date}，样式ID={style_id or '未提供'}，商品ID={product_id or '未提供'}"
        return (
            actual_arrival_price,
            f"{ARRIVAL_PRICE_WARNING_PREFIX}营销活动表未填写单件预估实收金额，"
            f"已按实际计算结果填写到手价，请补充或更新营销活动表，{detail}，"
            f"实际={actual_arrival_price:.2f}",
        )

    expected_price = number(expected_unit_receipt)
    detail = f"日期={date}，样式ID={style_id or '未提供'}，商品ID={product_id or '未提供'}"
    if expected_price <= 0:
        return (
            actual_arrival_price,
            f"{ARRIVAL_PRICE_WARNING_PREFIX}单件预估实收金额无效，已按实际到手价填写，请检查营销活动表，{detail}，"
            f"预估={expected_price:.2f}，实际={actual_arrival_price:.2f}",
        )
    if abs(actual_arrival_price - expected_price) < ARRIVAL_PRICE_MATCH_TOLERANCE:
        return expected_price, None
    if actual_arrival_price < expected_price:
        deviation = (expected_price - actual_arrival_price) / expected_price
        if deviation <= ARRIVAL_PRICE_LOW_DEVIATION_LIMIT:
            return (
                actual_arrival_price,
                f"{ARRIVAL_PRICE_WARNING_PREFIX}实际到手价低于单件预估实收金额5%以内，已按实际到手价填写；"
                f"可能为多件折扣或营销活动表未更新，请检查，{detail}，"
                f"预估={expected_price:.2f}，实际={actual_arrival_price:.2f}，偏差={deviation:.2%}",
            )
        return (
            "",
            f"{ARRIVAL_PRICE_WARNING_PREFIX}实际到手价低于单件预估实收金额超过5%，到手价未填写，请检查多件折扣或营销活动表，"
            f"{detail}，预估={expected_price:.2f}，实际={actual_arrival_price:.2f}，偏差={deviation:.2%}",
        )
    deviation = (actual_arrival_price - expected_price) / expected_price
    return (
        actual_arrival_price,
        f"{ARRIVAL_PRICE_WARNING_PREFIX}实际到手价高于单件预估实收金额，已按实际到手价填写；可能活动失效、促销失效导致价格上涨，"
        f"请检查，{detail}，预估={expected_price:.2f}，实际={actual_arrival_price:.2f}，偏差={deviation:.2%}",
    )


def load_template_products(template_path: Path) -> dict[tuple[str, str], dict[str, Any]]:
    data = read_workbook(template_path, TEMPLATE_SHEET)
    products: dict[tuple[str, str], dict[str, Any]] = {}
    for row in rows_to_dicts(data.rows):
        product_id = text(row.get("商品ID"))
        sku = text(row.get("商品SKU"))
        if not product_id and not sku:
            continue
        key = (product_id, sku)
        products[key] = {
            "品牌": row.get("品牌", ""),
            "所属运营": row.get("所属运营", ""),
            "平台": row.get("平台", ""),
            "店铺名称": row.get("店铺名称", ""),
            "大类": row.get("大类", ""),
            "商品ID": product_id,
            "商品SKU": sku,
            "产品成本": row.get("产品成本", ""),
            "活动价": row.get("活动价", ""),
            "报名价": row.get("报名价", ""),
            EXPECTED_UNIT_RECEIPT_COLUMN: row.get(EXPECTED_UNIT_RECEIPT_COLUMN, ""),
            PROMOTION_MECHANISM_COLUMN: marketing_promotion_mechanism(row),
            "到手价": row.get("到手价", ""),
            "大类辅助列": row.get("大类辅助列", ""),
            "辅助": row.get("辅助", ""),
        }
    return products


def load_marketing_rows(marketing_path: Path) -> list[dict[str, Any]]:
    try:
        data = read_workbook(marketing_path, "营销数据")
    except Exception as exc:
        raise ReportError(f"营销活动表读取失败，已停止生成日报,{marketing_path},{exc}") from exc
    rows = rows_to_dicts(data.rows)
    if not rows:
        raise ReportError(f"营销活动表未读到有效商品记录，已停止生成日报,{marketing_path}")
    headers = list(rows[0].keys())
    missing = [field for field in ("商品ID", "样式ID", "商品SKU") if not find_col(headers, [field])]
    if missing:
        raise ReportError(
            f"营销活动表缺少字段{'|'.join(missing)}，已停止生成日报,{marketing_path}"
        )
    valid_rows = [
        row for row in rows if text(row.get("商品ID")) or text(row.get("样式ID")) or text(row.get("商品SKU"))
    ]
    if not valid_rows:
        raise ReportError(f"营销活动表未读到有效商品记录，已停止生成日报,{marketing_path}")
    return valid_rows


def load_marketing_styles(marketing_rows: list[dict[str, Any]]) -> dict[str, dict[str, Any]]:
    styles: dict[str, dict[str, Any]] = {}
    for row in marketing_rows:
        style_id = text(row.get("样式ID"))
        if not style_id:
            continue
        styles[style_id] = {
            "品牌": row.get("品牌", ""),
            "所属运营": row.get("所属运营", ""),
            "平台": row.get("平台", ""),
            "店铺名称": row.get("店铺名称", ""),
            "大类": row.get("大类", ""),
            "商品ID": text(row.get("商品ID")),
            "商品SKU": row.get("商品SKU", ""),
            "产品成本": row.get("产品成本", ""),
            "活动价": row.get("活动价", ""),
            "报名价": row.get("报名价", ""),
            EXPECTED_UNIT_RECEIPT_COLUMN: row.get(EXPECTED_UNIT_RECEIPT_COLUMN, ""),
            PROMOTION_MECHANISM_COLUMN: marketing_promotion_mechanism(row),
        }
    return styles


def load_marketing_products(marketing_rows: list[dict[str, Any]]) -> dict[tuple[str, str], dict[str, Any]]:
    products: dict[tuple[str, str], dict[str, Any]] = {}
    for row in marketing_rows:
        product_id = text(row.get("商品ID"))
        sku = text(row.get("商品SKU"))
        if product_id or sku:
            products[(product_id, sku)] = {
                "品牌": row.get("品牌", ""),
                "所属运营": row.get("所属运营", ""),
                "平台": row.get("平台", ""),
                "店铺名称": row.get("店铺名称", ""),
                "大类": row.get("大类", ""),
                "商品ID": product_id,
                "商品SKU": sku,
                "产品成本": row.get("产品成本", ""),
                "活动价": row.get("活动价", ""),
                "报名价": row.get("报名价", ""),
                EXPECTED_UNIT_RECEIPT_COLUMN: row.get(EXPECTED_UNIT_RECEIPT_COLUMN, ""),
                PROMOTION_MECHANISM_COLUMN: marketing_promotion_mechanism(row),
            }
    return products


def load_costs_by_merchant_code(cost_table_path: Path) -> tuple[dict[str, float], list[str]]:
    costs: dict[str, float] = {}
    warnings: list[str] = []
    try:
        sheets = read_all_workbook_sheets(cost_table_path)
    except Exception as exc:
        return costs, [f"成本表读取失败,{cost_table_path},{exc}"]

    for data in sheets:
        rows = rows_to_dicts(data.rows)
        if not rows:
            continue
        headers = list(rows[0].keys())
        code_col = find_col(headers, ["商家编码", "商品编码"])
        cost_col = find_col(headers, ["6.11成本价", "成本价", "产品成本", "成本"])
        if not code_col or not cost_col:
            continue
        for row in rows:
            merchant_code = text(row.get(code_col))
            if not merchant_code:
                continue
            if text(row.get(cost_col)) == "":
                warnings.append(f"成本表成本为空,{data.sheet_name},{merchant_code}")
                continue
            costs[merchant_code] = number(row.get(cost_col))

    if not costs:
        warnings.append(f"成本表未读取到可用成本,{cost_table_path}")
    return costs, warnings


def load_product_exports(root: Path) -> dict[str, dict[str, Any]]:
    products: dict[str, dict[str, Any]] = {}
    for path in discover_files(root, "商品数据"):
        try:
            workbook = read_workbook(path, max_rows=ORDER_MAX_ROWS)
            if workbook.truncated:
                warnings.append(f"商品数据文件超过{ORDER_MAX_ROWS}行，仅读取前{ORDER_MAX_ROWS}行,{path}")
            rows = rows_to_dicts(workbook.rows)
        except Exception:
            continue
        for row in rows:
            product_id = text(row.get("商品ID（必填）") or row.get("商品ID") or row.get("商品id"))
            if not product_id:
                continue
            products.setdefault(product_id, {})
            products[product_id].update(
                {
                    "商品ID": product_id,
                    "商品SKU": text(row.get("商品名称") or row.get("商品SKU") or products[product_id].get("商品SKU", "")),
                }
            )
    return products


def is_valid_order(row: dict[str, Any]) -> bool:
    status = text(row.get("订单状态"))
    after_sale = text(row.get("售后状态"))
    blocked_status = ["未成交", "关闭", "退款成功", "全额退款", "已退款"]
    blocked_after_sale = ["退款成功", "全额退款", "已退款"]
    return not any(word in status for word in blocked_status) and not any(
        word in after_sale for word in blocked_after_sale
    )


def is_small_actual_receipt(value: Any) -> bool:
    amount = number(value)
    return 0.0 < amount < SMALL_RECEIPT_UPPER_BOUND


def matching_marketing_row(
    style_id: str,
    product_id: str,
    sku: str,
    template_products: dict[tuple[str, str], dict[str, Any]],
    template_styles: dict[str, dict[str, Any]],
) -> dict[str, Any]:
    return template_styles.get(style_id) or template_products.get((product_id, sku)) or next(
        (value for (candidate_product_id, _sku), value in template_products.items() if candidate_product_id == product_id),
        {},
    )


def resolve_order_unit_cost(
    style_id: str,
    product_id: str,
    sku: str,
    merchant_spec_code: str,
    template_products: dict[tuple[str, str], dict[str, Any]],
    template_styles: dict[str, dict[str, Any]],
    cost_by_merchant_code: dict[str, float],
) -> float | None:
    if merchant_spec_code and merchant_spec_code in cost_by_merchant_code:
        cost_value = cost_by_merchant_code[merchant_spec_code]
        if number(cost_value) > 0.0:
            return number(cost_value)
    marketing_row = matching_marketing_row(
        style_id, product_id, sku, template_products, template_styles
    )
    cost_value = marketing_row.get("产品成本", "")
    if text(cost_value) != "" and number(cost_value) > 0.0:
        return number(cost_value)
    return None


def below_cost_review_reason(
    actual_amount: float, quantity: float, unit_cost: float
) -> str:
    if quantity <= 0.0 or unit_cost <= 0.0:
        return ""
    actual_unit_receipt = actual_amount / quantity
    deviation = (unit_cost - actual_unit_receipt) / unit_cost
    if deviation <= COST_LOW_DEVIATION_THRESHOLD:
        return ""
    return (
        f"单件实际实收{actual_unit_receipt:.2f}低于单位成本{unit_cost:.2f}超过10%"
        f"（偏差{deviation:.2%}）"
    )


def order_review_confirmation_error(details: list[str]) -> ReportError:
    preview = "\n".join(f"- {detail}" for detail in details[:20])
    remainder = "" if len(details) <= 20 else f"\n- 另有 {len(details) - 20} 条。"
    return ReportError(
        "发现需要人工确认的订单：单笔实收金额大于0且小于1元，或单件实际实收低于单位成本超过10%；"
        "为避免误计入，尚未生成报表。请先询问用户是否计入这些订单；确认后以 "
        "--small-receipt-action include（计入）或 "
        "--small-receipt-action exclude（不计入）重新运行。\n"
        f"{preview}{remainder}"
    )


def load_order_aggregates(
    root: Path,
    target_dates: set[str],
    small_receipt_action: str = "confirm",
    template_products: dict[tuple[str, str], dict[str, Any]] | None = None,
    template_styles: dict[str, dict[str, Any]] | None = None,
    cost_by_merchant_code: dict[str, float] | None = None,
) -> tuple[dict[tuple[str, str], dict[str, Any]], list[str]]:
    if small_receipt_action not in SMALL_RECEIPT_ACTIONS:
        raise ReportError(f"Unsupported small receipt action: {small_receipt_action}")
    template_products = template_products or {}
    template_styles = template_styles or {}
    cost_by_merchant_code = cost_by_merchant_code or {}
    aggregates: dict[tuple[str, str], dict[str, Any]] = {}
    warnings: list[str] = []
    review_details: list[str] = []
    small_receipt_count = 0
    below_cost_count = 0
    missing_cost_rows: set[tuple[str, str, str, str, str]] = set()
    for path in discover_files(root, "订单数据"):
        try:
            rows = rows_to_dicts(read_workbook(path).rows)
        except Exception as exc:
            warnings.append(f"订单文件读取失败,{path},{exc}")
            continue
        if not rows:
            continue
        headers = list(rows[0].keys())
        if not is_pdd_order_headers(headers):
            continue
        merchant_code_col = find_col(headers, ["商家编码-规格维度"])
        order_id_col = find_col(headers, ["订单号", "订单编号", "子订单编号", "主订单编号", "订单ID"])
        required = {
            "date": find_col(headers, ["订单成交时间", "支付时间"]),
            "style_id": find_col(headers, ["样式ID", "样式id"]),
            "shop": find_col(headers, ["汇总_店铺名称", "店铺名称"]),
            "product_id": find_col(headers, ["商品id", "商品ID"]),
            "sku": find_col(headers, ["商品规格", "商品", "商品名称"]),
            "qty": find_col(headers, ["商品数量(件)", "商品数量", "数量"]),
            "amount": find_col(headers, ["商家实收金额(元)", "用户实付金额(元)", "商品总价(元)"]),
        }
        missing = [name for name, col in required.items() if not col]
        if missing:
            warnings.append(f"订单文件缺字段,{path},{'|'.join(missing)}")
            continue
        for row in rows:
            if not is_valid_order(row):
                continue
            date = parse_date(row.get(required["date"]))
            if date not in target_dates:
                continue
            style_id = text(row.get(required["style_id"]))
            product_id = text(row.get(required["product_id"]))
            sku = text(row.get(required["sku"]))
            shop = text(row.get(required["shop"]))
            merchant_spec_code = text(row.get(merchant_code_col)) if merchant_code_col else ""
            if not date or not style_id:
                continue
            actual_receipt = row.get(required["amount"])
            quantity = number(row.get(required["qty"]))
            review_reasons: list[str] = []
            if is_small_actual_receipt(actual_receipt):
                small_receipt_count += 1
                review_reasons.append(
                    f"单笔实收{number(actual_receipt):.2f}大于0且小于1元"
                )
            unit_cost = resolve_order_unit_cost(
                style_id,
                product_id,
                sku,
                merchant_spec_code,
                template_products,
                template_styles,
                cost_by_merchant_code,
            )
            if quantity > 0.0 and unit_cost is None:
                missing_cost_rows.add((date, shop, product_id, style_id, merchant_spec_code))
            elif unit_cost is not None:
                cost_reason = below_cost_review_reason(
                    number(actual_receipt), quantity, unit_cost
                )
                if cost_reason:
                    below_cost_count += 1
                    review_reasons.append(cost_reason)
            if review_reasons:
                order_id = text(row.get(order_id_col)) if order_id_col else ""
                detail = (
                    f"文件={path.name}；日期={date}；订单={order_id or '未提供'}；"
                    f"店铺={shop or '未提供'}；商品ID={product_id or '未提供'}；"
                    f"规格={sku or style_id or '未提供'}；数量={text(row.get(required['qty'])) or '未提供'}；"
                    f"实收={number(actual_receipt):.2f}；原因={'、'.join(review_reasons)}"
                )
                review_details.append(detail)
                if small_receipt_action == "exclude":
                    continue
            key = (date, style_id)
            current = aggregates.setdefault(
                key,
                {
                    "日期": date,
                    "样式ID": style_id,
                    "店铺名称": shop,
                    "商品ID": product_id,
                    "商品SKU": sku,
                    "商家编码-规格维度": merchant_spec_code,
                    "实际成交数量（去退款去补单后）": 0.0,
                    "实际成交金额（去退款去补单后）": 0.0,
                },
            )
            if merchant_spec_code and not current.get("商家编码-规格维度"):
                current["商家编码-规格维度"] = merchant_spec_code
            current["实际成交数量（去退款去补单后）"] += quantity
            current["实际成交金额（去退款去补单后）"] += number(actual_receipt)
    if review_details:
        if small_receipt_action == "confirm":
            raise order_review_confirmation_error(review_details)
        decision = "计入" if small_receipt_action == "include" else "不计入"
        warnings.append(
            f"异常订单已按确认{decision},{len(review_details)}条,"
            f"小额实收{small_receipt_count}条,低于成本超过10%共{below_cost_count}条"
        )
    for date, shop, product_id, style_id, merchant_spec_code in sorted(missing_cost_rows):
        warnings.append(
            f"订单成本缺失，未执行低于成本10%检查,{date},{shop},{product_id},{style_id},{merchant_spec_code}"
        )
    return aggregates, warnings


def promo_date_from_filename(path: Path) -> str:
    start_date, _end_date = date_bounds_from_filename(path)
    return start_date


def load_promo_costs(
    root: Path, target_dates: set[str]
) -> tuple[dict[tuple[str, str], float], list[str]]:
    warnings: list[str] = []
    snapshot_rows: list[PromotionSnapshotRow] = []
    promotion_paths, duplicate_warnings = deduplicate_exact_files(
        discover_files(root, "推广数据"), "拼多多推广费"
    )
    warnings.extend(duplicate_warnings)
    for path in promotion_paths:
        file_start_date, file_end_date = date_bounds_from_filename(path)
        try:
            rows = rows_to_dicts(read_workbook(path).rows)
        except Exception as exc:
            warnings.append(f"推广文件读取失败,{path},{exc}")
            continue
        if not rows:
            continue
        headers = list(rows[0].keys())
        product_col = find_col(headers, ["商品ID", "商品id"])
        cost_col = find_col(headers, ["总花费(元)", "成交花费(元)"])
        date_col = find_col(headers, ["日期", "统计日期", "报表日期"])
        shop_col = find_col(headers, list(PROMOTION_SHOP_CANDIDATES))
        plan_col = find_col(headers, list(PROMOTION_PLAN_CANDIDATES))
        if not product_col or not cost_col:
            warnings.append(f"推广文件缺字段,{path},商品ID或花费")
            continue
        if file_start_date and file_end_date and file_start_date != file_end_date and not date_col:
            raise ReportError(
                f"拼多多区间推广文件没有逐行日期，无法生成准确日报，已停止：{path}。"
                "请使用含日期列的逐日明细，或改为单日推广文件。"
            )
        if not date_col and not file_start_date:
            raise ReportError(
                f"拼多多推广文件既没有日期列，文件名也没有单日日期，已停止：{path}。"
            )
        path_shop, path_plan_id = infer_path_dimensions(root, "推广数据", path)
        for row in rows:
            amount = number(row.get(cost_col))
            if abs(amount) <= 1e-12:
                continue
            product_id = text(row.get(product_col))
            if not product_id or not product_id.isdigit():
                raise ReportError(
                    "拼多多推广行有非零花费但商品ID无效，已停止："
                    f"文件={path}，商品ID={product_id or '空'}，花费={amount:.2f}。"
                )
            if date_col:
                try:
                    require_single_day_value(row.get(date_col), path, date_col)
                except PromotionProtectionError as exc:
                    raise ReportError(str(exc)) from exc
                date = parse_date(row.get(date_col))
            else:
                date = file_start_date
            if not re.fullmatch(r"20\d{2}-\d{2}-\d{2}", date):
                raise ReportError(
                    "拼多多推广行有非零花费但日期无效，已停止："
                    f"文件={path}，日期={date or '空'}，商品ID={product_id}。"
                )
            if date not in target_dates:
                continue
            snapshot_rows.append(
                PromotionSnapshotRow(
                    platform="拼多多",
                    shop=text(row.get(shop_col)) if shop_col else path_shop,
                    promotion_type="拼多多推广费",
                    plan_id=text(row.get(plan_col)) if plan_col else path_plan_id,
                    date=date,
                    product_id=product_id,
                    amount=amount,
                    source_path=path,
                )
            )
    try:
        costs, snapshot_warnings = select_latest_snapshots(snapshot_rows)
    except PromotionProtectionError as exc:
        raise ReportError(str(exc)) from exc
    warnings.extend(snapshot_warnings)
    return costs, warnings


def enrich_row(
    row: dict[str, Any],
    template_products: dict[tuple[str, str], dict[str, Any]],
    template_styles: dict[str, dict[str, Any]],
    product_exports: dict[str, dict[str, Any]],
) -> dict[str, Any]:
    product_id = text(row.get("商品ID"))
    sku = text(row.get("商品SKU"))
    style_id = text(row.get("样式ID"))
    style_base = template_styles.get(style_id)
    product_base = template_products.get((product_id, sku)) or next(
        (value for (pid, _sku), value in template_products.items() if pid == product_id),
        {},
    )
    if style_base:
        base = style_base
    else:
        base = product_base
    export = product_exports.get(product_id, {})
    style_sku = text(style_base.get("商品SKU")) if style_base else ""
    fallback_sku = text(base.get("商品SKU")) if base else ""
    result = dict(base)
    result.update({k: v for k, v in export.items() if v and not result.get(k)})
    result.update(row)
    if style_sku:
        result["商品SKU"] = style_sku
    elif fallback_sku:
        result["商品SKU"] = fallback_sku
    elif style_id:
        result["商品SKU"] = ""
    if not result.get("平台"):
        result["平台"] = "拼多多"
    return result


def build_empty_burn_promotion_row(
    date: str,
    product_id: str,
    promotion_fee: float,
    template_products: dict[tuple[str, str], dict[str, Any]],
    template_styles: dict[str, dict[str, Any]],
    product_exports: dict[str, dict[str, Any]],
) -> dict[str, Any]:
    """Build a product-level row for promotion spend with no effective sales."""
    row = enrich_row(
        {
            "日期": date,
            "商品ID": product_id,
            "实际成交数量（去退款去补单后）": 0.0,
            "实际成交金额（去退款去补单后）": 0.0,
        },
        template_products,
        template_styles,
        product_exports,
    )
    # Promotion exports are product-level.  Do not falsely attribute spend to
    # one SKU when the marketing table contains multiple SKUs for the product.
    marketing_skus = {
        text(value.get("商品SKU"))
        for (candidate_product_id, _sku), value in template_products.items()
        if candidate_product_id == product_id and text(value.get("商品SKU"))
    }
    if len(marketing_skus) > 1:
        row["商品SKU"] = ""
    row.update(
        {
            "到手价": "",
            "每单补贴金额": "",
            "总补贴金额": "",
            "净销售额（实际成交+总补贴金额）": 0.0,
            "毛利": 0.0,
            "毛利率": "",
            "推广费用": promotion_fee,
            "店铺费用": 0.0,
            "平摊管理费用": 0.0,
            "净利润": -promotion_fee,
            "定价是否合理": "建议缩减推广",
            "备注": "空烧推广费：无有效销售",
        }
    )
    return row


def build_report(
    database_root: Path,
    marketing_path: Path,
    target_dates: set[str],
    cost_table_path: Path = DEFAULT_COST_TABLE,
    small_receipt_action: str = "confirm",
) -> tuple[list[str], list[dict[str, Any]], list[str]]:
    columns = read_template_columns()
    marketing_rows = load_marketing_rows(marketing_path)
    template_products = load_marketing_products(marketing_rows)
    template_styles = load_marketing_styles(marketing_rows)
    known_marketing_product_ids = {
        text(row.get("商品ID")) for row in template_styles.values() if text(row.get("商品ID"))
    }
    cost_by_merchant_code, cost_warnings = load_costs_by_merchant_code(cost_table_path)
    product_exports = load_product_exports(database_root)
    order_rows, order_warnings = load_order_aggregates(
        database_root,
        target_dates,
        small_receipt_action,
        template_products,
        template_styles,
        cost_by_merchant_code,
    )
    promo_costs, promo_warnings = load_promo_costs(database_root, target_dates)
    product_sales: dict[tuple[str, str], float] = defaultdict(float)
    for aggregate in order_rows.values():
        date = text(aggregate.get("日期"))
        product_id = text(aggregate.get("商品ID"))
        product_sales[(date, product_id)] += number(aggregate.get("实际成交金额（去退款去补单后）"))

    empty_burn_promo_costs = {
        key: promo_total
        for key, promo_total in promo_costs.items()
        if promo_total > 0.0 and product_sales.get(key, 0.0) <= 0.0
    }

    output_rows: list[dict[str, Any]] = []
    unmatched_styles: dict[str, str] = {}
    unmatched_cost_merchant_codes: set[str] = set()
    negative_gross_missing_price_product_ids: set[str] = set()
    arrival_price_warnings: list[str] = []
    for key, aggregate in order_rows.items():
        date, _style_id = key
        style_id = text(aggregate.get("样式ID"))
        if style_id and style_id not in template_styles:
            unmatched_styles[style_id] = text(aggregate.get("商品ID"))
        row = enrich_row(aggregate, template_products, template_styles, product_exports)
        qty = number(row.get("实际成交数量（去退款去补单后）"))
        amount = number(row.get("实际成交金额（去退款去补单后）"))
        has_subsidy = is_subsidy_activity(row)
        has_activity_price = has_subsidy and text(row.get("活动价")) != ""
        has_register_price = has_subsidy and text(row.get("报名价")) != ""
        per_order_subsidy = (
            number(row.get("报名价")) - number(row.get("活动价")) if has_activity_price and has_register_price else ""
        )
        row["每单补贴金额"] = per_order_subsidy
        row["总补贴金额"] = qty * per_order_subsidy if per_order_subsidy != "" else ""
        subsidy = number(row.get("总补贴金额"))
        merchant_spec_code = text(row.get("商家编码-规格维度"))
        marketing_cost = row.get("产品成本", "")
        if merchant_spec_code and merchant_spec_code in cost_by_merchant_code:
            row["产品成本"] = cost_by_merchant_code[merchant_spec_code]
        elif text(marketing_cost) != "":
            row["产品成本"] = marketing_cost
        else:
            row["产品成本"] = ""
            if merchant_spec_code:
                unmatched_cost_merchant_codes.add(merchant_spec_code)
        has_cost = text(row.get("产品成本")) != ""
        cost = number(row.get("产品成本"))
        product_id = text(row.get("商品ID"))
        promo_total = promo_costs.get((date, product_id), 0.0)
        product_amount = product_sales.get((date, product_id), 0.0)
        row["到手价"], arrival_price_warning = resolve_arrival_price(
            row.get(EXPECTED_UNIT_RECEIPT_COLUMN, ""), amount, qty, date, style_id, product_id
        )
        if arrival_price_warning:
            arrival_price_warnings.append(arrival_price_warning)
        allocated_promo_fee = promo_total * amount / product_amount if promo_total and product_amount else 0.0
        row["推广费用"] = allocated_promo_fee if allocated_promo_fee else ""
        net_sales = amount + subsidy if amount or subsidy else 0.0
        shop_fee = net_sales * 0.01 if net_sales else ""
        management_fee = net_sales * 0.01 if net_sales else ""
        promo_fee = number(row.get("推广费用"))
        row["净销售额（实际成交+总补贴金额）"] = net_sales if net_sales else ""
        row["店铺费用"] = shop_fee
        row["平摊管理费用"] = management_fee
        row["毛利"] = net_sales - cost * qty if has_cost and qty else ""
        if (
            has_subsidy
            and row["毛利"] != ""
            and row["毛利"] < 0
            and (not has_activity_price or not has_register_price)
            and product_id
        ):
            negative_gross_missing_price_product_ids.add(product_id)
        row["毛利率"] = row["毛利"] / net_sales if row.get("毛利") != "" and net_sales else ""
        base_profit = net_sales - qty * cost - number(shop_fee) - number(management_fee) if net_sales and has_cost else ""
        row["净利润"] = base_profit - promo_fee if base_profit != "" else ""
        if row["净利润"] != "":
            if base_profit < 0:
                row["定价是否合理"] = "建议改价"
            elif row["净利润"] < 0:
                row["定价是否合理"] = "建议缩减推广"
            else:
                row["定价是否合理"] = "合理"
        output_rows.append({column: row.get(column, "") for column in columns})

    empty_burn_rows: list[dict[str, Any]] = []
    empty_burn_warnings: list[str] = []
    for (date, product_id), promotion_fee in sorted(empty_burn_promo_costs.items()):
        empty_burn_row = build_empty_burn_promotion_row(
            date,
            product_id,
            promotion_fee,
            template_products,
            template_styles,
            product_exports,
        )
        empty_burn_rows.append({column: empty_burn_row.get(column, "") for column in columns})
        empty_burn_warnings.append(
            f"空烧推广费已写入报表,无有效销售,{date},{product_id},{promotion_fee:.2f}"
        )

    output_rows = order_output_rows(output_rows, empty_burn_rows)
    try:
        reconcile_promotion_costs(promo_costs, output_rows, "拼多多")
    except PromotionProtectionError as exc:
        raise ReportError(str(exc)) from exc
    warnings = (
        order_warnings
        + promo_warnings
        + cost_warnings
        + arrival_price_warnings
        + empty_burn_warnings
    )
    for product_id in sorted(negative_gross_missing_price_product_ids):
        warnings.append(f"商品ID{product_id} 官补活动报名价活动价缺失")
    for merchant_spec_code in sorted(unmatched_cost_merchant_codes):
        warnings.append(f"成本表未匹配商家编码且营销活动表成本为空,{merchant_spec_code}")
    for style_id, product_id in sorted(unmatched_styles.items()):
        if product_id in known_marketing_product_ids:
            warnings.append(
                f"新样式ID未在营销活动表匹配到，商品ID已收录，通常是该商品链接近期编辑修改过SKU；已按商品ID兜底取营销活动信息，请补齐营销活动表中的样式ID和SKU,{style_id},{product_id}"
            )
        else:
            warnings.append(
                f"新样式ID未在营销活动表匹配到，商品ID未收录，通常是新上架商品；请新增商品ID、样式ID、SKU和成本等营销活动表信息；商品ID也未收录，无法兜底取营销活动信息,{style_id},{product_id}"
            )
    if not output_rows:
        warnings.append(f"没有生成明细行,请检查指定日期是否有有效订单: {','.join(sorted(target_dates))}")
    return columns, output_rows, warnings


def write_workbook(path: Path, columns: list[str], rows: list[dict[str, Any]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    if path.exists():
        try:
            path.unlink()
        except PermissionError as exc:
            raise ReportError(f"输出文件正在被占用，请先关闭后重试: {path}") from exc
    excel = ensure_excel()
    try:
        wb = excel.Workbooks.Add()
        ws = wb.Worksheets(1)
        ws.Name = "每日销售数据"
        column_numbers = {
            header: index
            for index, header in enumerate(columns, start=1)
            if header
        }
        text_columns = {
            "商品SKU",
            PROMOTION_MECHANISM_COLUMN,
            "定价是否合理",
            "备注",
            "大类辅助列",
            "辅助",
            "项目组",
            "管理类型",
            "品种",
            "产线",
        }
        two_decimal_columns = {
            "产品成本",
            "活动价",
            "报名价",
            "到手价",
            "实际成交金额（去退款去补单后）",
            "每单补贴金额",
            "总补贴金额",
            "净销售额（实际成交+总补贴金额）",
            "毛利",
            "推广费用",
            "店铺费用",
            "平摊管理费用",
            "净利润",
        }
        for header in text_columns:
            if header in column_numbers:
                ws.Columns(column_numbers[header]).NumberFormat = "@"
        if "商品ID" in column_numbers:
            ws.Columns(column_numbers["商品ID"]).NumberFormat = "0"
        for c, header in enumerate(columns, start=1):
            ws.Cells(1, c).Value = header
        for r, row in enumerate(rows, start=2):
            for c, header in enumerate(columns, start=1):
                value = row.get(header, "")
                if header == "商品ID" and text(value).isdigit():
                    value = int(text(value))
                elif header in text_columns:
                    value = text(value)
                ws.Cells(r, c).Value = value
        ws.Rows(1).Font.Bold = True
        ws.Rows(1).WrapText = True
        ws.Columns.AutoFit()
        if "日期" in column_numbers:
            ws.Columns(column_numbers["日期"]).NumberFormat = "yyyy-mm-dd"
        for header in two_decimal_columns:
            if header in column_numbers:
                ws.Columns(column_numbers[header]).NumberFormat = "#,##0.00"
        if "实际成交数量（去退款去补单后）" in column_numbers:
            ws.Columns(column_numbers["实际成交数量（去退款去补单后）"]).NumberFormat = "#,##0"
        if "毛利率" in column_numbers:
            ws.Columns(column_numbers["毛利率"]).NumberFormat = "0.00%"
        wb.SaveAs(str(path), FileFormat=51)
        wb.Close(False)
    finally:
        excel.Quit()


def write_log(path: Path, rows_count: int, warnings: list[str]) -> None:
    log_path = path.with_name(path.stem + "_生成日志.csv")
    with log_path.open("w", encoding="utf-8-sig", newline="") as fh:
        writer = csv.writer(fh)
        writer.writerow(["time", "level", "message"])
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        writer.writerow([now, "info", f"generated_rows={rows_count}"])
        for warning in warnings:
            writer.writerow([now, "warning", warning])


def print_arrival_price_dialogue_reminders(warnings: list[str]) -> None:
    reminders = [
        warning for warning in warnings if warning.startswith(ARRIVAL_PRICE_WARNING_PREFIX)
    ]
    if not reminders:
        return
    print("到手价对话提醒（请直接发送给用户）:")
    for reminder in reminders[:20]:
        print(f"- {reminder}")
    if len(reminders) > 20:
        print(f"- 另有 {len(reminders) - 20} 条到手价提醒，详见生成日志。")


def main() -> int:
    parser = argparse.ArgumentParser(description="Generate 数据报表拼多多.xlsx")
    parser.add_argument("--database-root", default=str(DEFAULT_DATABASE_ROOT))
    parser.add_argument("--template", default=str(DEFAULT_TEMPLATE))
    parser.add_argument("--marketing", help="营销活动表路径。省略时从<数据库根目录>\\营销活动监控自动识别。")
    parser.add_argument("--cost-table", help="成本表路径。省略时从<数据库根目录>\\营销活动监控自动识别。")
    parser.add_argument("--output", default=str(DEFAULT_OUTPUT))
    parser.add_argument("--date", action="append", help="要统计的日期，格式 YYYY-MM-DD。可重复填写，或用逗号分隔多天。")
    parser.add_argument("--start-date", help="连续日期范围的开始日期，格式 YYYY-MM-DD。")
    parser.add_argument("--end-date", help="连续日期范围的结束日期，格式 YYYY-MM-DD。")
    parser.add_argument(
        "--small-receipt-action",
        choices=SMALL_RECEIPT_ACTIONS,
        default="confirm",
        help="小额实收或单件实收低于成本超过10%%时的处理：confirm（先停止并等待确认）、include（计入）、exclude（不计入）。",
    )
    args = parser.parse_args()

    database_root = Path(args.database_root)
    marketing_path = resolve_marketing_path(args.marketing, database_root)
    cost_table_path = resolve_cost_table_path(args.cost_table, database_root, marketing_path)
    output_path = Path(args.output)
    if not database_root.exists():
        raise ReportError(f"Database root does not exist: {database_root}")
    if not marketing_path.exists():
        raise ReportError(f"营销活动表不存在: {marketing_path}")
    if not cost_table_path.exists():
        raise ReportError(f"成本表不存在: {cost_table_path}")

    target_dates = requested_dates_from_args(args)
    columns, rows, warnings = build_report(
        database_root,
        marketing_path,
        target_dates,
        cost_table_path,
        args.small_receipt_action,
    )
    write_workbook(output_path, columns, rows)
    write_log(output_path, len(rows), warnings)
    print(f"输出文件: {output_path}")
    print(f"统计日期: {', '.join(sorted(target_dates))}")
    print(f"生成行数: {len(rows)}")
    other_warnings = [
        warning for warning in warnings if not warning.startswith(ARRIVAL_PRICE_WARNING_PREFIX)
    ]
    if other_warnings:
        print("提醒:")
        for warning in other_warnings[:20]:
            print(f"- {warning}")
        if len(other_warnings) > 20:
            print(f"- 还有 {len(other_warnings) - 20} 条提醒，见生成日志。")
    print_arrival_price_dialogue_reminders(warnings)
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except ReportError as exc:
        print(f"生成停止: {exc}", file=sys.stderr)
        raise SystemExit(2)
    except Exception as exc:
        print(f"生成失败: {exc}", file=sys.stderr)
        raise
