from __future__ import annotations

import math
import re
import unicodedata
from collections import Counter
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable, Mapping, Sequence


COST_TABLE_MAX_ROWS = 10_000
COST_SOURCE_LOG_COLUMNS = (
    "平台",
    "行类型",
    "日期",
    "店铺名称",
    "商品ID",
    "样式ID",
    "订单规格编码",
    "产品成本",
    "成本来源",
    "营销成本来源",
    "营销成本",
    "成本表成本",
    "成本差异（营销-成本表）",
    "快递成本",
    "运费来源",
    "运费匹配状态",
)
COST_TABLE_CODE_CANDIDATES = (
    "商家编码-规格维度",
    "商品SKU",
    "商品编码",
    "商品SKU名称",
    "SKU名称",
    "SKU编码",
)
COST_TABLE_COST_CANDIDATES = ("6.11成本价", "成本价", "产品成本", "成本")
COST_METADATA_FIELDS = ("产线", "项目组", "管理类型", "品种")
ORDER_SPEC_CODE_CANDIDATES = (
    "商家编码-规格维度",
    "商家编码（规格维度）",
    "规格维度商家编码",
    "规格商家编码",
    "SKU商家编码",
    "SKU编码",
)


class DataQualityError(RuntimeError):
    pass


@dataclass(frozen=True)
class MarketingCostMatch:
    record: Mapping[str, Any]
    source: str


@dataclass(frozen=True)
class CostResolution:
    cost: float | None
    cost_record: Mapping[str, Any]
    source: str
    marketing_record: Mapping[str, Any]
    specification_conflict_warning: str = ""


def text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def normalize_match_code(value: Any) -> str:
    raw = unicodedata.normalize("NFKC", text(value))
    return re.sub(r"\s+", "", raw).casefold()


def parse_number_strict(
    value: Any,
    *,
    field_name: str,
    context: str,
    blank_as_zero: bool = False,
) -> float:
    raw = text(value)
    if raw in {"", "-", "--"}:
        if blank_as_zero:
            return 0.0
        raise DataQualityError(f"{field_name}为空：{context}")
    if isinstance(value, (int, float)):
        result = float(value)
    else:
        cleaned = raw.replace(",", "").replace("￥", "").replace("元", "").strip()
        try:
            result = float(cleaned)
        except ValueError as exc:
            raise DataQualityError(
                f"{field_name}不是有效数字：值={raw}，{context}"
            ) from exc
    if not math.isfinite(result):
        raise DataQualityError(f"{field_name}不是有限数字：值={raw}，{context}")
    return result


def optional_number(value: Any) -> float | None:
    if text(value) == "":
        return None
    try:
        return parse_number_strict(
            value, field_name="数值", context="营销活动表", blank_as_zero=False
        )
    except DataQualityError:
        return None


def unit_subsidy_amount(marketing_row: Mapping[str, Any]) -> float:
    """Return the valid per-unit subsidy recorded by the marketing activity row."""
    activity_price = optional_number(marketing_row.get("活动价"))
    registration_price = optional_number(marketing_row.get("报名价"))
    if (
        activity_price is None
        or registration_price is None
        or registration_price <= activity_price
    ):
        return 0.0
    return registration_price - activity_price


def _header_key(value: Any) -> str:
    return normalize_match_code(value)


def _first_supported_header(headers: Sequence[str], candidates: Sequence[str]) -> str:
    normalized_headers = { _header_key(header): header for header in headers if text(header) }
    for candidate in candidates:
        header = normalized_headers.get(_header_key(candidate))
        if header:
            return header
    return ""


def _supported_code_headers(headers: Sequence[str]) -> list[str]:
    """Return actual code headers in required per-row fallback order."""
    normalized_headers = { _header_key(header): header for header in headers if text(header) }
    selected: list[str] = []
    for candidate in COST_TABLE_CODE_CANDIDATES:
        header = normalized_headers.get(_header_key(candidate))
        if header and header not in selected:
            selected.append(header)
    return selected


def find_product_cost_table(database_root: Path) -> Path | None:
    """Prefer the explicitly maintained V3 cost table, not product exports."""
    folder = database_root / "1.产品数据库"
    preferred = folder / "产品成本.xlsx"
    if preferred.is_file():
        return preferred
    candidates = sorted(
        p for p in folder.iterdir()
        if p.is_file() and p.suffix.lower() in {".xlsx", ".xls"}
        and "成本" in p.stem and not p.name.startswith("~$")
        and not any(word in p.stem for word in ("运费", "快递", "物流"))
        and not any(word in p.stem.lower() for word in ("模板", "备份", "副本", "backup"))
    ) if folder.is_dir() else []
    if len(candidates) > 1:
        raise DataQualityError(
            "产品数据库存在多个成本表候选，请用 --cost-table 指定："
            + "、".join(str(p) for p in candidates)
        )
    return candidates[0] if candidates else None


class CostRecords(dict):
    """Usable cost records plus independently matched specification names."""

    def __init__(self, sku_names: dict[str, str]):
        super().__init__()
        self.sku_names = sku_names


def load_cost_sku_names(sheets: Iterable[Any]) -> dict[str, str]:
    """Match cost-table product names by exact SKU, even without a valid cost."""
    names: dict[str, str] = {}
    for data in sheets:
        rows = getattr(data, "rows", [])
        for index, candidate in enumerate(rows[:20]):
            headers = [text(v) for v in candidate]
            name_header = _first_supported_header(headers, ("产品名称", "商品名称"))
            summary = {"编码类型", "组合装编码", "货品编码（组合内）"}.issubset(headers)
            code_headers = _supported_code_headers(headers)
            if not name_header or not (summary or code_headers):
                continue
            positions = {h: i for i, h in enumerate(headers) if h}
            for values in rows[index + 1:]:
                def value(header: str) -> str:
                    p = positions[header]
                    return text(values[p]) if p < len(values) else ""
                current_headers = code_headers
                if summary:
                    kind = value("编码类型")
                    if kind not in {"组合装", "货品"}:
                        continue
                    current_headers = ["组合装编码" if kind == "组合装" else "货品编码（组合内）"]
                raw_code = next((value(h) for h in current_headers if value(h)), "")
                code = normalize_match_code(raw_code)
                name = " ".join(value(name_header).split())
                if not code or not name:
                    continue
                if code in names and names[code] != name:
                    raise DataQualityError(f"成本表同一SKU编码对应不同商品名称，已停止：{raw_code}；{names[code]}；{name}")
                names[code] = name
            break
    return names


def load_cost_table_records(
    cost_table_path: Path,
    sheets: Iterable[Any],
) -> tuple[dict[str, dict[str, Any]], list[str]]:
    """Read cost rows once for both platforms with row-level code fallback.

    ``sheets`` deliberately uses the common WorkbookData shape (``rows``,
    ``sheet_name``, ``truncated``) so the two report builders share exactly the
    same matching and duplicate-protection behavior.
    """
    sheets = list(sheets)
    costs: dict[str, dict[str, Any]] = CostRecords(load_cost_sku_names(sheets))
    from goods_catalog import GoodsCatalog
    costs.goods_catalog = GoodsCatalog.from_sheets(sheets)
    warnings: list[str] = []
    for data in sheets:
        if getattr(data, "truncated", False):
            warnings.append(
                f"成本表工作表超过{COST_TABLE_MAX_ROWS}行，仅读取前{COST_TABLE_MAX_ROWS}行，"
                f"{cost_table_path}，{getattr(data, 'sheet_name', '未命名工作表')}"
            )
        rows = getattr(data, "rows", [])
        header_index: int | None = None
        headers: list[str] = []
        product_summary = False
        for index, candidate_row in enumerate(rows[:20]):
            candidate_headers = [text(value) for value in candidate_row]
            summary_headers = {"编码类型", "组合装编码", "货品编码（组合内）", "匹配成本(元)", "匹配状态"}
            is_summary = summary_headers.issubset(candidate_headers)
            if is_summary or (_supported_code_headers(candidate_headers) and _first_supported_header(
                candidate_headers, COST_TABLE_COST_CANDIDATES
            )):
                header_index = index
                headers = candidate_headers
                product_summary = is_summary
                break
        if header_index is None:
            continue

        code_headers = _supported_code_headers(headers)
        cost_header = "匹配成本(元)" if product_summary else _first_supported_header(headers, COST_TABLE_COST_CANDIDATES)
        header_positions = {header: index for index, header in enumerate(headers) if header}
        metadata_positions = {
            field: header_positions[actual]
            for field in COST_METADATA_FIELDS
            if (actual := _first_supported_header(headers, (field,)))
        }
        sheet_name = text(getattr(data, "sheet_name", "未命名工作表")) or "未命名工作表"
        skipped_statuses: Counter[str] = Counter()
        reference_count = 0

        for row_index, values in enumerate(rows[header_index + 1 :], start=header_index + 2):
            row_code_headers = code_headers
            row_cost_headers = [cost_header]
            if product_summary:
                def field(name: str) -> str:
                    position = header_positions[name]
                    return text(values[position]) if position < len(values) else ""
                if not any(text(v) for v in values):
                    continue
                status = field("匹配状态")
                if status not in {"已匹配", "参考测算"}:
                    skipped_statuses[status or "状态为空"] += 1
                    continue
                row_cost_headers = (["匹配成本(元)"] if status == "已匹配" else []) + ["参考测算成本(元)"]
                row_cost_headers = [header for header in row_cost_headers if header in header_positions]
                kind = field("编码类型")
                if kind not in {"组合装", "货品"}:
                    warnings.append(f"成本汇总编码类型无效，已忽略；工作表={sheet_name}，行号={row_index}，类型={kind}")
                    continue
                # Component expressions belong to the bundle, never to an order SKU.
                row_code_headers = ["组合装编码" if kind == "组合装" else "货品编码（组合内）"]
            selected_header = ""
            raw_code = ""
            for code_header in row_code_headers:
                raw_value = values[header_positions[code_header]] if header_positions[code_header] < len(values) else ""
                if text(raw_value):
                    selected_header = code_header
                    raw_code = text(raw_value)
                    break
            if not selected_header:
                continue
            match_code = normalize_match_code(raw_code)
            if not match_code:
                continue

            cost_value = None
            selected_cost_header = ""
            has_cost_value = False
            for candidate_header in row_cost_headers:
                position = header_positions[candidate_header]
                raw_cost = values[position] if position < len(values) else ""
                if text(raw_cost) == "":
                    continue
                has_cost_value = True
                try:
                    candidate_cost = parse_number_strict(
                        raw_cost,
                        field_name="成本",
                        context=f"文件={cost_table_path}，工作表={sheet_name}，行号={row_index}，编码={raw_code}，成本表头={candidate_header}",
                    )
                except DataQualityError as exc:
                    warnings.append(f"成本表成本无效，已忽略该值；{exc}")
                    continue
                if candidate_cost <= 0.0:
                    warnings.append(
                        f"成本表成本必须大于0，已忽略该值；工作表={sheet_name}，行号={row_index}，"
                        f"编码={raw_code}，成本表头={candidate_header}，成本={candidate_cost:.2f}"
                    )
                    continue
                cost_value = candidate_cost
                selected_cost_header = candidate_header
                break
            if cost_value is None:
                if not has_cost_value:
                    warnings.append(f"成本表成本为空，已忽略；工作表={sheet_name}，行号={row_index}，编码={raw_code}")
                continue

            source = f"成本表:{cost_table_path.name}/{sheet_name}/{selected_header}={raw_code}"
            cost_kind = "参考测算" if selected_cost_header == "参考测算成本(元)" else "已匹配"
            if product_summary:
                source += f"/{selected_cost_header}（{cost_kind}）"
            record = {
                "产品成本": cost_value,
                **{
                    field: values[position] if position < len(values) else ""
                    for field, position in metadata_positions.items()
                },
                "来源": source,
                "原始编码": raw_code,
                "工作表": sheet_name,
                "行号": row_index,
                "编码表头": selected_header,
                "成本表头": selected_cost_header,
                "成本性质": cost_kind,
                "SKU名称": costs.sku_names.get(match_code, ""),
            }
            previous = costs.get(match_code)
            if previous:
                previous_cost = float(previous["产品成本"])
                if math.isclose(previous_cost, cost_value, rel_tol=0.0, abs_tol=1e-9):
                    warnings.append(
                        "成本表规范化编码重复且成本相同，已去重并保留首条；"
                        f"标准化编码={match_code}，首条={previous['来源']}（成本={previous_cost:.2f}），"
                        f"重复条={source}（成本={cost_value:.2f}）"
                    )
                    continue
                raise DataQualityError(
                    "成本表规范化编码成本冲突，已停止生成日报；"
                    f"标准化编码={match_code}；"
                    f"记录1：工作表={previous['工作表']}，行号={previous['行号']}，"
                    f"原始编码={previous['原始编码']}，成本={previous_cost:.2f}；"
                    f"记录2：工作表={sheet_name}，行号={row_index}，"
                    f"原始编码={raw_code}，成本={cost_value:.2f}"
                )
            costs[match_code] = record
            if product_summary and cost_kind == "参考测算":
                reference_count += 1

        if reference_count:
            warnings.append(
                f"产品成本汇总采用参考测算成本{reference_count}条，成本来源已标注参考测算；"
                f"文件={cost_table_path.name}，工作表={sheet_name}"
            )
        if skipped_statuses:
            detail = "、".join(f"{status}{count}条" for status, count in sorted(skipped_statuses.items()))
            warnings.append(
                f"产品成本汇总采用已匹配成本或可用参考测算成本，无可用成本的项目留空；"
                f"文件={cost_table_path.name}，工作表={sheet_name}；跳过{detail}"
            )

    if not costs:
        raise DataQualityError(f"成本表未读取到可用正数成本，已停止生成日报，{cost_table_path}")
    return costs, warnings


def _positive_cost(record: Mapping[str, Any]) -> float | None:
    value = optional_number(record.get("产品成本", ""))
    return value if value is not None and value > 0.0 else None


def resolve_marketing_cost_record(
    style_id: str,
    product_id: str,
    sku: str,
    template_products: Mapping[tuple[str, str], Mapping[str, Any]],
    template_styles: Mapping[str, Mapping[str, Any]],
) -> MarketingCostMatch | None:
    """Apply the documented marketing-table fallback order without ambiguity."""
    style_record = template_styles.get(style_id, {}) if style_id else {}
    if style_record and _positive_cost(style_record) is not None:
        return MarketingCostMatch(style_record, f"样式ID={style_id}")

    exact_record = template_products.get((product_id, sku), {})
    if exact_record and _positive_cost(exact_record) is not None:
        return MarketingCostMatch(exact_record, f"商品ID+商品SKU={product_id}+{sku}")

    candidate_records = [
        record
        for (candidate_product_id, _candidate_sku), record in template_products.items()
        if candidate_product_id == product_id and _positive_cost(record) is not None
    ]
    candidate_costs = {_positive_cost(record) for record in candidate_records}
    if len(candidate_costs) == 1 and candidate_records:
        return MarketingCostMatch(candidate_records[0], f"商品ID唯一成本={product_id}")
    return None


def resolve_final_cost(
    *,
    style_id: str,
    product_id: str,
    final_order_sku: str,
    cost_by_merchant_code: Mapping[str, Mapping[str, Any]],
    template_products: Mapping[tuple[str, str], Mapping[str, Any]],
    template_styles: Mapping[str, Mapping[str, Any]],
    marketing_source_prefix: str,
) -> CostResolution:
    """Use final order SKU first, then the safe marketing-table fallback order."""
    normalized_sku = normalize_match_code(final_order_sku)
    cost_record = cost_by_merchant_code.get(normalized_sku, {}) if normalized_sku else {}
    table_cost = _positive_cost(cost_record)
    marketing_match = resolve_marketing_cost_record(
        style_id, product_id, final_order_sku, template_products, template_styles
    )
    marketing_record = marketing_match.record if marketing_match else {}
    if table_cost is not None:
        specification_conflict_warning = ""
        style_record = template_styles.get(style_id, {}) if style_id else {}
        marketing_sku = text(style_record.get("商品SKU"))
        if marketing_sku and normalize_match_code(marketing_sku) != normalized_sku:
            marketing_cost = _positive_cost(style_record)
            specification_conflict_warning = (
                "订单SKU与营销表样式SKU冲突，已使用成本表最终SKU精确匹配；"
                f"商品ID={product_id or '未提供'}；样式ID={style_id or '未提供'}；"
                f"订单商品SKU={final_order_sku or '未提供'}；营销表商品SKU={marketing_sku}；"
                f"成本表成本={table_cost:.2f}；营销表成本="
                f"{f'{marketing_cost:.2f}' if marketing_cost is not None else '未提供'}"
            )
        return CostResolution(
            table_cost,
            cost_record,
            text(cost_record.get("来源")) or "成本表:未提供来源",
            marketing_record,
            specification_conflict_warning,
        )
    if marketing_match:
        return CostResolution(
            _positive_cost(marketing_match.record),
            {},
            f"{marketing_source_prefix}/{marketing_match.source}",
            marketing_match.record,
        )
    return CostResolution(None, {}, "未匹配", {})


def duplicate_identifier_warnings(
    rows: Iterable[dict[str, Any]], field_name: str, label: str
) -> list[str]:
    counts = Counter(text(row.get(field_name)) for row in rows if text(row.get(field_name)))
    return [
        f"{label}重复，请检查；字段={field_name}，值={value}，出现{count}次"
        for value, count in sorted(counts.items())
        if count > 1
    ]


def marketing_price_warnings(rows: Iterable[dict[str, Any]]) -> list[str]:
    warnings: list[str] = []
    for row in rows:
        activity_raw = text(row.get("活动价"))
        registration_raw = text(row.get("报名价"))
        if not activity_raw or not registration_raw:
            continue
        activity_price = optional_number(activity_raw)
        registration_price = optional_number(registration_raw)
        identity = (
            f"样式ID={text(row.get('样式ID')) or '未提供'}，"
            f"商品ID={text(row.get('商品ID')) or '未提供'}，"
            f"商品SKU={text(row.get('商品SKU')) or '未提供'}"
        )
        if activity_price is None or registration_price is None:
            warnings.append(
                "营销活动价格不是有效数字，请检查；"
                f"{identity}，活动价={activity_raw}，报名价={registration_raw}"
            )
        elif registration_price < activity_price:
            warnings.append(
                "报名价小于活动价，疑似营销活动表数据错误；本行不计算补贴，请检查；"
                f"{identity}，活动价={activity_price:.2f}，报名价={registration_price:.2f}"
            )
    return warnings


def cost_source_log(
    *,
    platform: str,
    row_type: str,
    date: str,
    shop: str,
    product_id: str,
    style_id: str,
    order_spec_code: str,
    cost: Any,
    source: str,
    marketing_source: str = "",
    marketing_cost: Any = "",
    table_cost: Any = "",
    cost_difference: Any = "",
) -> dict[str, str]:
    return {
        "平台": platform,
        "行类型": row_type,
        "日期": date or "未提供",
        "店铺名称": shop or "未提供",
        "商品ID": product_id or "未提供",
        "样式ID": style_id or "未提供",
        "订单规格编码": order_spec_code or "未提供",
        "产品成本": text(cost) or "未匹配",
        "成本来源": source,
        "营销成本来源": marketing_source,
        "营销成本": text(marketing_cost),
        "成本表成本": text(table_cost),
        "成本差异（营销-成本表）": text(cost_difference),
        "快递成本": "0" if row_type == "空烧推广费" else "",
        "运费来源": "无有效销售，不计快递成本" if row_type == "空烧推广费" else "",
        "运费匹配状态": "不适用（空烧推广费）" if row_type == "空烧推广费" else "",
    }
