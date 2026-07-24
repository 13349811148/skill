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


def load_cost_table_records(
    cost_table_path: Path,
    sheets: Iterable[Any],
) -> tuple[dict[str, dict[str, Any]], list[str]]:
    """Read cost rows once for both platforms with row-level code fallback.

    ``sheets`` deliberately uses the common WorkbookData shape (``rows``,
    ``sheet_name``, ``truncated``) so the two report builders share exactly the
    same matching and duplicate-protection behavior.
    """
    costs: dict[str, dict[str, Any]] = {}
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
        for index, candidate_row in enumerate(rows[:20]):
            candidate_headers = [text(value) for value in candidate_row]
            if _supported_code_headers(candidate_headers) and _first_supported_header(
                candidate_headers, COST_TABLE_COST_CANDIDATES
            ):
                header_index = index
                headers = candidate_headers
                break
        if header_index is None:
            continue

        code_headers = _supported_code_headers(headers)
        cost_header = _first_supported_header(headers, COST_TABLE_COST_CANDIDATES)
        header_positions = {header: index for index, header in enumerate(headers) if header}
        cost_position = header_positions[cost_header]
        metadata_positions = {
            field: header_positions[actual]
            for field in COST_METADATA_FIELDS
            if (actual := _first_supported_header(headers, (field,)))
        }
        sheet_name = text(getattr(data, "sheet_name", "未命名工作表")) or "未命名工作表"

        for row_index, values in enumerate(rows[header_index + 1 :], start=header_index + 2):
            selected_header = ""
            raw_code = ""
            for code_header in code_headers:
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

            raw_cost = values[cost_position] if cost_position < len(values) else ""
            if text(raw_cost) == "":
                warnings.append(f"成本表成本为空，已忽略；工作表={sheet_name}，行号={row_index}，编码={raw_code}")
                continue
            try:
                cost_value = parse_number_strict(
                    raw_cost,
                    field_name="成本",
                    context=f"文件={cost_table_path}，工作表={sheet_name}，行号={row_index}，编码={raw_code}",
                )
            except DataQualityError as exc:
                warnings.append(f"成本表成本无效，已忽略；{exc}")
                continue
            if cost_value <= 0.0:
                warnings.append(
                    f"成本表成本必须大于0，已忽略；工作表={sheet_name}，行号={row_index}，"
                    f"编码={raw_code}，成本={cost_value:.2f}"
                )
                continue

            source = f"成本表:{cost_table_path.name}/{sheet_name}/{selected_header}={raw_code}"
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
    }
