from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Callable, Iterable


HEADER_SEARCH_ROWS = 20
ORDER_PAYMENT_DATE_CANDIDATES = (
    "订单支付时间",
    "支付时间",
    "订单付款时间",
    "付款时间",
)
ORDER_TRANSACTION_DATE_CANDIDATES = ("订单成交时间",)
ORDER_CREATION_DATE_CANDIDATES = ("订单创建时间",)
PROMOTION_DATE_CANDIDATES = (
    "归档日期",
    "日期",
    "业务日期",
    "统计日期",
    "统计时间",
    "推广日期",
    "报表日期",
)
PRODUCT_ID_CANDIDATES = (
    "归档商品ID",
    "商品ID",
    "商品id",
    "商品Id",
    "宝贝ID",
    "主体ID",
    "商品编号",
    "商品编码",
    "商品ID（必填）",
)
PROMOTION_ARCHIVE_TYPE_CANDIDATES = ("归档推广类型",)
SKU_ID_CANDIDATES = (
    "样式ID",
    "样式id",
    "SKU ID",
    "SKU_ID",
    "SKUID",
    "SKUID（必填，注意不是SKU编码）",
    "skuId",
)
SPEC_CODE_CANDIDATES = (
    "商家编码-规格维度",
    "商家编码（规格维度）",
    "规格维度商家编码",
    "规格商家编码",
    "SKU商家编码",
    "SKU编码",
    "规格编码",
    "商品SKU",
)

PROMOTION_PDD = "pdd"
PROMOTION_TMALL_WANXIANG = "tmall_wanxiang"
PROMOTION_TMALL_NEW = "tmall_new_customer"
PROMOTION_TMALL_OLD = "tmall_old_customer"
PROMOTION_TMALL_BRAND = "tmall_brand_enjoy"


def text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def normalize_header(value: Any) -> str:
    return re.sub(r"\s+", "", text(value)).lower()


def find_header(headers: Iterable[Any], candidates: Iterable[str]) -> str | None:
    original = [text(header) for header in headers]
    normalized = {normalize_header(header): header for header in original if header}
    for candidate in candidates:
        match = normalized.get(normalize_header(candidate))
        if match:
            return match
    for header in original:
        clean = normalize_header(header)
        if any(normalize_header(candidate) in clean for candidate in candidates):
            return header
    return None


def find_header_row(
    rows: list[list[Any]], predicate: Callable[[list[str]], bool]
) -> int | None:
    for index, row in enumerate(rows[:HEADER_SEARCH_ROWS]):
        headers = [text(value) for value in row]
        if predicate(headers):
            return index
    return None


def numbered_dict_rows(
    rows: list[list[Any]], header_row_index: int
) -> list[tuple[int, dict[str, Any]]]:
    headers = [text(value) for value in rows[header_row_index]]
    result: list[tuple[int, dict[str, Any]]] = []
    for physical_index, raw in enumerate(rows[header_row_index + 1 :], start=header_row_index + 2):
        if not any(text(value) for value in raw):
            continue
        item = {
            headers[index]: raw[index] if index < len(raw) else ""
            for index in range(len(headers))
            if headers[index]
        }
        result.append((physical_index, item))
    return result


def matching_row_headers(row: dict[str, Any], candidates: Iterable[str]) -> list[str]:
    matches: list[str] = []
    used: set[str] = set()
    for candidate in candidates:
        header = find_header(row.keys(), [candidate])
        if header and header not in used:
            matches.append(header)
            used.add(header)
    return matches


def resolve_order_date(
    row: dict[str, Any], platform: str, parse_date: Callable[[Any], str]
) -> tuple[str, dict[str, str]]:
    candidates = [
        *ORDER_PAYMENT_DATE_CANDIDATES,
        *ORDER_TRANSACTION_DATE_CANDIDATES,
    ]
    if platform == "天猫":
        candidates.extend(ORDER_CREATION_DATE_CANDIDATES)
    raw_values: dict[str, str] = {}
    for header in matching_row_headers(row, candidates):
        raw_value = text(row.get(header))
        raw_values[header] = raw_value
        parsed = parse_date(row.get(header))
        if re.fullmatch(r"20\d{2}-\d{2}-\d{2}", parsed):
            return parsed, raw_values
    return "", raw_values


def classify_promotion_headers(headers: Iterable[Any]) -> str | None:
    normalized = {normalize_header(header) for header in headers}
    if "预估新客加速费用" in normalized:
        return PROMOTION_TMALL_NEW
    if "预估老客加速费用" in normalized:
        return PROMOTION_TMALL_OLD
    if "预估抽佣金额" in normalized and "业务日期" in normalized:
        return PROMOTION_TMALL_BRAND
    if "主体id" in normalized and "花费" in normalized:
        return PROMOTION_TMALL_WANXIANG
    if (
        ("商品id" in normalized)
        and bool(normalized & {"总花费(元)", "成交花费(元)"})
    ):
        return PROMOTION_PDD
    return None


def classify_promotion_row(headers: Iterable[Any], row: dict[str, Any]) -> str | None:
    """Use the organizer's row-level type when present; fall back to raw headers."""
    archive_type_header = find_header(headers, PROMOTION_ARCHIVE_TYPE_CANDIDATES)
    if archive_type_header:
        archive_type = normalize_header(row.get(archive_type_header))
        archive_type_map = {
            "拼多多-商品推广": PROMOTION_PDD,
            "万相台": PROMOTION_TMALL_WANXIANG,
            "新客加速": PROMOTION_TMALL_NEW,
            "老客加速": PROMOTION_TMALL_OLD,
        }
        return archive_type_map.get(archive_type)
    return classify_promotion_headers(headers)


def has_promotion_amount_header(headers: Iterable[Any]) -> bool:
    normalized = {normalize_header(header) for header in headers}
    return bool(
        normalized
        & {
            "总花费(元)",
            "成交花费(元)",
            "花费",
            "消耗",
            "总消耗",
            "预估新客加速费用",
            "预估老客加速费用",
            "预估抽佣金额",
            "品牌新享费用",
        }
    )


def is_excluded_order_path(path: Path) -> bool:
    name = path.name
    lowered = name.casefold()
    return (
        name.startswith("订单汇总表_")
        or "整理记录" in name
        or "缺失提醒" in name
        or "中断" in name
        or "_合并前备份" in str(path)
        or "备份" in name
        or "临时" in name
        or name.startswith("~$")
        or lowered.endswith(".tmp")
    )


@dataclass
class ProductExportIndex:
    by_pair: dict[tuple[str, str], dict[str, str]] = field(default_factory=dict)
    by_product: dict[str, list[dict[str, str]]] = field(default_factory=dict)

    def add(self, product_id: str, sku_id: str, spec_code: str, source: str) -> None:
        product_id = text(product_id)
        sku_id = text(sku_id)
        spec_code = text(spec_code)
        if not product_id:
            return
        record = {
            "商品ID": product_id,
            "SKU ID": sku_id,
            "商品SKU": spec_code,
            "来源": source,
        }
        if sku_id:
            key = (product_id, sku_id)
            previous = self.by_pair.get(key)
            if previous and previous.get("商品SKU") and spec_code and previous["商品SKU"] != spec_code:
                raise ValueError(
                    "商品数据规格映射冲突："
                    f"商品ID={product_id}，SKU ID={sku_id}，"
                    f"前值={previous['商品SKU']}（{previous['来源']}），"
                    f"后值={spec_code}（{source}）"
                )
            if not previous or (not previous.get("商品SKU") and spec_code):
                self.by_pair[key] = record
        candidates = self.by_product.setdefault(product_id, [])
        signature = (sku_id, spec_code)
        if signature not in {
            (candidate.get("SKU ID", ""), candidate.get("商品SKU", ""))
            for candidate in candidates
        }:
            candidates.append(record)

    def resolve(self, product_id: str, sku_id: str = "") -> dict[str, str]:
        product_id = text(product_id)
        sku_id = text(sku_id)
        if sku_id and (product_id, sku_id) in self.by_pair:
            return dict(self.by_pair[(product_id, sku_id)])
        candidates = self.by_product.get(product_id, [])
        unique_pairs = {
            (candidate.get("SKU ID", ""), candidate.get("商品SKU", ""))
            for candidate in candidates
        }
        if len(unique_pairs) == 1 and candidates:
            return dict(candidates[0])
        return {}

    def sku_count(self, product_id: str) -> int:
        return len(
            {
                (candidate.get("SKU ID", ""), candidate.get("商品SKU", ""))
                for candidate in self.by_product.get(text(product_id), [])
            }
        )
