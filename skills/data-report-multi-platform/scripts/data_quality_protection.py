from __future__ import annotations

import math
import re
import unicodedata
from collections import Counter
from typing import Any, Iterable


COST_TABLE_MAX_ROWS = 10_000
COST_SOURCE_LOG_PREFIX = "成本来源,"
COST_TABLE_CODE_CANDIDATES = (
    "商品编码",
    "商品SKU名称",
    "商品SKU",
    "SKU名称",
    "SKU编码",
)
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
    date: str,
    shop: str,
    product_id: str,
    style_id: str,
    order_spec_code: str,
    cost: Any,
    source: str,
) -> str:
    return (
        f"{COST_SOURCE_LOG_PREFIX}平台={platform},日期={date or '未提供'},"
        f"店铺={shop or '未提供'},商品ID={product_id or '未提供'},"
        f"样式ID={style_id or '未提供'},订单规格编码={order_spec_code or '未提供'},"
        f"成本={text(cost) or '未匹配'},来源={source}"
    )
