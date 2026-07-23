"""Shared order-line sales classification for safe business reporting."""

from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass
from typing import Any, Mapping


ORDER_NATURE_HEADER = "订单性质"
SALES_INCLUDED_HEADER = "是否计销售"
ORDER_SALES_MARK_HEADERS = (ORDER_NATURE_HEADER, SALES_INCLUDED_HEADER)
PRODUCT_TITLE_HEADERS = (
    "商品标题",
    "商品名称",
    "商品名",
    "宝贝标题",
    "宝贝名称",
)
SUPPLEMENTARY_PRICE_KEYWORDS = ("补差价", "补收差价")


def text(value: Any) -> str:
    """Return a string without changing identifiers such as product IDs."""
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def normalized_header(value: Any) -> str:
    return re.sub(r"\s+", "", unicodedata.normalize("NFKC", text(value))).casefold()


def matching_header(row: Mapping[str, Any], candidates: tuple[str, ...]) -> str | None:
    normalized = {normalized_header(header): header for header in row if text(header)}
    for candidate in candidates:
        match = normalized.get(normalized_header(candidate))
        if match:
            return match
    for header in row:
        header_text = normalized_header(header)
        if any(normalized_header(candidate) in header_text for candidate in candidates):
            return header
    return None


def standardize_product_title(value: Any) -> str:
    """Normalize title spacing/width and remove bracketed marketing prompts."""
    normalized = unicodedata.normalize("NFKC", text(value))
    previous = None
    while normalized != previous:
        previous = normalized
        normalized = re.sub(r"(?:\[[^\[\]]*\]|【[^【】]*】|〔[^〔〕]*〕)", "", normalized)
    return re.sub(r"\s+", "", normalized).strip()


@dataclass(frozen=True)
class OrderSalesClassification:
    sales_excluded: bool
    is_supplementary_price: bool
    reason: str
    title_header: str = ""
    original_title: str = ""
    normalized_title: str = ""


def classify_order_line(row: Mapping[str, Any]) -> OrderSalesClassification:
    """Classify one order detail only; never infer from ID, money, quantity, or SKU."""
    nature_header = matching_header(row, (ORDER_NATURE_HEADER,))
    sales_header = matching_header(row, (SALES_INCLUDED_HEADER,))
    nature = text(row.get(nature_header)) if nature_header else ""
    sales_flag = text(row.get(sales_header)) if sales_header else ""

    for header in PRODUCT_TITLE_HEADERS:
        matched_header = matching_header(row, (header,))
        if not matched_header:
            continue
        original_title = text(row.get(matched_header))
        normalized_title = standardize_product_title(original_title)
        if any(keyword in normalized_title for keyword in SUPPLEMENTARY_PRICE_KEYWORDS):
            return OrderSalesClassification(
                sales_excluded=True,
                is_supplementary_price=True,
                reason="已识别为补差价，不计销售",
                title_header=matched_header,
                original_title=original_title,
                normalized_title=normalized_title,
            )

    normalized_nature = unicodedata.normalize("NFKC", nature)
    if any(keyword in normalized_nature for keyword in SUPPLEMENTARY_PRICE_KEYWORDS):
        return OrderSalesClassification(
            sales_excluded=True,
            is_supplementary_price=True,
            reason="已识别为补差价，不计销售",
        )

    if unicodedata.normalize("NFKC", sales_flag).casefold() in {"否", "no", "n", "false", "0"}:
        return OrderSalesClassification(
            sales_excluded=True,
            is_supplementary_price=False,
            reason="订单已标记为不计销售",
        )

    return OrderSalesClassification(False, False, "")


def apply_order_sales_marks(record: dict[str, str]) -> OrderSalesClassification:
    """Write stable archive marks while preserving non-sales marks already supplied upstream."""
    classification = classify_order_line(record)
    nature_header = matching_header(record, (ORDER_NATURE_HEADER,)) or ORDER_NATURE_HEADER
    sales_header = matching_header(record, (SALES_INCLUDED_HEADER,)) or SALES_INCLUDED_HEADER
    if classification.is_supplementary_price:
        record[nature_header] = "补差价"
        record[sales_header] = "否"
    elif classification.sales_excluded:
        if not text(record.get(sales_header)):
            record[sales_header] = "否"
    else:
        if not text(record.get(nature_header)):
            record[nature_header] = "正常"
        if not text(record.get(sales_header)):
            record[sales_header] = "是"
    return classification
