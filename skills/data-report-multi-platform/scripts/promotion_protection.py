from __future__ import annotations

import hashlib
import re
from collections import defaultdict
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable


UNKNOWN_DIMENSION = "<未提供>"
RECONCILIATION_TOLERANCE = 0.01
DATE_TOKEN_PATTERN = re.compile(
    r"(20\d{2})[-/.年](\d{1,2})[-/.月](\d{1,2})(?:日)?"
)
GENERIC_PATH_PARTS = {
    "推广数据",
    "品牌新享数据",
    "品牌新享",
    "品牌心享",
    "品牌心想",
    "万相台",
    "新客加速",
    "老客加速",
    "天猫",
    "拼多多",
}


class PromotionProtectionError(RuntimeError):
    pass


@dataclass(frozen=True)
class PromotionSnapshotRow:
    platform: str
    shop: str
    promotion_type: str
    date: str
    product_id: str
    amount: float
    source_path: Path

    @property
    def business_key(self) -> tuple[str, str, str, str, str]:
        return (
            self.platform,
            self.shop or UNKNOWN_DIMENSION,
            self.promotion_type,
            self.date,
            self.product_id,
        )


def dates_in_text(value: Any) -> list[str]:
    dates: list[str] = []
    for year, month, day in DATE_TOKEN_PATTERN.findall(str(value or "")):
        dates.append(f"{int(year):04d}-{int(month):02d}-{int(day):02d}")
    return dates


def date_bounds_from_filename(path: Path) -> tuple[str, str]:
    dates = dates_in_text(path.name)
    if not dates:
        return "", ""
    if len(dates) == 1:
        return dates[0], dates[0]
    return dates[0], dates[1]


def require_single_day_value(value: Any, source_path: Path, field_name: str) -> None:
    dates = dates_in_text(value)
    if len(set(dates)) > 1:
        raise PromotionProtectionError(
            f"推广日期不是逐日数据，已停止生成日报：文件={source_path}，字段={field_name}，值={value}"
        )


def infer_path_dimensions(
    root: Path, promotion_dir: str, source_path: Path
) -> tuple[str, str]:
    relative_parts: tuple[str, ...] = ()
    for base in (root / promotion_dir, root):
        try:
            relative = source_path.relative_to(base)
        except ValueError:
            continue
        relative_parts = relative.parts[:-1]
        if base == root and relative_parts:
            relative_parts = relative_parts[1:]
        break

    meaningful = [
        part
        for part in relative_parts
        if part not in GENERIC_PATH_PARTS and not dates_in_text(part)
    ]
    shop = meaningful[0] if meaningful else ""
    # The archive's year/month folders are not promotion-plan identities.
    # Keep the second return value only for backward-compatible callers; the
    # report workflow never reads or matches a plan ID.
    return shop, ""


def _file_rank(path: Path) -> tuple[int, str]:
    try:
        modified = path.stat().st_mtime_ns
    except OSError:
        modified = 0
    return modified, str(path).casefold()


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def deduplicate_exact_files(
    paths: Iterable[Path], promotion_type: str
) -> tuple[list[Path], list[str]]:
    by_digest: dict[str, list[Path]] = defaultdict(list)
    unreadable: list[Path] = []
    warnings: list[str] = []
    for path in paths:
        try:
            by_digest[_sha256(path)].append(path)
        except OSError as exc:
            unreadable.append(path)
            warnings.append(f"{promotion_type}文件哈希读取失败,{path},{exc}")

    selected = list(unreadable)
    for duplicates in by_digest.values():
        newest = max(duplicates, key=_file_rank)
        selected.append(newest)
        ignored = [path for path in duplicates if path != newest]
        if ignored:
            warnings.append(
                f"{promotion_type}完全重复文件已去重,保留{newest},忽略"
                + "|".join(str(path) for path in sorted(ignored))
            )
    return sorted(selected), warnings


def select_latest_snapshots(
    rows: Iterable[PromotionSnapshotRow],
) -> tuple[dict[tuple[str, str], float], list[str]]:
    per_source: dict[
        tuple[str, str, str, str, str], dict[Path, float]
    ] = defaultdict(lambda: defaultdict(float))
    for row in rows:
        per_source[row.business_key][row.source_path] += row.amount

    costs: dict[tuple[str, str], float] = defaultdict(float)
    warnings: list[str] = []
    for business_key, source_amounts in per_source.items():
        platform, shop, promotion_type, date, product_id = business_key
        if len(source_amounts) == 1:
            selected_path, amount = next(iter(source_amounts.items()))
        else:
            if shop == UNKNOWN_DIMENSION:
                sources = "|".join(str(path) for path in sorted(source_amounts))
                raise PromotionProtectionError(
                    "推广快照无法安全去重，已停止生成日报："
                    f"平台={platform}，推广类型={promotion_type}，日期={date}，商品ID={product_id}，"
                    f"店铺={shop}，重复来源={sources}。"
                    "请补充店铺信息，或只保留一份明确的最新快照。"
                )
            selected_path = max(source_amounts, key=_file_rank)
            amount = source_amounts[selected_path]
            ignored = [path for path in source_amounts if path != selected_path]
            warnings.append(
                "推广重复快照已去重，保留最新文件,"
                f"平台={platform},店铺={shop},推广类型={promotion_type},"
                f"日期={date},商品ID={product_id},保留={selected_path},忽略="
                + "|".join(str(path) for path in sorted(ignored))
            )
        costs[(date, product_id)] += amount
    return dict(costs), warnings


def _text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def _number(value: Any) -> float:
    if value is None or value == "":
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    cleaned = (
        str(value)
        .replace(",", "")
        .replace("￥", "")
        .replace("元", "")
        .strip()
    )
    if cleaned in {"", "-", "--"}:
        return 0.0
    try:
        return float(cleaned)
    except ValueError:
        return 0.0


def reconcile_promotion_costs(
    expected_costs: dict[tuple[str, str], float],
    output_rows: Iterable[dict[str, Any]],
    platform: str,
    tolerance: float = RECONCILIATION_TOLERANCE,
) -> tuple[float, float]:
    actual_costs: dict[tuple[str, str], float] = defaultdict(float)
    for row in output_rows:
        amount = _number(row.get("推广费用"))
        if abs(amount) <= 1e-12:
            continue
        date = _text(row.get("日期"))
        product_id = _text(row.get("商品ID"))
        if not date or not product_id:
            raise PromotionProtectionError(
                f"{platform}推广费对账失败：输出行有推广费但缺少日期或商品ID，推广费={amount:.2f}。"
            )
        actual_costs[(date, product_id)] += amount

    mismatches: list[str] = []
    for date, product_id in sorted(set(expected_costs) | set(actual_costs)):
        expected = expected_costs.get((date, product_id), 0.0)
        actual = actual_costs.get((date, product_id), 0.0)
        difference = actual - expected
        if abs(difference) > tolerance + 1e-9:
            mismatches.append(
                f"日期={date}，商品ID={product_id}，导入={expected:.2f}，输出={actual:.2f}，差额={difference:.2f}"
            )

    expected_total = sum(expected_costs.values())
    actual_total = sum(actual_costs.values())
    if abs(actual_total - expected_total) > tolerance + 1e-9:
        mismatches.append(
            f"总计，导入={expected_total:.2f}，输出={actual_total:.2f}，差额={actual_total - expected_total:.2f}"
        )
    if mismatches:
        preview = "\n".join(f"- {item}" for item in mismatches[:20])
        remainder = "" if len(mismatches) <= 20 else f"\n- 另有{len(mismatches) - 20}项。"
        raise PromotionProtectionError(
            f"{platform}推广费输入输出对账失败，差额超过{tolerance:.2f}元，已停止生成日报。\n"
            f"{preview}{remainder}"
        )
    return expected_total, actual_total
