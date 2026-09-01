"""Regression checks for same-batch product-ID reminder reconciliation."""

from __future__ import annotations

import sys
from pathlib import Path


SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

import organize_reports as organizer


def analysis(
    *,
    report_type: str,
    shop_name: str,
    product_ids: list[str],
    new_product_ids: list[str] | None = None,
    status: str = "ready",
) -> organizer.Analysis:
    return organizer.Analysis(
        source=Path(f"{shop_name}-{report_type}-{len(product_ids)}.xlsx"),
        status=status,
        reason="ok" if status == "ready" else "missing_required_fields",
        report_type=report_type,
        target_folder=report_type,
        shop_name=shop_name,
        shop_source="test",
        period="2026-09-01",
        period_source="test",
        download_date="2026-09-01",
        target_path=Path("target.xlsx"),
        product_counts="",
        matched_headers="",
        missing_fields="" if status == "ready" else "period",
        map_updates=[],
        product_ids=product_ids,
        new_product_ids=list(new_product_ids or []),
    )


def test_ready_product_data_resolves_same_shop_reminders() -> None:
    product_data = analysis(
        report_type="商品数据",
        shop_name="测试店",
        product_ids=["1000001", "1000002"],
    )
    order = analysis(
        report_type="订单数据",
        shop_name="测试店",
        product_ids=["1000001", "1000003"],
        new_product_ids=["1000001", "1000003"],
    )
    promotion = analysis(
        report_type="推广数据",
        shop_name="测试店",
        product_ids=["1000002"],
        new_product_ids=["1000002"],
    )

    organizer.reconcile_same_batch_product_data_ids(
        [order, promotion, product_data]
    )

    assert order.new_product_ids == ["1000003"]
    assert promotion.new_product_ids == []


def test_pending_or_cross_shop_product_data_does_not_resolve_reminder() -> None:
    pending_product_data = analysis(
        report_type="商品数据",
        shop_name="测试店",
        product_ids=["2000001"],
        status="pending",
    )
    other_shop_product_data = analysis(
        report_type="商品数据",
        shop_name="另一店",
        product_ids=["2000002"],
    )
    order = analysis(
        report_type="订单数据",
        shop_name="测试店",
        product_ids=["2000001", "2000002"],
        new_product_ids=["2000001", "2000002"],
    )

    organizer.reconcile_same_batch_product_data_ids(
        [order, pending_product_data, other_shop_product_data]
    )

    assert order.new_product_ids == ["2000001", "2000002"]


def test_conflicting_product_data_shops_keep_reminder() -> None:
    shop_a_product_data = analysis(
        report_type="商品数据",
        shop_name="测试店",
        product_ids=["3000001"],
    )
    shop_b_product_data = analysis(
        report_type="商品数据",
        shop_name="另一店",
        product_ids=["3000001"],
    )
    order = analysis(
        report_type="订单数据",
        shop_name="测试店",
        product_ids=["3000001"],
        new_product_ids=["3000001"],
    )

    organizer.reconcile_same_batch_product_data_ids(
        [order, shop_a_product_data, shop_b_product_data]
    )

    assert order.new_product_ids == ["3000001"]


def main() -> None:
    test_ready_product_data_resolves_same_shop_reminders()
    test_pending_or_cross_shop_product_data_does_not_resolve_reminder()
    test_conflicting_product_data_shops_keep_reminder()
    print("PASS: V2同批商品ID映射提醒回归测试全部通过")


if __name__ == "__main__":
    main()
