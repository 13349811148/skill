"""Regression checks for PDD marketing fallback and shop-scoped promotion spend."""

from __future__ import annotations

import sys
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch


SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

import build_pdd_report as pdd
from data_quality_protection import normalize_match_code
from promotion_protection import (
    PromotionProtectionError,
    PromotionSnapshotRow,
    reconcile_promotion_costs,
    select_latest_snapshots,
)


def sheet(headers: list[str], rows: list[list[object]], name: str = "明细") -> SimpleNamespace:
    return SimpleNamespace(
        rows=[headers, *rows],
        sheet_name=name,
        truncated=False,
        total_rows=len(rows) + 1,
    )


def marketing_row(
    *,
    shop: str,
    product_id: str,
    style_id: str,
    sku: str,
    cost: object = 20,
    activity_price: object = 25,
    registration_price: object = 30,
) -> dict[str, object]:
    return {
        "品牌": "测试品牌",
        "所属运营": "测试运营",
        "平台": "拼多多",
        "店铺名称": shop,
        "大类": "测试大类",
        "商品ID": product_id,
        "样式ID": style_id,
        "商品SKU": sku,
        "产品成本": cost,
        "活动价": activity_price,
        "报名价": registration_price,
        "活动名称": "测试活动",
        pdd.EXPECTED_UNIT_RECEIPT_COLUMN: 50,
    }


def test_changed_product_id_uses_shop_and_normalized_sku_without_overwriting_order() -> None:
    rows = [
        marketing_row(
            shop="甲店",
            product_id="旧商品",
            style_id="旧样式",
            sku="Ab １２",
            cost=22,
        )
    ]
    resolution = pdd.resolve_marketing_fields(
        rows,
        style_id="新样式",
        product_id="新商品",
        shop="甲店",
        sku="Ａ Ｂ12",
    )
    assert resolution.values["品牌"] == "测试品牌"
    assert resolution.values["产品成本"] == 22
    assert resolution.sources["产品成本"].startswith("店铺名称+规范化商品SKU=")

    enriched = pdd.enrich_row(
        {
            "日期": "2026-09-01",
            "店铺名称": "甲店",
            "商品ID": "新商品",
            "样式ID": "新样式",
            "商品SKU": "Ａ Ｂ12",
        },
        pdd.load_marketing_products(rows),
        pdd.load_marketing_styles(rows),
        pdd.ProductExportIndex(),
        rows,
    )
    assert enriched["商品ID"] == "新商品"
    assert enriched["店铺名称"] == "甲店"
    assert enriched["商品SKU"] == "Ａ Ｂ12"


def test_cross_shop_conflicts_are_field_specific_and_never_take_first_row() -> None:
    rows = [
        marketing_row(
            shop="甲店",
            product_id="商品甲",
            style_id="样式甲",
            sku="same sku",
            cost=20,
            activity_price=10,
        ),
        marketing_row(
            shop="乙店",
            product_id="商品乙",
            style_id="样式乙",
            sku="ＳＡＭＥＳＫＵ",
            cost=20,
            activity_price=11,
        ),
    ]
    resolution = pdd.resolve_marketing_fields(
        rows,
        style_id="未知样式",
        product_id="新商品",
        shop="丙店",
        sku="Same Sku",
    )
    assert resolution.values["产品成本"] == 20
    assert resolution.sources["产品成本"].startswith("商品SKU跨店逐字段一致值=")
    assert resolution.values["活动价"] == ""
    assert any("字段=活动价" in warning for warning in resolution.conflicts)


def test_same_shop_marketing_cost_overrides_old_cost_table_and_logs_difference() -> None:
    rows = [
        marketing_row(
            shop="甲店",
            product_id="旧商品",
            style_id="旧样式",
            sku="SKU-1",
            cost=22,
        )
    ]
    resolution = pdd.resolve_marketing_fields(
        rows,
        style_id="新样式",
        product_id="新商品",
        shop="甲店",
        sku="sku － 1",
    )
    cost_table = {
        normalize_match_code("SKU-1"): {
            "产品成本": 18,
            "来源": "成本表:成本表.xlsx/成本/商品SKU=SKU-1",
        }
    }
    cost = pdd.resolve_pdd_cost(
        {"商品SKU": "ｓｋｕ－１"},
        resolution,
        cost_table,
        Path("营销活动.xlsx"),
    )
    assert cost.cost == 22
    assert cost.source.startswith("营销活动表:营销活动.xlsx/店铺名称+")
    assert cost.marketing_cost == 22
    assert cost.table_cost == 18
    assert cost.cost_difference == 4


def test_full_report_logs_marketing_cost_override() -> None:
    marketing_rows = [
        marketing_row(
            shop="甲店",
            product_id="旧商品",
            style_id="旧样式",
            sku="SKU-1",
            cost=22,
        )
    ]
    aggregates = {
        ("2026-09-01", "甲店", "新商品", "新样式"): {
            "日期": "2026-09-01",
            "店铺名称": "甲店",
            "商品ID": "新商品",
            "样式ID": "新样式",
            "商品SKU": "ｓｋｕ－１",
            "实际成交数量（去退款去补单后）": 1,
            "实际成交金额（去退款去补单后）": 50,
        }
    }
    cost_table = {
        normalize_match_code("SKU-1"): {
            "产品成本": 18,
            "来源": "成本表:成本表.xlsx/成本/商品SKU=SKU-1",
        }
    }
    columns = [
        "日期",
        "品牌",
        "所属运营",
        "平台",
        "店铺名称",
        "大类",
        "商品ID",
        "商品SKU",
        "产品成本",
        "活动价",
        "报名价",
        pdd.PROMOTION_MECHANISM_COLUMN,
        "到手价",
        "实际成交数量（去退款去补单后）",
        "实际成交金额（去退款去补单后）",
        "推广费用",
        "SKU ID",
    ]
    with (
        patch.object(pdd, "read_template_columns", return_value=columns),
        patch.object(pdd, "load_marketing_rows", return_value=marketing_rows),
        patch.object(pdd, "load_costs_by_merchant_code", return_value=(cost_table, [])),
        patch.object(pdd, "load_product_exports", return_value=(pdd.ProductExportIndex(), [])),
        patch.object(pdd, "load_order_aggregates", return_value=(aggregates, [])),
        patch.object(pdd, "load_promo_costs", return_value=({}, [])),
    ):
        _columns, rows, warnings, logs = pdd.build_report(
            Path("数据库"),
            Path("营销活动.xlsx"),
            {"2026-09-01"},
            Path("成本表.xlsx"),
            "include",
        )
    assert rows[0]["产品成本"] == 22
    assert rows[0]["商品ID"] == "新商品"
    assert rows[0]["店铺名称"] == "甲店"
    assert rows[0]["商品SKU"] == "ｓｋｕ－１"
    assert logs[0]["营销成本来源"].startswith("营销活动表:营销活动.xlsx/店铺名称+")
    assert logs[0]["营销成本"] == "22"
    assert logs[0]["成本表成本"] == "18"
    assert logs[0]["成本差异（营销-成本表）"] == "4"
    assert any("营销成本覆盖成本表旧成本" in warning for warning in warnings)


def test_empty_burn_multi_spec_and_invalid_numbers_are_visible() -> None:
    rows = [
        marketing_row(shop="甲店", product_id="1001", style_id="S1", sku="SKU-1"),
        marketing_row(shop="甲店", product_id="1001", style_id="S2", sku="SKU-2"),
    ]
    empty = pdd.build_empty_burn_promotion_row(
        "2026-09-01",
        "甲店",
        "1001",
        12.5,
        pdd.load_marketing_products(rows),
        pdd.load_marketing_styles(rows),
        pdd.ProductExportIndex(),
        rows,
    )
    empty["产品成本"] = "#N/A"
    warnings = pdd.finalize_marketing_display_fields(empty, empty_burn=True)
    assert empty["店铺名称"] == "甲店"
    assert empty["商品SKU"] == pdd.MULTI_SPEC_VALUE
    assert empty["SKU ID"] == pdd.MULTI_SPEC_VALUE
    assert empty["到手价"] == pdd.EMPTY_BURN_ARRIVAL_PRICE
    assert empty["产品成本"] == pdd.MISSING_MARKETING_VALUE
    assert any("#N/A" in warning and "未转换为0" in warning for warning in warnings)


def test_arrival_price_always_uses_actual_even_when_more_than_five_percent_low() -> None:
    almost_equal_arrival, no_warning = pdd.resolve_arrival_price(
        50,
        actual_amount=99.998,
        quantity=2,
        date="2026-09-01",
        style_id="S1",
        product_id="1001",
    )
    assert almost_equal_arrival == 49.999
    assert no_warning is None

    arrival, warning = pdd.resolve_arrival_price(
        50,
        actual_amount=80,
        quantity=2,
        date="2026-09-01",
        style_id="S1",
        product_id="1001",
    )
    assert arrival == 40
    assert warning and "超过5%" in warning and "已按实际到手价填写" in warning

    invalid_arrival, invalid_warning = pdd.resolve_arrival_price(
        "#N/A",
        actual_amount=80,
        quantity=2,
        date="2026-09-01",
        style_id="S1",
        product_id="1001",
    )
    assert invalid_arrival == 40
    assert invalid_warning and "#N/A" in invalid_warning


def test_promotion_shop_dimension_is_preserved_and_reconciled() -> None:
    snapshots = [
        PromotionSnapshotRow("拼多多", "甲店", "拼多多推广费", "2026-09-01", "1001", 8, Path("甲.xlsx")),
        PromotionSnapshotRow("拼多多", "乙店", "拼多多推广费", "2026-09-01", "1001", 9, Path("乙.xlsx")),
    ]
    costs, _warnings = select_latest_snapshots(snapshots, include_shop=True)
    assert costs == {
        ("2026-09-01", "甲店", "1001"): 8,
        ("2026-09-01", "乙店", "1001"): 9,
    }
    output = [
        {"日期": "2026-09-01", "店铺名称": "甲店", "商品ID": "1001", "推广费用": 8},
        {"日期": "2026-09-01", "店铺名称": "乙店", "商品ID": "1001", "推广费用": 9},
    ]
    reconcile_promotion_costs(
        costs,
        output,
        "拼多多",
        key_fields=("日期", "店铺名称", "商品ID"),
    )
    wrong_shop_output = [dict(output[0]), {**output[1], "店铺名称": "甲店"}]
    try:
        reconcile_promotion_costs(
            costs,
            wrong_shop_output,
            "拼多多",
            key_fields=("日期", "店铺名称", "商品ID"),
        )
    except PromotionProtectionError as exc:
        assert "店铺名称=乙店" in str(exc)
    else:
        raise AssertionError("拼多多推广费跨店错配没有被店铺维度对账拦截")


def test_pdd_promotion_loader_keeps_row_or_source_path_shop() -> None:
    root = Path("数据库")
    shop_a = root / "推广数据" / "甲店" / "甲.xlsx"
    shop_b = root / "推广数据" / "乙店" / "乙.xlsx"
    headers = ["日期", "店铺名称", "商品ID", "总花费(元)", "净实际投产比"]
    sheets = {
        shop_a: [sheet(headers, [["2026-09-01", "甲店", "1001", 8, 2]])],
        shop_b: [sheet(headers, [["2026-09-01", "", "1001", 9, 2]])],
    }
    with (
        patch.object(pdd, "discover_files", return_value=[shop_a, shop_b]),
        patch.object(
            pdd,
            "deduplicate_exact_files",
            return_value=([shop_a, shop_b], []),
        ),
        patch.object(pdd, "read_workbook_sheets", side_effect=lambda path: sheets[path]),
    ):
        costs, _warnings = pdd.load_promo_costs(root, {"2026-09-01"})
    assert costs == {
        ("2026-09-01", "甲店", "1001"): 8,
        ("2026-09-01", "乙店", "1001"): 9,
    }


def test_full_report_keeps_same_product_promotion_spend_in_source_shop() -> None:
    marketing_rows = [
        marketing_row(shop="甲店", product_id="1001", style_id="SA", sku="SKU-A", cost=20),
        marketing_row(shop="乙店", product_id="1001", style_id="SB", sku="SKU-B", cost=21),
    ]
    costs = {
        normalize_match_code("SKU-A"): {
            "产品成本": 19,
            "来源": "成本表:成本表.xlsx/成本/商品SKU=SKU-A",
        },
        normalize_match_code("SKU-B"): {
            "产品成本": 20,
            "来源": "成本表:成本表.xlsx/成本/商品SKU=SKU-B",
        },
    }
    aggregates = {
        ("2026-09-01", "甲店", "1001", "SA"): {
            "日期": "2026-09-01",
            "店铺名称": "甲店",
            "商品ID": "1001",
            "样式ID": "SA",
            "商品SKU": "SKU-A",
            "实际成交数量（去退款去补单后）": 2,
            "实际成交金额（去退款去补单后）": 100,
        }
    }
    columns = [
        "日期",
        "品牌",
        "所属运营",
        "平台",
        "店铺名称",
        "大类",
        "商品ID",
        "商品SKU",
        "产品成本",
        "活动价",
        "报名价",
        pdd.PROMOTION_MECHANISM_COLUMN,
        "到手价",
        "实际成交数量（去退款去补单后）",
        "实际成交金额（去退款去补单后）",
        "推广费用",
        "备注",
        "SKU ID",
    ]
    with (
        patch.object(pdd, "read_template_columns", return_value=columns),
        patch.object(pdd, "load_marketing_rows", return_value=marketing_rows),
        patch.object(pdd, "load_costs_by_merchant_code", return_value=(costs, [])),
        patch.object(pdd, "load_product_exports", return_value=(pdd.ProductExportIndex(), [])),
        patch.object(pdd, "load_order_aggregates", return_value=(aggregates, [])),
        patch.object(
            pdd,
            "load_promo_costs",
            return_value=({("2026-09-01", "乙店", "1001"): 9}, []),
        ),
    ):
        _columns, output_rows, warnings, logs = pdd.build_report(
            Path("数据库"),
            Path("营销活动.xlsx"),
            {"2026-09-01"},
            Path("成本表.xlsx"),
            "include",
        )
    assert len(output_rows) == 2
    sales_row, empty_burn_row = output_rows
    assert sales_row["店铺名称"] == "甲店" and sales_row["推广费用"] == ""
    assert empty_burn_row["店铺名称"] == "乙店" and empty_burn_row["推广费用"] == 9
    assert empty_burn_row["到手价"] == pdd.EMPTY_BURN_ARRIVAL_PRICE
    assert any("乙店" in warning and "空烧推广费" in warning for warning in warnings)
    assert all(log["店铺名称"] in {"甲店", "乙店"} for log in logs)


def test_cancelled_and_refunded_orders_remain_filtered() -> None:
    assert not pdd.is_valid_order({"订单状态": "已取消", "售后状态": ""})
    assert not pdd.is_valid_order({"订单状态": "已发货", "售后状态": "退款成功"})
    assert pdd.is_valid_order({"订单状态": "已发货", "售后状态": ""})

    headers = [
        "订单号",
        "订单支付时间",
        "订单状态",
        "售后状态",
        "汇总_店铺名称",
        "商品id",
        "样式ID",
        "商家编码-规格维度",
        "商品数量(件)",
        "商家实收金额(元)",
    ]
    rows = [
        ["正常", "2026-09-01", "已发货", "", "甲店", "1001", "S1", "SKU-1", 1, 50],
        ["取消", "2026-09-01", "已取消", "", "甲店", "1001", "S1", "SKU-1", 1, 50],
        ["退款", "2026-09-01", "已发货", "退款成功", "甲店", "1001", "S1", "SKU-1", 1, 50],
    ]
    source = Path("订单.xlsx")
    with (
        patch.object(pdd, "discover_files", return_value=[source]),
        patch.object(pdd, "deduplicate_exact_files", return_value=([source], [])),
        patch.object(pdd, "read_workbook_sheets", return_value=[sheet(headers, rows)]),
    ):
        aggregates, _warnings = pdd.load_order_aggregates(
            Path("数据库"),
            {"2026-09-01"},
            small_receipt_action="include",
            cost_by_merchant_code={
                normalize_match_code("SKU-1"): {"产品成本": 10, "来源": "测试成本"}
            },
        )
    record = aggregates[("2026-09-01", "甲店", "1001", "S1")]
    assert record["实际成交数量（去退款去补单后）"] == 1
    assert record["实际成交金额（去退款去补单后）"] == 50


def main() -> None:
    test_changed_product_id_uses_shop_and_normalized_sku_without_overwriting_order()
    test_cross_shop_conflicts_are_field_specific_and_never_take_first_row()
    test_same_shop_marketing_cost_overrides_old_cost_table_and_logs_difference()
    test_full_report_logs_marketing_cost_override()
    test_empty_burn_multi_spec_and_invalid_numbers_are_visible()
    test_arrival_price_always_uses_actual_even_when_more_than_five_percent_low()
    test_promotion_shop_dimension_is_preserved_and_reconciled()
    test_pdd_promotion_loader_keeps_row_or_source_path_shop()
    test_full_report_keeps_same_product_promotion_spend_in_source_shop()
    test_cancelled_and_refunded_orders_remain_filtered()
    print("PASS: 拼多多营销匹配、到手价、推广店铺隔离与订单过滤回归测试全部通过")


if __name__ == "__main__":
    main()
