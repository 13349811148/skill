"""Regression checks for shared cost-table matching in PDD and Tmall reports.

Run with the bundled Python from this directory, for example:
    python -X utf8 scripts/test_cost_matching_regression.py
"""

from __future__ import annotations

import sys
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch


SCRIPT_DIR = Path(__file__).resolve().parent
if str(SCRIPT_DIR) not in sys.path:
    sys.path.insert(0, str(SCRIPT_DIR))

import build_pdd_report as pdd
import build_tmall_report as tmall
from data_quality_protection import (
    COST_TABLE_CODE_CANDIDATES,
    DataQualityError,
    cost_source_log,
    load_cost_table_records,
    resolve_final_cost,
)


def sheet(headers: list[str], rows: list[list[object]], name: str = "成本") -> SimpleNamespace:
    return SimpleNamespace(
        rows=[headers, *rows],
        sheet_name=name,
        truncated=False,
        total_rows=len(rows) + 1,
    )


def marketing_data() -> tuple[dict[tuple[str, str], dict[str, object]], dict[str, dict[str, object]]]:
    record = {"商品ID": "1001", "商品SKU": "测试商品5kg", "产品成本": 28}
    return {("1001", "测试商品5kg"): record}, {"style-wrong": record}


def assert_platform_cost_resolvers(costs: dict[str, dict[str, object]]) -> None:
    products, styles = marketing_data()
    for name, module in (("拼多多", pdd), ("天猫", tmall)):
        resolved = module.resolve_order_unit_cost(
            "style-wrong", "1001", "测试商品2.5kg", "测试商品2.5kg", products, styles, costs
        )
        assert resolved == 15.8, f"{name}未优先使用成本表最终SKU，得到={resolved}"


def test_header_alias_and_specification_conflict() -> None:
    costs, warnings = load_cost_table_records(
        Path("成本表.xlsx"),
        [sheet(["商家编码-规格维度", "成本价"], [["测试商品2.5kg", 15.8]])],
    )
    assert not warnings
    record = costs["测试商品2.5kg"]
    assert record["产品成本"] == 15.8
    assert record["来源"] == "成本表:成本表.xlsx/成本/商家编码-规格维度=测试商品2.5kg"

    products, styles = marketing_data()
    resolution = resolve_final_cost(
        style_id="style-wrong",
        product_id="1001",
        final_order_sku="测试 商品 ２．５ＫＧ",
        cost_by_merchant_code=costs,
        template_products=products,
        template_styles=styles,
        marketing_source_prefix="营销活动表:营销活动.xlsx",
    )
    assert resolution.cost == 15.8
    assert resolution.source.startswith("成本表:")
    assert "订单SKU与营销表样式SKU冲突" in resolution.specification_conflict_warning
    assert "商品ID=1001" in resolution.specification_conflict_warning
    assert "样式ID=style-wrong" in resolution.specification_conflict_warning
    assert "订单商品SKU=测试 商品 ２．５ＫＧ" in resolution.specification_conflict_warning
    assert "营销表商品SKU=测试商品5kg" in resolution.specification_conflict_warning
    assert "成本表成本=15.80" in resolution.specification_conflict_warning
    assert "营销表成本=28.00" in resolution.specification_conflict_warning
    assert_platform_cost_resolvers(costs)

    log = cost_source_log(
        platform="拼多多",
        row_type="正常销售",
        date="2026-07-24",
        shop="测试店",
        product_id="1001",
        style_id="style-wrong",
        order_spec_code="测试商品2.5kg",
        cost=resolution.cost,
        source=resolution.source,
    )
    assert log["成本来源"] == record["来源"]


def test_product_sku_header_and_platform_loaders() -> None:
    source_sheet = sheet(["商品SKU", "产品成本"], [["测试商品2.5kg", 15.8]])
    costs, _warnings = load_cost_table_records(Path("成本表.xlsx"), [source_sheet])
    assert costs["测试商品2.5kg"]["产品成本"] == 15.8
    assert costs["测试商品2.5kg"]["来源"].endswith("/商品SKU=测试商品2.5kg")

    for name, module in (("拼多多", pdd), ("天猫", tmall)):
        original_reader = module.read_all_workbook_sheets
        try:
            module.read_all_workbook_sheets = lambda _path: [source_sheet]
            loaded, _warnings = module.load_costs_by_merchant_code(Path("成本表.xlsx"))
        finally:
            module.read_all_workbook_sheets = original_reader
        assert loaded["测试商品2.5kg"]["产品成本"] == 15.8, f"{name}未使用共享成本读取器"


def test_per_row_fallback_and_duplicate_protection() -> None:
    costs, _warnings = load_cost_table_records(
        Path("成本表.xlsx"),
        [
            sheet(
                ["商家编码-规格维度", "商品SKU", "成本"],
                [["", "测试商品2.5kg", 15.8], ["测试商品3kg", "", 20]],
            )
        ],
    )
    assert costs["测试商品2.5kg"]["来源"].endswith("/商品SKU=测试商品2.5kg")
    assert costs["测试商品3kg"]["来源"].endswith("/商家编码-规格维度=测试商品3kg")

    _costs, duplicate_warnings = load_cost_table_records(
        Path("成本表.xlsx"),
        [sheet(["商品SKU", "成本"], [["测试商品2.5kg", 15.8], ["测试 商品 2.5KG", 15.8]])],
    )
    assert any("重复且成本相同" in warning for warning in duplicate_warnings)

    try:
        load_cost_table_records(
            Path("成本表.xlsx"),
            [sheet(["商品SKU", "成本"], [["测试商品2.5kg", 15.8], ["测试 商品 2.5KG", 16]])],
        )
    except DataQualityError as exc:
        message = str(exc)
        assert "成本冲突" in message and "行号=2" in message and "行号=3" in message
    else:
        raise AssertionError("同一规范化编码的不同成本没有停止生成")


def test_marketing_fallback_only_when_cost_table_misses() -> None:
    products, styles = marketing_data()
    resolution = resolve_final_cost(
        style_id="style-wrong",
        product_id="1001",
        final_order_sku="成本表不存在的SKU",
        cost_by_merchant_code={},
        template_products=products,
        template_styles=styles,
        marketing_source_prefix="营销活动表:营销活动.xlsx",
    )
    assert resolution.cost == 28
    assert resolution.source.endswith("/样式ID=style-wrong")
    assert not resolution.specification_conflict_warning


def test_full_report_cost_regression_for_both_platforms() -> None:
    """Exercise each build_report pipeline without Excel/file I/O dependencies."""
    cost_rows, _warnings = load_cost_table_records(
        Path("成本表.xlsx"),
        [sheet(["商家编码-规格维度", "成本"], [["测试商品2.5kg", 15.8]])],
    )
    marketing_rows = [
        {
            "商品ID": "1001",
            "样式ID": "style-wrong",
            "商品SKU": "测试商品5kg",
            "产品成本": 28,
            "店铺名称": "测试店",
        }
    ]
    aggregate = {
        ("2026-07-24", "style-wrong"): {
            "日期": "2026-07-24",
            "商品ID": "1001",
            "样式ID": "style-wrong",
            "商品SKU": "测试商品2.5kg",
            "店铺名称": "测试店",
            "实际成交数量（去退款去补单后）": 2,
            "实际成交金额（去退款去补单后）": 200,
        }
    }
    for name, module in (("拼多多", pdd), ("天猫", tmall)):
        with (
            patch.object(
                module,
                "read_template_columns",
                return_value=[
                    "商品ID",
                    "SKU ID",
                    "商品SKU",
                    "产品成本",
                    "推广费用",
                    "实际成交数量（去退款去补单后）",
                    "实际成交金额（去退款去补单后）",
                ],
            ),
            patch.object(module, "load_marketing_rows", return_value=marketing_rows),
            patch.object(module, "load_costs_by_merchant_code", return_value=(cost_rows, [])),
            patch.object(module, "load_product_exports", return_value=(module.ProductExportIndex(), [])),
            patch.object(module, "load_order_aggregates", return_value=(aggregate, [])),
            patch.object(module, "load_promo_costs", return_value=({}, [])),
        ):
            if module is pdd:
                _columns, rows, warnings, logs = module.build_report(
                    Path("数据库"), Path("营销活动.xlsx"), {"2026-07-24"}, Path("成本表.xlsx"), "include"
                )
            else:
                _columns, rows, warnings, logs = module.build_report(
                    Path("数据库"),
                    Path("营销活动.xlsx"),
                    {"2026-07-24"},
                    Path("成本表.xlsx"),
                    "订单数据",
                    "推广数据",
                    "品牌新享数据",
                    "include",
                )
        assert len(rows) == 1, f"{name}回归测试没有生成预期订单行"
        assert rows[0]["产品成本"] == 15.8, f"{name}整表生成未保留成本表成本"
        assert rows[0]["实际成交数量（去退款去补单后）"] == 2
        assert rows[0]["实际成交金额（去退款去补单后）"] == 200
        assert rows[0]["推广费用"] == ""
        assert logs[0]["成本来源"].startswith("成本表:成本表.xlsx/成本/商家编码-规格维度=")
        assert any("订单SKU与营销表样式SKU冲突" in warning for warning in warnings)


def main() -> None:
    assert "商家编码-规格维度" in COST_TABLE_CODE_CANDIDATES
    assert "商品SKU" in COST_TABLE_CODE_CANDIDATES
    test_header_alias_and_specification_conflict()
    test_product_sku_header_and_platform_loaders()
    test_per_row_fallback_and_duplicate_protection()
    test_marketing_fallback_only_when_cost_table_misses()
    test_full_report_cost_regression_for_both_platforms()
    print("PASS: 拼多多与天猫共享成本匹配回归测试全部通过")


if __name__ == "__main__":
    main()
