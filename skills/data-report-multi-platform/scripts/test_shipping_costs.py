"""Freight lookup tests use saved prices and never substitute a flat rate."""
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from openpyxl import Workbook

from data_quality_protection import DataQualityError
from shipping_costs import (load_shipping_table, ShippingTable, ShippingQuote,
                            capture_shipping_order, apply_shipping_cost)
from product_cost_rules import apply_product_cost_rules


class ShippingTableTests(unittest.TestCase):
    def write(self, path, rows):
        wb = Workbook()
        wb.active.title = "编码匹配"
        wb.active["B4"] = "安徽"
        for name in ("圆通各地运费", "邮政各地运费"):
            sheet = wb.create_sheet(name)
            sheet.append(["说明"])
            sheet.append(["编码", "货品名称", "计费重(kg)", "安徽", "深圳"])
            for row in rows:
                sheet.append(row)
        wb.save(path)

    def test_normalized_code_and_specific_region_select_saved_quote(self):
        with TemporaryDirectory() as directory:
            path = Path(directory) / "运费成本.xlsx"
            self.write(path, [["ＰＡＣＫ 01", "组合装", 1, 2.65, 4.2]])
            table, _ = load_shipping_table(path)
            self.assertEqual(table.selected_region, "安徽")
            self.assertEqual(table.lookup("pack01", "圆通", "深圳").amount, 4.2)
            self.assertIn("/E3/", table.lookup("pack01", "圆通", "深圳").source)
            self.assertIsNone(table.lookup("unknown", "圆通", "安徽").amount)
            self.assertIsNone(table.lookup("pack01", "EMS", "安徽").amount)
            self.assertIsNone(table.lookup("pack01", "圆通", "").amount)

    def test_invalid_prices_and_uncached_formula_are_unknown(self):
        with TemporaryDirectory() as directory:
            path = Path(directory) / "运费成本.xlsx"
            self.write(path, [["A", "缺资料", 1, "#N/A", "=1+2"],
                              ["B", "非法价格", 1, 0, -2]])
            table, _ = load_shipping_table(path)
            for code in ("A", "B"):
                for region in ("安徽", "深圳"):
                    self.assertIsNone(table.lookup(code, "圆通", region).amount)

    def test_normalized_conflicting_duplicate_stops(self):
        with TemporaryDirectory() as directory:
            path = Path(directory) / "运费成本.xlsx"
            self.write(path, [["A", "产品", 1, 2, 3], ["Ａ", "产品", 1, 2.5, 3]])
            with self.assertRaises(DataQualityError):
                load_shipping_table(path)

    def test_same_price_duplicates_warn(self):
        with TemporaryDirectory() as directory:
            path = Path(directory) / "运费成本.xlsx"
            self.write(path, [["A", "产品", 1, 2, 3], ["Ａ", "产品", 1, 2, 3]])
            table, warnings = load_shipping_table(path)
            self.assertEqual(table.lookup("a", "邮政", "安徽").amount, 2)
            self.assertTrue(any("重复" in warning for warning in warnings))

    def test_missing_file_is_visible_and_unknown(self):
        table, warnings = load_shipping_table(Path("不存在/运费成本.xlsx"))
        self.assertIsNone(table.lookup("A", "圆通", "安徽").amount)
        self.assertIn("不存在", warnings[0])


class ShippingPolicyTests(unittest.TestCase):
    def table(self):
        prices = {}
        for region in ("黑龙江", "吉林", "辽宁", "安徽", "广东", "深圳"):
            yto, postal = ((9, 1) if region in ("黑龙江", "吉林", "辽宁")
                           else (5, 4) if region == "深圳" else (3, 2))
            for carrier, amount in (("圆通", yto), ("邮政", postal)):
                prices[("a", carrier, region)] = ShippingQuote(amount, f"{carrier}/{region}")
        return ShippingTable(prices, selected_region="安徽")

    def test_northeast_always_uses_yto_even_when_postal_is_cheaper(self):
        table = self.table()
        for province in ("黑龙江省", "吉林", "辽宁省", "黑", "吉", "辽"):
            self.assertEqual(table.select("Ａ", province).amount, 9)

    def test_other_regions_take_lower_valid_quote_and_shenzhen_has_own_price(self):
        table = self.table()
        self.assertEqual(table.select("A", "安徽省").amount, 2)
        self.assertEqual(table.select("A", "广东省", "深圳市").amount, 4)
        self.assertIsNone(table.select("A", "广东省", "").amount)
        self.assertIsNone(table.select("A", "辽宁", "深圳").amount)
        self.assertIsNone(table.select("A", "未知").amount)
        self.assertIsNone(table.select("A", "").amount)  # No selected-region fallback.

    def test_one_missing_quote_only_allows_fallback_outside_northeast(self):
        table = self.table()
        table.prices[("a", "圆通", "安徽")] = ShippingQuote(None, "圆通/安徽", "缺体积")
        self.assertEqual(table.select("A", "安徽").amount, 2)
        table.prices[("a", "圆通", "辽宁")] = ShippingQuote(None, "圆通/辽宁", "缺体积")
        self.assertIsNone(table.select("A", "辽宁").amount)
        table.prices[("a", "邮政", "安徽")] = ShippingQuote(None, "邮政/安徽", "缺体积")
        self.assertIsNone(table.select("A", "安徽").amount)

    def test_unit_shipping_uses_accepted_quantity_weighted_destinations(self):
        aggregate = {}
        capture_shipping_order(aggregate, {"省": "辽宁省", "市": "沈阳市"}, 1, "A", "订单/行2")
        capture_shipping_order(aggregate, {"省": "安徽省", "市": "合肥市"}, 3, "A", "订单/行3")
        row = {"SKU编码": "A", "实际成交数量（去退款去补单后）": 4, "到手价": 100, "产品成本": 10}
        warnings, log = [], {}
        apply_shipping_cost(row, aggregate, self.table(), warnings, log)
        apply_product_cost_rules(row)
        self.assertEqual(row["快递成本"], 3.75)
        self.assertAlmostEqual(row["毛利"], 68.85)
        self.assertIn("4/4件", log["运费匹配状态"])
        self.assertIn("邮政/安徽", log["运费来源"])
        self.assertEqual(warnings, [])

    def test_partial_match_never_averages_only_the_known_orders(self):
        aggregate = {}
        capture_shipping_order(aggregate, {"省": "安徽"}, 3, "A", "订单/行2")
        capture_shipping_order(aggregate, {}, 1, "A", "订单/行3")
        row = {"SKU编码": "A", "实际成交数量（去退款去补单后）": 4, "到手价": 100, "产品成本": 10}
        warnings, log = [], {}
        apply_shipping_cost(row, aggregate, self.table(), warnings, log)
        apply_product_cost_rules(row)
        for field in ("快递成本", "综合成本", "毛利", "净利润"):
            self.assertEqual(row[field], "")
        self.assertEqual(row["产品成本"], 10)
        self.assertIn("3/4件", log["运费匹配状态"])
        self.assertTrue(any("订单/行3" in warning for warning in warnings))


if __name__ == "__main__":
    unittest.main()
