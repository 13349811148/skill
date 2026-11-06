"""Check both builders keep their order gates before freight aggregation."""
import csv
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch

from openpyxl import Workbook

import build_pdd_report as pdd
import build_tmall_report as tmall
from portable_reports import read_workbook_sheets


class ShippingIntegrationTests(unittest.TestCase):
    def fixture(self, root, module, missing_destination=False):
        product = root / "1.产品数据库"
        product.mkdir()
        freight = Workbook()
        freight.active.title = "编码匹配"
        freight.active["B4"] = "安徽"
        for name, prices in (("圆通各地运费", (9, 3)), ("邮政各地运费", (1, 2))):
            sheet = freight.create_sheet(name)
            sheet.append(["编码", "货品名称", "计费重(kg)", "辽宁", "安徽"])
            sheet.append(["A", "产品", 1, *prices])
        freight.save(product / "运费成本.xlsx")
        costs = Workbook()
        costs.active.append(["SKU编码", "产品成本", "商品名称"])
        costs.active.append(["A", 10, "测试产品"])
        cost_path = product / "产品成本.xlsx"
        costs.save(cost_path)
        marketing = root / "营销活动.xlsx"
        wb = Workbook()
        wb.active.append(["商品ID", "样式ID", "商品SKU", "店铺名称", "产品成本",
                          "品牌", "所属运营", "平台", "大类", "活动价", "报名价", "活动名称"])
        wb.active.append(["P1", "S1", "A", "测试店", 10, "测试品牌", "测试运营",
                          "拼多多" if module is pdd else "天猫", "测试类", 100, 100, "普通活动"])
        wb.save(marketing)
        orders = root / "订单数据"
        orders.mkdir()
        style_header = "样式ID" if module is pdd else "SKU ID"
        product_header = "商品ID" if module is pdd else "宝贝ID"
        headers = ["订单号", "订单支付时间", style_header, "汇总_店铺名称", product_header,
                   "商品数量(件)", "商家实收金额(元)", "商家编码-规格维度", "订单状态",
                   "售后状态", "商品名称", "省", "市"]
        def order(order_id, quantity, receipt, province, title="测试产品", after_sale=""):
            return [order_id, "2026-11-05 12:00:00", "S1", "测试店", "P1", quantity,
                    receipt, "A", "已发货，待收货", after_sale, title, province, ""]
        rows = [order("ONE", 1, 100, "辽宁省"), order("TWO", 3, 300, "安徽省"),
                order("ONE", 1, 100, "辽宁省"),  # Identified duplicate must not add freight.
                order("TINY", 99, 0.01, "无报价地区"),
                order("SUPPLEMENT", 100, 1000, "无报价地区", title="补差价"),
                order("REFUNDED", 100, 1000, "无报价地区", after_sale="退款成功")]
        if missing_destination:
            rows.append(order("MISSING", 1, 100, ""))
        with (orders / "合成订单.csv").open("w", encoding="utf-8-sig", newline="") as file:
            writer = csv.writer(file)
            writer.writerow(headers)
            writer.writerows(rows)
        return marketing, cost_path

    def build(self, root, module, missing_destination=False):
        marketing, cost_path = self.fixture(root, module, missing_destination)
        with patch.object(module, "read_workbook_sheets", read_workbook_sheets):
            return module.build_report(root, marketing, {"2026-11-05"}, cost_path,
                                       small_receipt_action="exclude")

    def test_both_platforms_use_only_accepted_order_quantities(self):
        for module in (pdd, tmall):
            with self.subTest(platform=module.__name__), TemporaryDirectory() as directory:
                columns, rows, warnings, logs = self.build(Path(directory), module)
                self.assertEqual(len(rows), 1)
                row = rows[0]
                self.assertEqual(len(columns), 32)
                self.assertFalse(set(row)&{'活动价','报名价','每单补贴金额','总补贴金额','净销售额（实际成交+总补贴金额）'})
                self.assertEqual(row["实际成交数量（去退款去补单后）"], 4)
                self.assertEqual(row["实际成交金额（去退款去补单后）"], 400)
                self.assertEqual(row["快递成本"], 3.75)
                self.assertAlmostEqual(row["综合成本"], 31.15)
                self.assertAlmostEqual(row["毛利"], 275.4)
                self.assertAlmostEqual(row["净利润"], 275.4)
                self.assertIn("4/4件", logs[0]["运费匹配状态"])
                self.assertNotIn("无报价地区", logs[0]["运费来源"])
                self.assertFalse(any("运费未完整匹配" in warning for warning in warnings))

    def test_both_platforms_leave_profit_unknown_when_any_effective_order_lacks_region(self):
        for module in (pdd, tmall):
            with self.subTest(platform=module.__name__), TemporaryDirectory() as directory:
                _, rows, warnings, logs = self.build(Path(directory), module, True)
                row = rows[0]
                self.assertEqual(row["实际成交数量（去退款去补单后）"], 5)
                self.assertEqual(row["产品成本"], 10)
                for field in ("快递成本", "综合成本", "毛利", "净利润"):
                    self.assertEqual(row[field], "")
                self.assertIn("4/5件", logs[0]["运费匹配状态"])
                self.assertTrue(any("运费未完整匹配" in warning for warning in warnings))


if __name__ == "__main__":
    unittest.main()
