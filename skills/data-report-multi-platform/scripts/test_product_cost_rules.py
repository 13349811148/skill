"""Check cost units, missing values and the requested direct profit subtraction."""
import unittest
from product_cost_rules import apply_product_cost_rules


class ProductCostRulesTests(unittest.TestCase):
    def row(self, **overrides):
        row = {"到手价":20, "产品成本":10, "实际成交数量（去退款去补单后）":3,
               "净销售额（实际成交+总补贴金额）":60, "推广费用":5, "快递成本":2.65,
               "店铺费用":100, "平摊管理费用":100}
        row.update(overrides)
        apply_product_cost_rules(row)
        return row

    def test_net_profit_is_gross_minus_promotion_without_quantity_multiplier(self):
        row = self.row()
        self.assertAlmostEqual(row["其他均摊成本"],3.34)
        self.assertAlmostEqual(row["平台费用"],0.14)
        self.assertEqual(row["快递成本"],2.65)
        self.assertAlmostEqual(row["综合成本"],16.13)
        self.assertAlmostEqual(row["毛利"],3.87)
        self.assertAlmostEqual(row["毛利率"],0.1935)
        self.assertAlmostEqual(row["净利润"],-1.13)
        self.assertNotIn("店铺费用",row)
        self.assertNotIn("平摊管理费用",row)

    def test_shipping_stays_per_unit_for_multiple_units(self):
        self.assertEqual(self.row(**{"实际成交数量（去退款去补单后）":10})["快递成本"],2.65)

    def test_unknown_product_cost_does_not_become_zero(self):
        for value in [None,"",0,"未匹配"]:
            row=self.row(**{"产品成本":value})
            for field in ["综合成本","毛利","毛利率","净利润"]:
                self.assertEqual(row[field],"")
            self.assertEqual(row["快递成本"],2.65)
            self.assertAlmostEqual(row["其他均摊成本"],3.34)

    def test_known_unit_gross_does_not_require_unknown_subsidy(self):
        row=self.row(**{"净销售额（实际成交+总补贴金额）":""})
        self.assertAlmostEqual(row["毛利"],3.87)
        self.assertAlmostEqual(row["净利润"],-1.13)

    def test_subsidy_and_quantity_do_not_change_direct_profit_definition(self):
        self.assertAlmostEqual(self.row(**{"净销售额（实际成交+总补贴金额）":63})["净利润"],-1.13)
        self.assertAlmostEqual(self.row(**{"实际成交数量（去退款去补单后）":10})["净利润"],-1.13)

    def test_missing_shipping_keeps_known_components_and_blanks_dependent_profit(self):
        for value in (None, "", "#N/A", 0, -2):
            row = self.row(**{"快递成本": value})
            for field in ("快递成本", "综合成本", "毛利", "毛利率", "净利润"):
                self.assertEqual(row[field], "")
            self.assertEqual(row["产品成本"], 10)
            self.assertAlmostEqual(row["其他均摊成本"], 3.34)
            self.assertAlmostEqual(row["平台费用"], 0.14)

    def test_partial_reference_reapplication_keeps_matched_shipping(self):
        row = self.row()
        apply_product_cost_rules(row)
        self.assertEqual(row["快递成本"], 2.65)
        self.assertAlmostEqual(row["净利润"], -1.13)

    def test_empty_burn_has_no_shipping_or_percentage_costs(self):
        row=self.row(**{"实际成交数量（去退款去补单后）":0,"到手价":"不适用（无有效销量）","备注":"空烧推广费：无有效销售"})
        for field in ["快递成本","其他均摊成本","平台费用","毛利"]:
            self.assertEqual(row[field],0)
        self.assertEqual(row["综合成本"],"")
        self.assertEqual(row["净利润"],-5)


if __name__ == '__main__':
    unittest.main()
