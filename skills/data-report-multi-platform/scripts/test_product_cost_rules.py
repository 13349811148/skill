"""Check cost units, missing values and row-total profit independently."""
import unittest
from product_cost_rules import apply_product_cost_rules


class ProductCostRulesTests(unittest.TestCase):
    def row(self, **overrides):
        row = {"到手价":20, "产品成本":10, "实际成交数量（去退款去补单后）":3,
               "净销售额（实际成交+总补贴金额）":60, "推广费用":5,
               "店铺费用":100, "平摊管理费用":100}
        row.update(overrides)
        apply_product_cost_rules(row)
        return row

    def test_unit_gross_and_total_net_profit(self):
        row = self.row()
        self.assertAlmostEqual(row["其他均摊成本"],3.34)
        self.assertAlmostEqual(row["平台费用"],0.14)
        self.assertEqual(row["快递成本"],2.1)
        self.assertAlmostEqual(row["综合成本"],15.58)
        self.assertAlmostEqual(row["毛利"],4.42)
        self.assertAlmostEqual(row["毛利率"],0.221)
        self.assertAlmostEqual(row["净利润"],8.26)
        self.assertNotIn("店铺费用",row)
        self.assertNotIn("平摊管理费用",row)

    def test_shipping_stays_per_unit_for_multiple_units(self):
        self.assertEqual(self.row(**{"实际成交数量（去退款去补单后）":10})["快递成本"],2.1)

    def test_unknown_product_cost_does_not_become_zero(self):
        for value in [None,"",0,"未匹配"]:
            row=self.row(**{"产品成本":value})
            for field in ["综合成本","毛利","毛利率","净利润"]:
                self.assertEqual(row[field],"")
            self.assertEqual(row["快递成本"],2.1)
            self.assertAlmostEqual(row["其他均摊成本"],3.34)

    def test_known_unit_gross_does_not_require_unknown_subsidy(self):
        row=self.row(**{"净销售额（实际成交+总补贴金额）":""})
        self.assertAlmostEqual(row["毛利"],4.42)
        self.assertEqual(row["净利润"],"")

    def test_known_subsidy_is_included_once_in_total_profit(self):
        self.assertAlmostEqual(self.row(**{"净销售额（实际成交+总补贴金额）":63})["净利润"],11.26)

    def test_empty_burn_has_no_shipping_or_percentage_costs(self):
        row=self.row(**{"实际成交数量（去退款去补单后）":0,"到手价":"不适用（无有效销量）","备注":"空烧推广费：无有效销售"})
        for field in ["快递成本","其他均摊成本","平台费用","毛利"]:
            self.assertEqual(row[field],0)
        self.assertEqual(row["综合成本"],"")
        self.assertEqual(row["净利润"],-5)


if __name__ == '__main__':
    unittest.main()
