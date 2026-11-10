import unittest
from product_cost_rules import apply_name_profit_override, EXPIRY_PROFIT_NOTE


class NameProfitOverrideTests(unittest.TestCase):
    def test_expiry_name_zeroes_net_only_and_is_idempotent(self):
        for name in ['【临期】吐司300g','临期-吐司300g','普通面包\n临期-吐司300g']:
            row={'货品名称':name,'净利润':-8.25,'毛利':3,'推广费用':11.25,'销售额':20,'备注':'原备注'}
            self.assertTrue(apply_name_profit_override(row))
            self.assertEqual(row['净利润'],0)
            self.assertEqual((row['毛利'],row['推广费用'],row['销售额']),(3,11.25,20))
            once=dict(row);apply_name_profit_override(row);self.assertEqual(row,once)
            self.assertIn(EXPIRY_PROFIT_NOTE,row['备注'])

    def test_unknown_cost_can_have_explicitly_confirmed_zero_profit(self):
        row={'货品名称':'临期-面包','净利润':None,'毛利':None,'产品成本':None}
        apply_name_profit_override(row)
        self.assertEqual(row['净利润'],0)
        self.assertIsNone(row['毛利']);self.assertIsNone(row['产品成本'])

    def test_code_or_sku_name_alone_does_not_match_goods_name(self):
        for name in [None,'','普通吐司']:
            row={'货品名称':name,'SKU编码':'X临期','SKU名称':'【临期】商品标题','净利润':-5}
            self.assertFalse(apply_name_profit_override(row));self.assertEqual(row['净利润'],-5)


if __name__=='__main__':unittest.main()
