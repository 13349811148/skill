import unittest
from decimal import Decimal
from types import SimpleNamespace
from goods_catalog import GoodsCatalog, goods_rows
from goods_financials import PRICE


class GoodsFinancialTests(unittest.TestCase):
    def catalog(self):
        rows=[['编码类型','组合装编码','货品编码（组合内）','产品名称','基础单位'],
              ['货品','','A','甲','袋'],['货品','','B','乙','盒'],['货品','','C','丙','袋'],
              ['组合装','AB','A×2；B×1','甲乙套装','套'],['组合装','AA','A×2','两袋甲','套']]
        cat=GoodsCatalog.from_sheets([SimpleNamespace(rows=rows)])
        cat.unit_costs={'a':Decimal(10),'b':Decimal(20),'c':Decimal(5)}
        return cat

    def row(self,code='AB',date='2026-11-08',qty=4,sales=80,gross=20,promo=8,net=12,**extra):
        return {'日期':date,'店铺名称':'店','SKU编码':code,'商品ID':'P','SKU ID':code,
                '实际成交数量（去退款去补单后）':qty,'实际成交金额（去退款去补单后）':sales,
                '毛利':gross,'推广费用':promo,'净利润':net,**extra}

    def test_mixed_bundle_uses_cost_times_quantity_and_not_plain_quantity(self):
        audit={};rows,issues=goods_rows([self.row()],self.catalog(),audit)
        self.assertFalse(issues)
        a,b=rows
        self.assertEqual([a['货品数量'],b['货品数量']],[8,4])
        self.assertEqual([a['销售额'],b['销售额']],[40,40])
        self.assertEqual([a[PRICE],b[PRICE]],[5,10])
        self.assertEqual([a['毛利'],a['推广费用'],a['净利润']],[10,4,6])
        self.assertTrue(all(r['complete'] for r in audit['daily_reconciliation']))

    def test_missing_one_child_cost_keeps_all_mixed_money_blank(self):
        cat=self.catalog();cat.unit_costs.pop('b');audit={}
        rows,_=goods_rows([self.row()],cat,audit)
        for row in rows:
            for field in ['销售额','毛利','推广费用','净利润',PRICE]:self.assertIsNone(row[field])
        self.assertEqual([r['货品数量'] for r in rows],[8,4])
        self.assertEqual(len(audit['unallocated']),1)
        self.assertTrue(all(not r['complete'] for r in audit['daily_reconciliation']))

    def test_missing_contribution_masks_same_goods_total_without_hiding_other_goods(self):
        cat=self.catalog();cat.unit_costs.pop('b')
        rows,_=goods_rows([self.row(),self.row('A',qty=1,sales=30),self.row('C',qty=1,sales=7)],cat)
        by={r['货品编码']:r for r in rows}
        self.assertIsNone(by['A']['销售额'])
        self.assertEqual(by['A']['货品数量'],9)
        self.assertEqual(by['C']['销售额'],7)

    def test_single_leaf_keeps_known_money_even_without_cost_and_does_not_restore_missing_gross(self):
        cat=self.catalog();cat.unit_costs={}
        rows,_=goods_rows([self.row('AA',qty=3,sales=60,gross=None,promo=None,net=0)],cat)
        row=rows[0]
        self.assertEqual(row['货品数量'],6)
        self.assertEqual(row['销售额'],60)
        self.assertEqual(row[PRICE],10)
        self.assertEqual(row['净利润'],0)
        self.assertIsNone(row['毛利'])

    def test_zero_sales_promotion_stays_independent_and_has_no_unit_price(self):
        source=self.row('UNKNOWN',qty=0,sales=0,gross=0,promo=9,net=-9,备注='空烧推广费')
        rows,issues=goods_rows([source,self.row('A',qty=1,sales=30)],self.catalog())
        self.assertFalse(issues)
        self.assertTrue(rows[-1]['货品编码'].startswith('空烧推广费:'))
        self.assertEqual(rows[-1]['净利润'],-9)
        self.assertEqual(rows[-1]['推广费用'],9)
        self.assertIsNone(rows[-1][PRICE])
        self.assertEqual(rows[0]['净利润'],12)

    def test_daily_rounding_preserves_total_and_is_deterministic(self):
        sources=[self.row(c,qty=1,sales=10,gross=.005,promo=0,net=.005) for c in ['A','B','C']]
        audit={};rows,_=goods_rows(sources,self.catalog(),audit)
        self.assertEqual(sum(Decimal(str(r['净利润'])) for r in rows),Decimal('.02'))
        reversed_rows,_=goods_rows(list(reversed(sources)),self.catalog())
        self.assertEqual(rows,reversed_rows)

    def test_unresolved_sku_keeps_receipt_visible_but_does_not_invent_quantity(self):
        rows,issues=goods_rows([self.row('UNKNOWN',qty=1,sales=19.9,gross=None,promo=0,net=None)],self.catalog())
        self.assertEqual(len(issues),1)
        self.assertTrue(rows[0]['货品编码'].startswith('待拆分:'))
        self.assertEqual(rows[0]['销售额'],19.9)
        self.assertIsNone(rows[0]['货品数量'])
        self.assertIsNone(rows[0][PRICE])


if __name__=='__main__':unittest.main()
