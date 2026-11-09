from decimal import Decimal
from types import SimpleNamespace
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import patch
import csv
import unittest
from goods_catalog import GoodsCatalog, goods_rows
from report_extensions import daily_profit
from report_exclusions import exclude_existing_sales, excluded_pdd_product
from product_cost_rules import apply_product_cost_rules


class GoodsReportingTests(unittest.TestCase):
    def catalog(self):
        rows=[['编码类型','组合装编码','货品编码（组合内）','产品名称','基础单位'],
              ['货品','','A','货品甲','袋'],['货品','','B','货品乙','盒'],
              ['组合装','PACK','A×2；B×1','两甲一乙','套'],
              ['组合装','NEST','PACK×3；A×1','嵌套装','套']]
        return GoodsCatalog.from_sheets([SimpleNamespace(rows=rows)])

    def row(self, code='PACK', quantity=4, **extra):
        return {'日期':'2026-11-08','店铺名称':'店','SKU编码':code,
                '实际成交数量（去退款去补单后）':quantity, '净利润':12, **extra}

    def test_nested_bom_multiplies_and_combines_same_leaf(self):
        self.assertEqual(self.catalog().expand(' NEST '), {'a':Decimal(7),'b':Decimal(3)})

    def test_product_annotations_and_goods_rollup_do_not_duplicate_profit(self):
        rows=[self.row(),self.row('A',3)]
        annotated,parts,error=self.catalog().annotate(rows[0])
        self.assertEqual(annotated['货品名称'],'货品甲\n货品乙')
        self.assertEqual(annotated['货品数量'],'8\n4')
        goods,issues=goods_rows(rows,self.catalog())
        self.assertFalse(issues)
        self.assertEqual([(r['货品编码'],r['货品数量']) for r in goods],[('A',11),('B',4)])
        self.assertFalse(any('净利润' in r for r in goods))
        self.assertEqual(daily_profit(rows)[('2026-11-08','店')]['value'],24)

    def test_cycles_and_missing_child_do_not_become_partial_quantities(self):
        catalog=self.catalog();catalog.records['a'].update(kind='组合装',bom='NEST×1')
        with self.assertRaisesRegex(ValueError,'循环引用'):catalog.expand('PACK')
        _,parts,error=self.catalog().annotate(self.row('unknown'))
        self.assertEqual(parts,{})
        self.assertIn('缺少货品',error)

    def test_empty_burn_has_no_goods_and_reduces_store_profit(self):
        row=self.row('unknown',0,净利润=-9,备注='空烧推广费')
        self.assertEqual(goods_rows([row],self.catalog()),([],[]))
        self.assertEqual(daily_profit([row])[('2026-11-08','店')]['value'],-9)

    def test_unresolved_skus_remain_separate_and_traceable(self):
        rows=[self.row('unknown',商品ID='P',**{'SKU ID':'1'}), self.row('unknown',商品ID='P',**{'SKU ID':'2'})]
        goods,issues=goods_rows(rows,self.catalog())
        self.assertEqual(len(goods),2)
        self.assertTrue(all(r['货品数量'] is None for r in goods))
        self.assertEqual({r['SKU ID'] for r in issues},{'1','2'})

    def test_missing_profit_is_not_zero_but_confirmed_numeric_zero_counts(self):
        self.assertIsNone(daily_profit([self.row(净利润=None)])[('2026-11-08','店')]['value'])
        self.assertEqual(daily_profit([self.row(净利润=0)])[('2026-11-08','店')]['value'],0)

    def test_exclusion_preserves_actual_promotion_and_is_idempotent(self):
        row=self.row(商品ID='991147507035',推广费用=8)
        rows,removed=exclude_existing_sales([row],'pdd')
        self.assertEqual(len(removed),1)
        self.assertEqual(rows[0]['净利润'],-8)
        self.assertEqual(rows[0]['实际成交数量（去退款去补单后）'],0)
        self.assertEqual(exclude_existing_sales(rows,'pdd'),(rows,[]))
        self.assertFalse(excluded_pdd_product('other'))
        self.assertEqual(exclude_existing_sales([row],'tmall'),([row],[]))

    def test_temporary_zero_ends_when_cost_becomes_known(self):
        row=self.row('ZB27',2,到手价=20,产品成本='',快递成本=2,推广费用=3,
                     实际成交金额=40)
        row['实际成交金额（去退款去补单后）']=40
        apply_product_cost_rules(row)
        self.assertEqual(row['净利润'],0)
        self.assertEqual(row['毛利'],'')
        row['产品成本']=8
        apply_product_cost_rules(row)
        self.assertAlmostEqual(row['净利润'],10.04)

    def test_confirmed_product_exclusion_precedes_order_quantity_and_freight(self):
        import build_pdd_report as pdd
        from portable_reports import read_workbook_sheets
        with TemporaryDirectory() as directory:
            root=Path(directory);path=root/'orders.csv'
            headers=['订单号','支付时间','样式ID','汇总_店铺名称','商品ID','商品数量(件)',
                     '商家实收金额(元)','商家编码-规格维度','订单状态','售后状态','商品名称']
            with path.open('w',encoding='utf-8-sig',newline='') as file:
                writer=csv.writer(file);writer.writerow(headers)
                writer.writerow(['EX','2026-11-05 12:00:00','S1','店','991147507035',99,9900,'A','已发货，待收货','','普通商品'])
                writer.writerow(['OK','2026-11-05 12:00:00','S2','店','OTHER',1,100,'A','已发货，待收货','','普通商品'])
            with patch.object(pdd,'discover_files',return_value=[path]),patch.object(pdd,'read_workbook_sheets',read_workbook_sheets):
                rows,warnings=pdd.load_order_aggregates(root,{'2026-11-05'},'exclude',cost_by_merchant_code={'a':{'产品成本':10}})
            self.assertEqual(len(rows),1)
            self.assertTrue(any('固定商品ID排除' in w for w in warnings))


if __name__=='__main__':unittest.main()
