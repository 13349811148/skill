"""Validate approved reference costs, exact SKU identity and unmatched blanks."""
from pathlib import Path
from types import SimpleNamespace
from tempfile import TemporaryDirectory
import unittest

import build_pdd_report as pdd
import build_tmall_report as tmall
from data_quality_protection import DataQualityError, load_cost_table_records

HEADERS = ['编码类型','组合装编码','货品编码（组合内）','匹配成本(元)','参考测算成本(元)','匹配状态']


def summary(rows):
    return SimpleNamespace(sheet_name='编码成本汇总', rows=[['说明'], HEADERS, *rows], truncated=False)


class ProductCostTests(unittest.TestCase):
    def test_bundle_and_stock_identity_with_reference_fallback(self):
        costs, warnings = load_cost_table_records(Path('产品成本.xlsx'), [summary([
            ['组合装','PACK02','ITEM01×2',18,19,'已匹配'],
            ['货品','','ITEM01',9,None,'已匹配'],
            ['组合装','REF01','ITEM01×3',99,27,'参考测算'],
            ['货品','','PARTIAL01',1,2,'未完整匹配'],
            ['货品','','FALLBACK01',None,3,'已匹配'],
            ['货品','','EMPTY01',None,None,'已匹配'],
        ])])
        self.assertEqual(set(costs), {'pack02','item01','ref01','fallback01'})
        self.assertEqual(costs['pack02']['产品成本'],18)
        self.assertEqual(costs['item01']['产品成本'],9)
        self.assertEqual(costs['pack02']['编码表头'],'组合装编码')
        self.assertEqual(costs['item01']['编码表头'],'货品编码（组合内）')
        self.assertEqual(costs['pack02']['行号'],3)
        self.assertEqual(costs['ref01']['产品成本'],27)
        self.assertEqual(costs['fallback01']['产品成本'],3)
        self.assertEqual(costs['ref01']['成本性质'],'参考测算')
        self.assertEqual(costs['pack02']['成本性质'],'已匹配')
        self.assertEqual(costs['ref01']['成本表头'],'参考测算成本(元)')
        self.assertIn('参考测算成本(元)（参考测算）',costs['ref01']['来源'])
        self.assertTrue(any('采用参考测算成本2条' in warning for warning in warnings))

    def test_normalized_duplicate_conflict_still_stops(self):
        with self.assertRaises(DataQualityError):
            load_cost_table_records(Path('产品成本.xlsx'), [summary([
                ['组合装','ＰＡＣＫ 02','ITEM01×2',18,None,'已匹配'],
                ['组合装','pack02','ITEM01×2',19,None,'已匹配'],
            ])])

    def test_reference_only_is_a_valid_cost_table(self):
        costs,_=load_cost_table_records(Path('产品成本.xlsx'), [summary([
            ['组合装','REF01','ITEM01×3',None,27,'参考测算'],
        ])])
        self.assertEqual(costs['ref01']['产品成本'],27)

    def test_only_missing_or_invalid_costs_stop_generation(self):
        with self.assertRaises(DataQualityError):
            load_cost_table_records(Path('产品成本.xlsx'), [summary([
                ['组合装','REF01','ITEM01×3',None,None,'参考测算'],
                ['货品','','ZERO',0,0,'已匹配'],
                ['货品','','INVALID','不是数字',-1,'已匹配'],
            ])])

    def test_invalid_matched_cost_can_use_a_valid_reference(self):
        costs,warnings=load_cost_table_records(Path('产品成本.xlsx'), [summary([
            ['货品','','REF01',0,7,'已匹配'],
            ['货品','','REF02','无效',8,'已匹配'],
        ])])
        self.assertEqual(costs['ref01']['产品成本'],7)
        self.assertEqual(costs['ref02']['产品成本'],8)
        self.assertTrue(any('必须大于0' in warning for warning in warnings))
        self.assertTrue(any('成本无效' in warning for warning in warnings))

    def test_discovery_precedence_and_ambiguity_for_both_platforms(self):
        with TemporaryDirectory() as name:
            root=Path(name);product=root/'1.产品数据库';product.mkdir()
            legacy=root/'营销活动监控';legacy.mkdir();(legacy/'成本表.xlsx').touch()
            (product/'运费成本.xlsx').touch()
            (product/'快递成本.xlsx').touch()
            preferred=product/'产品成本.xlsx';preferred.touch()
            for module in (pdd,tmall):
                self.assertEqual(module.resolve_cost_table_path(None,root,legacy/'营销活动.xlsx'),preferred)
                explicit=root/'指定成本.xlsx'
                self.assertEqual(module.resolve_cost_table_path(str(explicit),root,legacy/'营销活动.xlsx'),explicit)
            preferred.unlink();(product/'成本甲.xlsx').touch();(product/'成本乙.xlsx').touch()
            for module in (pdd,tmall):
                with self.assertRaises(module.ReportError):
                    module.resolve_cost_table_path(None,root,legacy/'营销活动.xlsx')
            (product/'成本甲.xlsx').unlink();(product/'成本乙.xlsx').unlink()
            for module in (pdd,tmall):
                self.assertEqual(module.resolve_cost_table_path(None,root,legacy/'营销活动.xlsx'),legacy/'成本表.xlsx')


if __name__ == '__main__':
    unittest.main()
