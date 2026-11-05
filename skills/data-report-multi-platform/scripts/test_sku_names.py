from pathlib import Path
from types import SimpleNamespace
import unittest
from data_quality_protection import CostRecords, DataQualityError, load_cost_table_records, load_cost_sku_names
from product_cost_rules import product_output_row
from report_archive import product_key


class SkuNameTests(unittest.TestCase):
    def sheets(self,rows):
        return [SimpleNamespace(sheet_name='编码成本汇总',truncated=False,rows=[['说明'],['编码类型','组合装编码','货品编码（组合内）','产品名称','匹配成本(元)','匹配状态'],*rows])]

    def test_names_are_matched_even_when_cost_is_unavailable(self):
        sheets=self.sheets([['组合装','PACK','PART×2','整套商品',20,'已匹配'],['货品','','PART','单品商品',None,'未完整匹配']])
        costs,_=load_cost_table_records(Path('产品成本.xlsx'),sheets)
        self.assertEqual(len(costs),1)
        self.assertEqual(costs.sku_names,{'pack':'整套商品','part':'单品商品'})
        self.assertEqual(costs['pack']['SKU名称'],'整套商品')
        output=product_output_row({'商品SKU':'ＰＡＲＴ'},['SKU编码','SKU名称','产品成本'],costs)
        self.assertEqual(output,{'SKU编码':'ＰＡＲＴ','SKU名称':'单品商品','产品成本':''})
        self.assertNotIn('商品SKU',output)

    def test_name_conflict_stops_instead_of_guessing(self):
        with self.assertRaises(DataQualityError):
            load_cost_sku_names(self.sheets([['货品','','PART','商品甲',None,'无成本'],['货品','','part','商品乙',None,'无成本']]))

    def test_unmatched_name_is_blank_and_business_key_uses_renamed_code(self):
        self.assertEqual(product_output_row({'商品SKU':'UNKNOWN'},['SKU编码','SKU名称'],CostRecords({}))['SKU名称'],'')
        row={'日期':'2026-11-01','店铺名称':'测试店','商品ID':'1','SKU ID':'2'}
        self.assertEqual(product_key({**row,'商品SKU':'A'}),product_key({**row,'SKU编码':'A'}))
        self.assertNotEqual(product_key({**row,'SKU编码':'A'}),product_key({**row,'SKU编码':'B'}))


if __name__=='__main__':unittest.main()
