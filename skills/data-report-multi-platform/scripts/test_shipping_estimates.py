from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from openpyxl import Workbook
from shipping_costs import ShippingTable, ShippingQuote, load_shipping_table
from shipping_estimates import name_grams, fill_same_weight_quotes


class ShippingEstimatesTests(unittest.TestCase):
    def test_nominal_weight_is_total_per_sale_unit_and_unknown_gifts_are_not_ignored(self):
        self.assertEqual(name_grams('吐司300g【2袋】'),600)
        self.assertEqual(name_grams('香蕉280g（4根）*1箱'),280)
        self.assertEqual(name_grams('面包2kg*1盒'),2000)
        self.assertIsNone(name_grams('吐司400g【2袋】+赠品1个'))
        self.assertIsNone(name_grams('纸箱400g'))
        self.assertIsNone(name_grams('【托盒】面包400g托盒1只'))

    def test_same_weight_median_uses_measured_donors_without_repeating_packing(self):
        info = {c:{'货品名称':n,'重量(g)':w,'体积(cm³)':v} for c,n,w,v in [
            ('a','面包300g【1袋】',None,None),('b','吐司300g【1袋】',420,800),
            ('c','蛋糕300g【1袋】',430,900),('d','面包301g【1袋】',440,900),
            ('e','面包300g【1袋】',None,None),('f','面包206g【1袋】',None,None)]}
        prices = {(c,'圆通','安徽'):ShippingQuote(p,c) for c,p in [('a',None),('b',2.4),('c',2.6),('d',10),('e',None),('f',None)]}
        table=ShippingTable(prices)
        fill_same_weight_quotes(table,info,{})
        self.assertEqual(table.lookup('a','圆通','安徽').amount,2.5)
        self.assertEqual(table.lookup('e','圆通','安徽').amount,2.5)
        self.assertIsNone(table.lookup('f','圆通','安徽').amount)
        self.assertNotIn('e的',table.lookup('a','圆通','安徽').source)

    def test_workbook_loader_propagates_weight_only_and_runtime_reference_labels(self):
        with TemporaryDirectory() as folder:
            wb=Workbook();s=wb.active;s.title='编码匹配'
            s.append(['编码','货品名称','重量(g)','体积(cm³)','数据状态'])
            s.append(['A','吐司300g【1袋】',420,None,'缺体积，按重量估算'])
            s.append(['B','面包300g【1袋】',None,None,'缺重量及体积'])
            for title in ('圆通各地运费','邮政各地运费'):
                sheet=wb.create_sheet(title);sheet.append(['编码','货品名称','计费重(kg)','安徽'])
                sheet.append(['A','产品',0.42,2.4]);sheet.append(['B','产品',None,None])
            path=Path(folder)/'运费成本.xlsx';wb.save(path)
            table,_=load_shipping_table(path)
            self.assertIn('按重量估算',table.lookup('A','圆通','安徽').estimate)
            self.assertEqual(table.lookup('B','圆通','安徽').amount,2.4)
            self.assertIn('同克重估算',table.lookup('B','圆通','安徽').estimate)


if __name__=='__main__':unittest.main()
