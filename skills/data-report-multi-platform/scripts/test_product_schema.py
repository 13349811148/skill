"""The public schema and partial-reference pass use the same total-profit rules."""
from pathlib import Path
from tempfile import TemporaryDirectory
from types import SimpleNamespace
from unittest.mock import patch
import sys
import unittest

import build_pdd_report as pdd
import build_tmall_report as tmall
import pdd_local_runtime as runtime
from product_cost_rules import apply_product_cost_rules, product_output_row, REMOVED_REPORT_COLUMNS
from product_schema import PRODUCT_COLUMNS, LEGACY_COLUMNS_32, compatible_headers, normalize_headers


class ProductSchemaTests(unittest.TestCase):
    def test_windows_csv_line_endings_keep_the_new_schema(self):
        with TemporaryDirectory() as directory:
            path=Path(directory)/'columns.csv'
            raw=pdd.COLUMNS_CSV.read_bytes().replace(b'\r\n',b'\n').replace(b'\n',b'\r\n')
            path.write_bytes(raw)
            for module in (pdd,tmall):
                with patch.object(module,'COLUMNS_CSV',path):
                    self.assertEqual(module.read_template_columns(),list(PRODUCT_COLUMNS))
        legacy=[c.replace('\n','\r\n') for c in LEGACY_COLUMNS_32]
        self.assertTrue(compatible_headers(legacy,list(PRODUCT_COLUMNS)))

    def test_both_platforms_expose_20_columns_with_no_removed_fields(self):
        for module in (pdd,tmall):
            columns=module.read_template_columns()
            self.assertEqual(len(columns),20)
            self.assertFalse(set(columns)&set(REMOVED_REPORT_COLUMNS))
            self.assertNotIn(module.PROMOTION_MECHANISM_COLUMN,columns)
            self.assertEqual(columns[columns.index('商品ID')+1],'SKU ID')
            self.assertEqual(columns[-1],'备注')
            self.assertEqual(columns.count('SKU ID'),1)

    def test_unknown_legacy_headers_are_rejected(self):
        self.assertFalse(compatible_headers([*LEGACY_COLUMNS_32,'人工新增'],list(PRODUCT_COLUMNS)))
        self.assertFalse(compatible_headers([c for c in LEGACY_COLUMNS_32 if c!='SKU ID'],list(PRODUCT_COLUMNS)))

    def test_partial_reference_pass_keeps_unit_cost_row_profit_and_deleted_columns(self):
        columns=pdd.read_template_columns()
        module=SimpleNamespace(load_marketing_rows=lambda path:[],load_costs_by_merchant_code=lambda path:({},[]),
                               resolve_marketing_path=lambda *a:None,resolve_cost_table_path=lambda *a:None,
                               PROMOTION_MECHANISM_COLUMN=pdd.PROMOTION_MECHANISM_COLUMN)
        def build():
            module.load_marketing_rows(Path('missing-reference-fixture/营销活动.xlsx'))
            row={'商品SKU':'A','实际成交数量（去退款去补单后）':3,'实际成交金额（去退款去补单后）':60,
                 '到手价':20,'产品成本':10,'快递成本':2.65,'推广费用':5}
            apply_product_cost_rules(row)
            return columns,[product_output_row(row,columns,{})],[],[]
        module.build_report=build
        with patch.object(sys,'argv',['test','--allow-missing-reference']),patch.object(runtime,'portable_install'):
            runtime.install(module)
            _,rows,_,_=module.build_report()
        self.assertEqual(set(rows[0]),set(columns))
        self.assertAlmostEqual(rows[0]['综合成本'],16.13)
        self.assertAlmostEqual(rows[0]['毛利'],11.61)
        self.assertAlmostEqual(rows[0]['净利润'],6.61)


if __name__=='__main__':unittest.main()
