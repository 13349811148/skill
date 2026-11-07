"""Column migration keeps untouched months and protects confirmed exclusions."""
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import patch
import json
import unittest
import report_archive as archive
from product_schema import PRODUCT_COLUMNS

class ArchiveSchemaTests(unittest.TestCase):
    def rows(self):
        return {'2026-10':[{'日期':'2026-10-27','店铺名称':'店','商品ID':'100','SKU ID':'200','SKU编码':'A','实际成交数量（去退款去补单后）':3,'实际成交金额（去退款去补单后）':60,'快递成本':2.15,'产品成本':5.6,'毛利':20,'净利润':15,'备注':'已执行剔除'}],
                '2026-11':[{'日期':'2026-11-01','店铺名称':'店','商品ID':'101','SKU ID':'201','SKU编码':'B','实际成交数量（去退款去补单后）':2,'备注':''}]}

    def test_migration_retains_other_month_and_clears_only_requested_date(self):
        original=self.rows()
        with TemporaryDirectory() as folder:
            output=Path(folder)/'商品.xlsx';output.touch()
            output.with_suffix('.product.json').write_text(json.dumps({'date_notes':{'2026-10-27':'仅27日剔除'}}),encoding='utf-8')
            def read(path,columns,info):info['legacy']=True;return {m:[dict(r) for r in rows] for m,rows in original.items()}
            with patch.object(archive,'read_product',side_effect=read),patch.object(archive,'backup_legacy_product') as backup,patch.object(archive,'commit') as commit:
                archive.product(output,list(PRODUCT_COLUMNS),[],['2026-11-01'])
            payload=commit.call_args.args[1]
            self.assertIsNone(payload['existingWorkbook'])
            self.assertEqual([p['name'] for p in payload['sheets']],['2026-10','2026-11'])
            self.assertEqual(payload['sheets'][0]['rows'],original['2026-10'])
            self.assertEqual(payload['sheets'][1]['rows'],[])
            self.assertEqual(payload['date_notes'],{'2026-10-27':'仅27日剔除'})
            backup.assert_called_once_with(output)

    def test_protected_date_cannot_be_overwritten_by_schema_migration(self):
        with TemporaryDirectory() as folder:
            output=Path(folder)/'商品.xlsx';output.touch()
            output.with_suffix('.product.json').write_text(json.dumps({'date_notes':{'2026-10-27':'已剔除'}}),encoding='utf-8')
            with patch.object(archive,'commit') as commit,self.assertRaisesRegex(ValueError,'已有确认'):
                archive.product(output,list(PRODUCT_COLUMNS),[],['2026-10-27'])
            commit.assert_not_called()

if __name__=='__main__':unittest.main()
