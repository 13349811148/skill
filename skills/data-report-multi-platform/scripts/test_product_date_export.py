"""WPS compatibility: dates are numeric Excel dates, never ISO t=d cells."""
from pathlib import Path
from tempfile import TemporaryDirectory
import json,os,subprocess,unittest,xml.etree.ElementTree as ET,zipfile
from openpyxl import load_workbook
from product_schema import PRODUCT_COLUMNS

DEPS=Path(os.environ.get('CODEX_WORKSPACE_DEPENDENCIES',str(Path.home()/'.cache/codex-runtimes/codex-primary-runtime/dependencies')))
NODE=DEPS/('node/bin/node.exe' if os.name=='nt' else 'node/bin/node')
WRITER=Path(__file__).with_name('write_product_report.mjs')

@unittest.skipUnless(NODE.exists(),'bundled artifact runtime is not installed')
class DateExportTests(unittest.TestCase):
    def test_each_business_date_remains_distinct_without_an_iso_timestamp(self):
        dates=['2026-10-19','2026-10-20T00:00:00.000Z','2026-11-01T00:30:00+08:00']
        with TemporaryDirectory() as directory:
            root=Path(directory);data=root/'rows.json';output=root/'report.xlsx'
            rows=[{'日期':d,'商品ID':str(123456789012+i),'SKU ID':str(987654321012+i),'SKU编码':f'A{i}'} for i,d in enumerate(dates)]
            data.write_text(json.dumps({'columns':list(PRODUCT_COLUMNS),'rows':rows},ensure_ascii=False),encoding='utf-8')
            result=subprocess.run([str(NODE),str(WRITER),str(data),str(output)],env={**os.environ,'DAILY_REPORT_PREVIEW':'0'},capture_output=True,text=True)
            self.assertEqual(result.returncode,0,result.stderr)
            workbook=load_workbook(output,data_only=True)
            sheet=workbook.active
            self.assertEqual([sheet.cell(i,1).value.strftime('%Y-%m-%d') for i in range(2,5)],['2026-10-19','2026-10-20','2026-11-01'])
            self.assertTrue(all(sheet.cell(i,1).number_format=='yyyy-mm-dd' for i in range(2,5)))
            workbook.close()
            ns={'s':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
            with zipfile.ZipFile(output) as archive:
                tree=ET.fromstring(archive.read('xl/worksheets/sheet1.xml'))
                for index in range(2,5):
                    cell=tree.find(f's:sheetData/s:row/s:c[@r="A{index}"]',ns)
                    self.assertIn(cell.get('t'),(None,'n'))
                    self.assertTrue(cell.find('s:v',ns).text.isdecimal())

if __name__=='__main__':unittest.main()
