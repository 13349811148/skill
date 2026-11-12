from pathlib import Path
from tempfile import TemporaryDirectory
from unittest import TestCase
import csv
import os

from openpyxl import Workbook

from build_report import detect_platforms
from build_store_report import build_payload
from build_tmall_report import load_product_exports

SHOP = '测试旗舰店（天猫）'
DAY = '2026-11-11'
CONFIG = {'shops': [], 'shop_aliases': {}}


def csv_archive(root, folder, category, name, headers, rows):
    path = root / folder / SHOP / category / name
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open('w', newline='', encoding='utf-8') as handle:
        writer = csv.writer(handle)
        writer.writerow(headers)
        writer.writerows(rows)
    return path


def scenes(root):
    path = root / '4.推广数据库' / SHOP / '万相台场景结案' / '场景结案.xlsx'
    path.parent.mkdir(parents=True)
    workbook = Workbook()
    header = ['日期', '计划ID', '主体ID', '主体类型', '展现量', '点击量', '花费', '总成交金额', '转化周期']
    for sheet, key, subject_type, subject, metrics in (
        (workbook.active, '关键词推广', '商品', '1234567', (200, 10, 5, 100)),
        (workbook.create_sheet(), '店铺直达', '店铺', '7654321', (100, 5, 2, 50)),
    ):
        sheet.title = key
        sheet.append(header)
        sheet.append([DAY, 'plan', subject, subject_type, *metrics, 30])
        sheet.append([DAY, 'plan', subject, subject_type, *metrics[:3], 1, 15])
    workbook.save(path)


class TmallV3Tests(TestCase):
    def test_main_order_recognition_and_store_calculation(self):
        with TemporaryDirectory() as temp:
            root = Path(temp)
            order_headers = ['订单编号', '买家实付金额', '买家应付货款', '订单状态',
                             '订单创建时间', '订单付款时间', '宝贝总数量', '店铺名称']
            old = csv_archive(root, '2.订单数据库', '天猫主订单原表', 'old.csv', order_headers,
                              [('order1', '99', '100', '交易成功', DAY, DAY, 1, SHOP)])
            newer = csv_archive(root, '2.订单数据库', '天猫主订单原表', 'new.csv', order_headers, [
                ('order1', '20', '100', '交易关闭', DAY, DAY, 1, '测试旗舰店'),
                ('order2', '30', '30', '交易成功', '2026-11-10', DAY, 1, SHOP),
                ('unpaid', '0', '40', '交易关闭', DAY, '', 1, SHOP),
            ])
            os.utime(old, (100, 100))
            os.utime(newer, (200, 200))
            csv_archive(root, '3.售后数据库', '天猫退款原表', 'refunds.csv',
                        ['订单编号', '退款编号', '退款状态', '退款完结时间', '退款总额'], [
                            ('older_order', 'return1', '退款成功', DAY, '3.50'),
                            ('older_order', 'return2', '退款成功', DAY, '1.25'),
                            ('order1', 'return3', '退款关闭', DAY, '20'),
                        ])
            scenes(root)
            self.assertEqual(detect_platforms(root), {'tmall'})
            payload = build_payload(root, {DAY}, config=CONFIG, platforms={'天猫'})
            self.assertEqual(len(payload['rows']), 1)
            row = payload['rows'][0]
            expected = {'店铺': SHOP, '支付金额': 50.0, '支付订单数': 2, '退款金额': 4.75,
                        '退款订单数': 1, '净成交额': 45.25, '推广花费': 7, '推广成交额': 150,
                        '曝光次数': 300, '点击次数': 15, '_万相台未归属商品花费': 2}
            self.assertEqual({field: row[field] for field in expected}, expected)
            self.assertIsNone(row['访客数'])
            self.assertIsNone(row['店铺净利润'])
            self.assertEqual(row['推广ROI'], 45.25 / 7)

    def test_missing_refund_completion_does_not_turn_into_zero(self):
        with TemporaryDirectory() as temp:
            root = Path(temp)
            csv_archive(root, '3.售后数据库', '天猫退款原表', 'refunds.csv',
                        ['订单编号', '退款编号', '退款状态', '退款完结时间', '退款总额'], [
                            ('one', 'one', '退款成功', DAY, 10),
                            ('two', 'two', '退款成功', '', 2),
                        ])
            row, = build_payload(root, {DAY}, config=CONFIG, platforms={'天猫'})['rows']
            self.assertIsNone(row['退款金额'])
            self.assertIsNone(row['净成交额'])
            self.assertEqual(len(row['_现金流明细']['refunds']), 1)

    def test_product_template_uses_specification_code(self):
        with TemporaryDirectory() as temp:
            root = Path(temp)
            path = root / '1.产品数据库' / SHOP / '天猫商品发布模板' / '发布模板.xlsx'
            path.parent.mkdir(parents=True)
            workbook = Workbook()
            sheet = workbook.active
            sheet.title = '发布模板'
            sheet.append(['说明'])
            sheet.append(['商品信息'])
            sheet.append(['商品Id', '商家编码', 'skuId', '价格（元）', '商家编码'])
            sheet.append(['1234567', 'SPU', '7654321', 20, 'SKU'])
            workbook.save(path)
            index, warnings = load_product_exports(root)
            self.assertFalse(warnings)
            self.assertEqual(index.resolve('1234567', '7654321')['商品SKU'], 'SKU')
