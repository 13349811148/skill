from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import patch
from decimal import Decimal
import csv,json,os,unittest
from store_cashflow import load,apply,day,amount,POLICY_NOTE

CONFIG={'shops':[{'平台':'拼多多','店铺':'店'}],'shop_aliases':{'店别名':'店'}}
ORDER_HEADERS=['订单号','样式ID','商品ID','支付时间','订单成交时间','商家实收金额(元)','汇总_店铺名称','订单状态','商品名称','订单性质','是否计销售']
REFUND_HEADERS=['售后编号','订单编号','售后状态','退款金额','同意退款时间','申请时间','店铺名称','退款成功时间']


def order(oid,dt,money='10',product='P',title='正常商品'):
    return [oid,'SKU',product,dt+' 12:00:00',dt+' 12:00:01',money,'店别名','已退款',title,'正常','是']


def refund(rid,oid,dt,money='3',status='退款成功'):
    return [rid,oid,status,money,dt+' 12:00:00' if dt else '', '2026-11-07 10:00:00','店','2026-11-07 12:00:00']


class StoreCashflowTests(unittest.TestCase):
    def write(self,root,folder,name,headers,rows,mtime=100):
        path=root/folder/name;path.parent.mkdir(parents=True,exist_ok=True)
        with path.open('w',encoding='utf-8-sig',newline='') as file:
            writer=csv.writer(file);writer.writerow(headers);writer.writerows(rows)
        os.utime(path,(mtime,mtime));return path

    def ledger(self,root,orders,refunds,dates=('2026-11-07','2026-11-08','2026-11-09')):
        self.write(root,'订单数据','orders.csv',ORDER_HEADERS,orders)
        self.write(root,'售后数据','refunds.csv',REFUND_HEADERS,refunds)
        return load(root,set(dates),CONFIG)

    def test_refund_is_booked_by_event_day_and_not_order_payment_day(self):
        with TemporaryDirectory() as directory:
            root=Path(directory)
            d=self.ledger(root,[order('O','2026-11-07','20'),order('OTHER','2026-11-09','10')],
                [refund('R1','O','2026-11-08','3'),refund('R2','O','2026-11-08','2')])
        r=d['rows']['2026-11-08','拼多多','店']
        self.assertEqual(r['支付金额'],0)
        self.assertEqual(r['退款金额'],5)
        self.assertEqual(r['退款订单数'],1)
        self.assertEqual(d['rows']['2026-11-07','拼多多','店']['支付金额'],20)
        self.assertIsNone(d['rows']['2026-11-07','拼多多','店']['退款金额'])

    def test_success_only_and_latest_snapshots_not_accumulated(self):
        with TemporaryDirectory() as directory:
            root=Path(directory)
            self.ledger(root,[order('O','2026-11-08')],[refund('R','O','2026-11-08','3')])
            self.write(root,'订单数据','new.csv',ORDER_HEADERS,[order('O','2026-11-08','12')],mtime=200)
            self.write(root,'售后数据','new.csv',REFUND_HEADERS,[refund('R','O','2026-11-08','4'),
                refund('PENDING','O','2026-11-08','99','待商家处理')],mtime=200)
            d=load(root,{'2026-11-08'},CONFIG)
        r=d['rows']['2026-11-08','拼多多','店']
        self.assertEqual(r['支付金额'],12)
        self.assertEqual(r['支付订单数'],1)
        self.assertEqual(r['退款金额'],4)
        self.assertEqual(d['counts']['ignored_refund_snapshots'],1)

    def test_preserves_confirmed_order_exclusions_without_extending_product_exclusions(self):
        with TemporaryDirectory() as directory:
            root=Path(directory);registry=root/'_系统/日报配置/订单剔除.json';registry.parent.mkdir(parents=True)
            registry.write_text(json.dumps({'groups':[{'平台':'拼多多','店铺':'店','订单号':['EX']}]}),encoding='utf-8')
            d=self.ledger(root,[order('EX','2026-11-08','99'),order('KEEP','2026-11-08','10','1021621380194'),
                order('SUP','2026-11-08','8',title='补差价')],
                [refund('R','EX','2026-11-08','7'),refund('S','SUP','2026-11-08','8')])
        r=d['rows']['2026-11-08','拼多多','店']
        self.assertEqual(r['支付金额'],10)
        self.assertEqual(r['支付订单数'],1)
        self.assertEqual(r['退款金额'],0)
        self.assertEqual(d['counts']['excluded_refund_events'],2)

    def test_missing_success_date_does_not_become_a_zero_refund(self):
        with TemporaryDirectory() as directory:
            d=self.ledger(Path(directory),[order('O','2026-11-08')],
                [refund('MISSING','O','','5'),refund('OTHER','O','2026-11-08','2')])
        self.assertIsNone(d['rows']['2026-11-08','拼多多','店']['退款金额'])
        self.assertTrue(d['logs'])

    def test_conflicting_same_snapshot_order_is_rejected(self):
        with TemporaryDirectory() as directory:
            with self.assertRaisesRegex(ValueError,'冲突支付行'):
                self.ledger(Path(directory),[order('O','2026-11-08','10'),order('O','2026-11-08','20')],[])

    def test_nonzero_receipt_cannot_use_creation_or_transaction_date(self):
        with TemporaryDirectory() as directory:
            row=order('O','2026-11-08');row[3]=''
            with self.assertRaisesRegex(ValueError,'无有效支付时间'):
                self.ledger(Path(directory),[row],[])
        with TemporaryDirectory() as directory:
            row=order('UNPAID','2026-11-08','0.000');row[3]=''
            data=self.ledger(Path(directory),[row],[])
            self.assertIsNone(data['rows']['2026-11-08','拼多多','店']['支付金额'])

    def test_backend_is_check_only_and_legacy_retained_net_cannot_override_new_flow(self):
        row={'日期':'2026-11-08','平台':'拼多多','店铺':'店','支付金额':999,'退款金额':123,'支付订单数':88,
             '推广花费':2,'店铺净利润':7,'访客数':300,'_原值保留':{'净成交额':876,'推广ROI':438,'新客数':6},'_口径说明':'历史说明'}
        ledger={'rows':{('2026-11-08','拼多多','店'):{'支付金额':10,'支付订单数':1,'退款金额':20,'退款订单数':2}},'logs':[]}
        logs=[];evidence={};apply([row],ledger,logs,evidence)
        self.assertEqual(row['净成交额'],-10)
        self.assertIsNone(row['推广ROI']);self.assertIsNone(row['推广费比'])
        self.assertEqual(row['_原值保留'],{'新客数':6})
        self.assertEqual(row['_历史现金流保留']['净成交额'],876)
        self.assertEqual(row['店铺净利润'],7);self.assertEqual(row['访客数'],300)
        self.assertEqual(row['_后台金额复核'][0]['后台复核'],999)
        self.assertEqual(row['_口径说明'],POLICY_NOTE)

    def test_missing_local_source_cannot_fall_back_to_backend(self):
        row={'日期':'2026-11-08','平台':'拼多多','店铺':'店','支付金额':999,'退款金额':123,'推广花费':2}
        apply([row],{'rows':{},'logs':[]},[],{})
        self.assertIsNone(row['支付金额']);self.assertIsNone(row['退款金额']);self.assertIsNone(row['净成交额'])

    def test_zero_and_negative_net_ratio_rules(self):
        for paid,refund_value,spend,roi,ratio in [(10,10,2,0,None),(10,0,0,None,0),(10,12,2,None,None)]:
            row={'日期':'2026-11-08','平台':'拼多多','店铺':'店','推广花费':spend}
            ledger={'rows':{('2026-11-08','拼多多','店'):{'支付金额':paid,'支付订单数':1,'退款金额':refund_value,'退款订单数':1}},'logs':[]}
            apply([row],ledger,[],{})
            self.assertEqual(row['推广ROI'],roi);self.assertEqual(row['推广费比'],ratio)
        self.assertEqual(day('2026/1/2 12:00:00'),'2026-01-02')
        with self.assertRaises(ValueError):amount('NaN','test')


if __name__=='__main__':unittest.main()
