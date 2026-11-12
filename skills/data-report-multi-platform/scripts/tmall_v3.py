"""Readers for original Tmall workbooks archived by the V3 operations database."""
from collections import defaultdict
from decimal import Decimal
from functools import lru_cache
from pathlib import Path
import json
import re

from handoff_protection import is_excluded_order_path
from order_line_protection import classify_order_line
from store_cashflow import amount, day, exclusions, money, records, text

PLATFORM = '天猫'
ORDER = '天猫主订单原表'
REFUND = '天猫退款原表'
PROMOTION = '万相台场景结案'
PRODUCT = '天猫商品发布模板'
SCENES = {'获客易', '关键词推广', '人群推广', '店铺直达'}
MAIN_HEADERS = {'订单编号', '买家实付金额', '订单状态', '订单创建时间',
                '订单付款时间', '宝贝总数量', '店铺名称'}
REFUND_HEADERS = {'订单编号', '退款编号', '退款完结时间', '退款总额', '退款状态'}
PROMO_HEADERS = {'日期', '计划ID', '主体ID', '主体类型', '展现量', '点击量', '花费', '总成交金额'}
PROMO_FIELDS = {'曝光次数': '展现量', '点击次数': '点击量',
                '推广花费': '花费', '推广成交额': '总成交金额'}
POLICY = ('天猫支付按主订单付款日汇总；优先商家实收，原表仅有买家实付时使用买家实付金额。'
          '退款只计退款成功，按退款完结时间汇总退款总额；同一订单的多笔退款分别计金额，退款订单数按订单编号去重。'
          '净成交额=当日支付金额−当日成功退款。')


def is_main_order(headers):
    normalized = {text(value) for value in headers}
    detail_fields = {'子订单编号', '子订单号', '宝贝id', '商品id', 'skuid', 'sku_id', '样式id'}
    return MAIN_HEADERS <= normalized and not {value.replace(' ', '').lower() for value in normalized} & detail_fields


def shop_name(value):
    name = re.sub(r'[（(]\s*天猫\s*[）)]$', '', text(value)).strip()
    return name + '（天猫）' if name else ''


def is_tmall_path(path):
    return any(part.endswith(('（天猫）', '(天猫)')) or part in {ORDER, REFUND, PROMOTION, PRODUCT}
               for part in path.parts[:-1])


def archive_metadata(path):
    receipt = path.with_suffix(path.suffix + '.archive.json')
    return json.loads(receipt.read_text(encoding='utf-8-sig')) if receipt.exists() else {}


def archive_shop(path, category=None, row_shop=None):
    metadata = archive_metadata(path)
    if metadata.get('platform') not in (None, PLATFORM):
        raise ValueError(f'天猫归档的平台与目录不符: {path}')
    folder_shop = next((part for part in reversed(path.parts[:-1]) if part.endswith(('（天猫）', '(天猫)'))), '')
    if category and category in path.parts:
        folder_shop = path.parts[path.parts.index(category) - 1]
    labels = {shop_name(value) for value in (metadata.get('shop'), folder_shop, row_shop) if text(value)}
    if len(labels) != 1:
        raise ValueError(f'天猫归档缺店铺或店铺信息冲突: {path}')
    return labels.pop()


def archive_files(root, folder, category):
    folder = Path(root) / folder
    paths = [path for path in folder.glob(f'*/{category}/**/*')
             if path.suffix.lower() in {'.xlsx', '.xls', '.csv'} and path.is_file()
             and not path.name.startswith('~$') and not is_excluded_order_path(path)]
    return sorted(paths, key=lambda path: (path.stat().st_mtime_ns, path.name), reverse=True)


def _put_latest(target, key, value, path, source):
    previous = target.setdefault(key, (path, value))
    if previous[0] == path and previous[1][:-1] != value[:-1]:
        raise ValueError(f'天猫原表中业务编号重复且数据冲突: {source}')


@lru_cache(maxsize=4096)
def _non_sale(title, nature, sales_flag):
    return classify_order_line({'商品标题': title, '订单性质': nature, '是否计销售': sales_flag}).sales_excluded


def cashflow(root, dates, config):
    excluded, _ = exclusions(root, config, PLATFORM)
    orders = {}
    refunds = {}
    logs = []
    order_sources = []
    refund_sources = []
    shops = set()
    order_coverage = defaultdict(list)
    refund_coverage = defaultdict(list)
    for path in archive_files(root, '2.订单数据库', ORDER):
        shop = archive_shop(path, ORDER)
        shops.add(shop)
        order_sources.append(str(path))
        source_days = set()
        for row, source in records(path, is_main_order):
            if shop_name(row.get('店铺名称')) != shop:
                raise ValueError(f'主订单店铺与归档不符: {source}')
            oid = text(row.get('订单编号'))
            if not oid:
                raise ValueError(f'天猫主订单缺订单编号: {source}')
            paid = day(row.get('订单付款时间'))
            if text(row.get('订单付款时间')) and not paid:
                raise ValueError(f'天猫订单付款时间无效: {source}')
            receipt = next((row[field] for field in ('商家实收金额(元)', '商家实收金额', '买家实付金额')
                            if field in row), None)
            value = amount(receipt, source) if paid else None
            if paid:
                source_days.add(paid)
            title = row.get('商品标题') if text(row.get('宝贝种类')) == '1' else None
            non_sale = _non_sale(title, row.get('订单性质'), row.get('是否计销售'))
            _put_latest(orders, (shop, oid), (paid, value, non_sale, source), path, source)
        if source_days:
            order_coverage[shop].append((min(source_days), max(source_days)))

    for path in archive_files(root, '3.售后数据库', REFUND):
        shop = archive_shop(path, REFUND)
        shops.add(shop)
        refund_sources.append(str(path))
        source_days = set()
        for row, source in records(path, REFUND_HEADERS.issubset):
            rid, oid = text(row.get('退款编号')), text(row.get('订单编号'))
            if not rid or not oid:
                raise ValueError(f'天猫退款缺退款编号或订单编号: {source}')
            finished = day(row.get('退款完结时间'))
            if finished:
                source_days.add(finished)
            success = text(row.get('退款状态')) == '退款成功'
            value = amount(row.get('退款总额'), source) if success else None
            applied = day(row.get('退款申请时间') or row.get('申请时间'))
            _put_latest(refunds, (shop, rid), (oid, success, finished, applied, value, source), path, source)
        if source_days:
            refund_coverage[shop].append((min(source_days), max(source_days)))

    payments = defaultdict(list)
    returned = defaultdict(list)
    uncertain = defaultdict(set)
    for (shop, oid), (_, (paid, value, non_sale, source)) in orders.items():
        if paid in dates and (shop, oid) not in excluded and not non_sale:
            payments[paid, shop].append({'order': oid, 'amount': str(value), 'source': source})
    for (shop, rid), (_, (oid, success, finished, applied, value, source)) in refunds.items():
        if not success or (shop, oid) in excluded:
            continue
        original = orders.get((shop, oid))
        if original and original[1][2]:
            continue
        if not finished:
            uncertain[shop].update(dt for dt in dates if not applied or dt >= applied)
            logs.append({'shop': shop, 'source': source, 'message': '天猫退款成功记录缺完结日期，受影响日期的退款留空'})
        elif finished in dates:
            returned[finished, shop].append({'order': oid, 'refund': rid, 'amount': str(value), 'source': source})

    rows = {}
    for shop in shops:
        for dt in dates:
            paid_known = any(start <= dt <= end for start, end in order_coverage[shop])
            refund_known = dt not in uncertain[shop] and any(start <= dt <= end for start, end in refund_coverage[shop])
            if not paid_known and not refund_known and dt not in uncertain[shop]:
                continue
            paid = payments[dt, shop]
            refunded = returned[dt, shop]
            rows[dt, PLATFORM, shop] = {
                '支付金额': money(sum((Decimal(row['amount']) for row in paid), Decimal(0))) if paid_known else None,
                '支付订单数': len(paid) if paid_known else None,
                '退款金额': money(sum((Decimal(row['amount']) for row in refunded), Decimal(0))) if refund_known else None,
                '退款订单数': len({row['order'] for row in refunded}) if refund_known else None,
                '_现金流明细': {'payments': paid, 'refunds': refunded},
            }
    return {'rows': rows, 'logs': logs, 'shops': shops, 'order_sources': order_sources, 'refund_sources': refund_sources}


def promotion(root, dates):
    """Use the latest daily scene export per shop, and retain non-product spend."""
    snapshots = {}
    for path in archive_files(root, '4.推广数据库', PROMOTION):
        shop = archive_shop(path, PROMOTION)
        daily = defaultdict(lambda: defaultdict(dict))
        for row, source in records(path, PROMO_HEADERS.issubset):
            if text(row.get('日期')) in {'合计', '总计', '汇总'}:
                continue
            dt = day(row.get('日期'))
            if not dt:
                raise ValueError(f'万相台场景结案缺少逐行日期: {source}')
            if dt not in dates or (dt, shop) in snapshots:
                continue
            # The sheet name distinguishes ad scenes; same-day plans and subjects are additive.
            scene = source.rsplit('/', 2)[1]
            if scene not in SCENES:
                continue
            subject_type, subject = text(row.get('主体类型')), text(row.get('主体ID'))
            plan = text(row.get('计划ID'))
            if not subject_type or not subject or not plan:
                raise ValueError(f'万相台明细缺计划或主体: {source}')
            metrics = tuple(amount(row.get(column), source) for column in PROMO_FIELDS.values())
            if any(value != value.to_integral_value() for value in metrics[:2]):
                raise ValueError(f'万相台展现量或点击量不是整数: {source}')
            key = scene, plan, subject_type, subject
            raw_window = row.get('转化周期')
            window = amount(raw_window, source) if text(raw_window) else Decimal(0)
            if daily[dt][key].setdefault(window, metrics) != metrics:
                raise ValueError(f'万相台同日同计划主体、同转化周期数据冲突: {source}')
        for dt, lines in daily.items():
            totals = [Decimal(0)] * len(PROMO_FIELDS)
            products = defaultdict(lambda: Decimal(0))
            unallocated = Decimal(0)
            common_windows = set.intersection(*(set(windows) for windows in lines.values()))
            attribution_window = max(common_windows) if common_windows else None
            for (_, _, subject_type, subject), windows in lines.items():
                delivery = {metrics[:3] for metrics in windows.values()}
                if len(delivery) != 1:
                    raise ValueError(f'万相台不同转化周期的展现、点击或花费不一致: {path}/{dt}/{subject}')
                metrics = delivery.pop()
                for index, value in enumerate(metrics):
                    totals[index] += value
                if attribution_window is not None:
                    totals[3] += windows[attribution_window][3]
                if subject_type in {'商品', '宝贝'}:
                    products[subject] += metrics[2]
                else:
                    unallocated += metrics[2]
            snapshots[dt, shop] = {
                'source': path,
                'fields': dict(zip(PROMO_FIELDS, [int(totals[0]), int(totals[1]), money(totals[2]),
                                                money(totals[3]) if attribution_window is not None else None])),
                'attribution_window': int(attribution_window) if attribution_window is not None else None,
                'products': {key: money(value) for key, value in products.items()},
                'unallocated': money(unallocated),
            }
    return snapshots
