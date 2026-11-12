"""京东店铺日报，推广与订单口径以可核验的原始数据为准。"""

from collections import defaultdict
from decimal import Decimal
from io import BytesIO
from pathlib import Path
import json
import os

from jd_promotion import archives, as_day, as_money, identifier, load, records
from report_archive import store
from store_schema import STORE_HEADERS

ORDER = '京东订单原表'
AFTER_SALE = '京东售后原表'
SKU = '京东SKU原表'
ORDER_HEADERS = {'订单号', '商品ID', '商品名称', '订购数量', '商家应收', '订单状态', '付款确认时间'}
AFTER_SALE_HEADERS = {'服务单号', '订单号', '售后状态', '退款金额', '售后申请时间'}
REFUND_DATES = ('退款成功时间', '退款完成时间', '退款成功日期', '退款完成日期')
BLANK = (None, '', '-', '--')
PROVISIONAL_SUFFIX = '（系统计算中，该值仅供参考）'


def decrypt_order(path, password):
    try:
        import msoffcrypto
    except ModuleNotFoundError as error:
        raise ValueError('读取加密京东订单需要 Python msoffcrypto-tool 组件') from error
    with path.open('rb') as source:
        document = msoffcrypto.OfficeFile(source)
        try:
            document.load_key(password=password, verify_password=True)
            decoded = BytesIO()
            document.decrypt(decoded)
        except msoffcrypto.exceptions.DecryptionError as error:
            raise ValueError(f'京东订单解密失败：{path.name}') from error
    decoded.seek(0)
    return decoded


def payment_rows(root, dates, unlocked_order=None):
    orders = {}
    order_items = {}
    covered = defaultdict(set)
    blocked = []
    sources = list(archives(root, '2.订单数据库', {ORDER}))
    shops = {shop for _, shop, _ in sources}
    if unlocked_order:
        if len(shops) != 1:
            raise ValueError('提供订单解密副本时，原归档须能唯一确认京东店铺')
        encrypted = [entry for entry in sources if entry[2].get('workbook_encrypted')]
        if len(encrypted) != 1:
            raise ValueError('订单副本须对应唯一的加密归档')
        source = encrypted[0]
        sources = [entry for entry in sources if entry != source]
        sources.append((Path(unlocked_order), source[1], dict(source[2], workbook_encrypted=False)))
    sources.sort(key=lambda item: (item[0].stat().st_mtime_ns, str(item[0])), reverse=True)
    for path, shop, meta in sources:
        period = meta.get('period', '').split('至')
        if len(period) != 2:
            raise ValueError(f'京东订单归档缺少可验证的日期范围：{path}')
        start, end = map(as_day, period)
        applicable = {day for day in dates if start <= day <= end}
        if not applicable:
            continue
        encrypted = meta.get('workbook_encrypted')
        password = os.environ.get('DAILY_REPORT_JD_PASSWORD')
        if encrypted and not password:
            blocked.append((shop, str(path)))
            continue
        batch = defaultdict(list)
        items = defaultdict(list)
        stream = decrypt_order(path, password) if encrypted else None
        try:
            for row, context in records(path, ORDER_HEADERS, stream=stream):
                if row['付款确认时间'] in BLANK:
                    continue
                day = as_day(row['付款确认时间'])
                if day not in dates:
                    continue
                order_id = identifier(row['订单号'])
                sku_id = identifier(row['商品ID'])
                if not order_id or not sku_id:
                    raise ValueError(f'京东已支付订单缺少订单号或商品ID：{context}')
                receipt = str(row['商家应收']).strip()
                provisional = receipt.endswith(PROVISIONAL_SUFFIX)
                amount = as_money(receipt.removesuffix(PROVISIONAL_SUFFIX), context)
                buyer = identifier(row.get('下单帐号') or row.get('下单账号'))
                batch[order_id].append((day, amount, buyer if '*' not in buyer else '', provisional))
                count = as_money(row['订购数量'], context)
                if not count or count != count.to_integral_value():
                    raise ValueError(f'京东商品订购数量不是正整数：{context}')
                items[order_id].append((sku_id, count))
        finally:
            if stream is not None:
                stream.close()
        for order_id, lines in batch.items():
            if len(set(lines)) != 1:
                raise ValueError(f'京东同订单付款时间、金额或买家信息冲突：{path.name}')
            # Merchant receivable is repeated on every SKU of the same order.
            orders.setdefault((shop, order_id), lines[0])
            order_items.setdefault((shop, order_id), items[order_id])
        covered[shop].update(applicable)
    return orders, order_items, covered, blocked, shops


def refund_rows(root, dates):
    records_by_id = {}
    covered = defaultdict(set)
    blocked = defaultdict(set)
    sources = list(archives(root, '3.售后数据库', {AFTER_SALE}))
    sources.sort(key=lambda item: (item[0].stat().st_mtime_ns, str(item[0])), reverse=True)
    for path, shop, meta in sources:
        period = meta.get('period', '').split('至')
        if len(period) == 2:
            start, end = map(as_day, period)
            covered[shop].update(day for day in dates if start <= day <= end)
        for row, context in records(path, AFTER_SALE_HEADERS):
            service = identifier(row['服务单号'])
            if not service:
                raise ValueError(f'京东售后明细缺少服务单号：{context}')
            records_by_id.setdefault((shop, service), (row, context))
    result = defaultdict(dict)
    for (shop, _), (row, context) in records_by_id.items():
        if row['售后状态'] not in ('完成','退款成功','退款完成','已退款') or row['退款金额'] in BLANK:
            continue
        amount = as_money(row['退款金额'], context)
        if not amount:
            continue
        refund_days = {as_day(row[field]) for field in REFUND_DATES if row.get(field) not in BLANK}
        if len(refund_days) > 1:
            raise ValueError(f'京东退款成功日期冲突：{context}')
        if not refund_days:
            applied = as_day(row['售后申请时间'])
            blocked[shop].update(day for day in dates if day >= applied)
            continue
        day = next(iter(refund_days))
        order_id = identifier(row['订单号'])
        if not order_id:
            raise ValueError(f'京东成功退款缺少订单号：{context}')
        result[shop, day][order_id] = result[shop, day].get(order_id, Decimal(0)) + amount
    return result, covered, blocked


def merchandise_costs(root, orders, items):
    from report_extensions import load_catalog
    from shipping_costs import load_database_shipping
    from data_quality_protection import normalize_match_code
    catalog = load_catalog(root)
    shipping, _ = load_database_shipping(root)
    sku_codes = {}
    sources = list(archives(root, '1.产品数据库', {SKU}))
    sources.sort(key=lambda item: (item[0].stat().st_mtime_ns, str(item[0])), reverse=True)
    for path, shop, _ in sources:
        batch = {}
        for row, context in records(path, {'SKUID','商家SKU'}):
            sku = identifier(row['SKUID'])
            code = normalize_match_code(row['商家SKU'])
            if sku in batch and batch[sku] != code:
                raise ValueError(f'京东SKU商家编码冲突：{context}')
            batch[sku] = code
        for sku, code in batch.items():
            sku_codes.setdefault((shop, sku), code)
    totals = defaultdict(Decimal)
    missing = defaultdict(set)
    for order_key, (day, _, _, _) in orders.items():
        shop, _ = order_key
        for sku, quantity in items[order_key]:
            code = sku_codes.get((shop, sku), '')
            cost = catalog.unit_costs.get(code)
            quote = shipping.majority_region(code) if cost is not None else None
            if cost is None or quote.amount is None:
                missing[shop, day].add(code or sku)
            else:
                totals[shop, day] += quantity * (cost + Decimal(str(quote.amount)))
    return totals, missing


def build(root, dates, unlocked_order=None, platform_fee_rate=None):
    promotion, promo_sources = load(root, dates)
    orders, items, paid_coverage, blocked_orders, order_shops = payment_rows(root, dates, unlocked_order)
    refunds, refund_coverage, blocked_refunds = refund_rows(root, dates)
    shops = order_shops | {shop for shop, _ in promotion}
    if not shops:
        raise ValueError('未找到京东订单或京准通财务归档')
    payments = defaultdict(list)
    for (shop, _), (day, amount, buyer, provisional) in orders.items():
        payments[shop, day].append((amount, buyer, provisional))
    totals, missing = merchandise_costs(root, orders, items) if orders else ({}, {})
    if platform_fee_rate is not None:
        platform_fee_rate = as_money(platform_fee_rate, '京东平台费率')
        if platform_fee_rate > 1:
            raise ValueError('京东平台费率请填写0到1之间的小数')
    overhead_rate = Decimal('0.167')

    result = {'headers': STORE_HEADERS, 'rows': [], 'logs': [], 'evidence': {}}
    audit = {'日期': sorted(dates), '推广来源': promo_sources, '加密订单原表': blocked_orders, '分日': []}
    for shop in sorted(shops):
        for day in sorted(dates):
            row = dict.fromkeys(STORE_HEADERS)
            row.update({'日期':day, '平台':'京东', '店铺':shop})
            row['_净利润以本次为准'] = True
            stats = promotion.get((shop, day), {})
            cash = stats.get('cash')
            row['推广花费'] = float(cash) if cash is not None else None
            row['曝光次数'] = stats.get('impressions')
            row['点击次数'] = stats.get('clicks')
            paid = payments[shop, day]
            if day in paid_coverage[shop]:
                row['支付订单数'] = len(paid)
                row['支付金额'] = float(sum((amount for amount, _, _ in paid), Decimal(0)))
                if all(buyer for _, buyer, _ in paid):
                    row['支付买家数'] = len({buyer for _, buyer, _ in paid})
            if day in refund_coverage[shop] and day not in blocked_refunds[shop]:
                returned = refunds[shop, day]
                row['退款金额'] = float(sum(returned.values(), Decimal(0)))
                row['退款订单数'] = len(returned)
            messages = ['推广花费取京准通现金流水，按投放日期归日']
            if row['支付金额'] is None:
                messages.append('订单待解密、日期未覆盖或商家应收未能安全拆分')
            provisional_count = sum(flag for _, _, flag in paid)
            if provisional_count:
                messages.append(f'支付金额包含{provisional_count}笔平台标注系统计算中的商家应收')
            if row['退款金额'] is None:
                messages.append('售后未提供可核实的退款成功日期')
            missing_costs = sorted(missing.get((shop, day), ()))
            if (row['支付金额'] is not None and row['退款金额'] is not None and cash is not None
                    and platform_fee_rate is not None and not missing_costs and not provisional_count):
                net = Decimal(str(row['支付金额'])) - Decimal(str(row['退款金额']))
                row['店铺净利润'] = float(net * (1 - overhead_rate - platform_fee_rate)
                                         - totals.get((shop, day), Decimal(0)) - cash)
                messages.append('店铺净利润按净成交额扣商品成本、估算运费、均摊、京东平台费和现金推广费计算')
            else:
                messages.append('店铺净利润待退款成功日期、订单应收、SKU成本、运费和京东费率齐全后核算')
            result['logs'].append({'date': day, 'shop': shop, 'message': '；'.join(messages)})
            key = f'{day}|京东|{shop}|'
            result['evidence'][key + '推广花费'] = '京准通现金消费／投放日期'
            if row['支付金额'] is not None:
                result['evidence'][key + '支付金额'] = '京东订单原表／付款确认时间／每笔订单商家应收计一次'
            if stats.get('impressions') is not None:
                messages.append('曝光和点击只覆盖已归档的推广明细')
                result['evidence'][key + '曝光次数'] = '已归档京准通推广明细（仅已覆盖渠道）'
                result['evidence'][key + '点击次数'] = '已归档京准通推广明细（仅已覆盖渠道）'
            result['logs'][-1]['message'] = '；'.join(messages)
            result['rows'].append(row)
            audit['分日'].append({'店铺':shop, '日期':day, '已支付订单数':row['支付订单数'],
                                 '含暂估应收订单数':provisional_count, '缺成本或运费编码':missing_costs, **stats})
    return result, audit


def run(root, dates, archive_dir, output=None, unlocked_order=None, platform_fee_rate=None):
    payload, audit = build(root, dates, unlocked_order, platform_fee_rate)
    from report_extensions import store_payload
    by_month = defaultdict(list)
    for row in payload['rows']:
        by_month[row['日期'][:7]].append(row)
    if output and len(by_month) > 1:
        raise ValueError('京东店铺单个输出路径只能保存一个月')
    for month, rows in sorted(by_month.items()):
        path = Path(output) if output else archive_dir / f'店铺日报_{month}.xlsx'
        # A promotion-only refresh must not erase prior, already verified JD orders or profit.
        if path.exists():
            prior = {(r['日期'], r['店铺']):r for r in store_payload(path)['rows'] if r['平台'] == '京东'}
            for row in rows:
                previous = prior.get((row['日期'], row['店铺']), {})
                for field in ('支付金额','支付订单数','支付买家数','退款金额','退款订单数','_口径说明'):
                    if row.get(field) is None and field in previous:
                        row[field] = previous[field]
                if (row['店铺净利润'] is None and previous.get('店铺净利润') is not None and
                        all(row[field] == previous.get(field) for field in ('支付金额','退款金额','推广花费'))):
                    row['店铺净利润'] = previous['店铺净利润']
        store(dict(payload, rows=rows), archive_dir, path)
        audit_path = path.with_name(path.stem + '_京东核对.json')
        previous = json.loads(audit_path.read_text(encoding='utf-8')) if audit_path.exists() else {}
        keys = {(row['日期'],row['店铺']) for row in rows}
        entries = [item for item in previous.get('分日',[]) if (item['日期'],item['店铺']) not in keys]
        entries += [item for item in audit['分日'] if (item['日期'],item['店铺']) in keys]
        temp = audit_path.with_suffix('.tmp')
        temp.write_text(json.dumps(dict(audit, 分日=entries), ensure_ascii=False, indent=2, default=str), encoding='utf-8')
        temp.replace(audit_path)
        print(f'京东店铺日报：{path}')
    if audit['加密订单原表']:
        print('京东订单仍有打开密码，需通过 --jd-order-input 提供可读取副本以补算支付等指标。')
