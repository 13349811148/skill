"""PDD daily cash flow: paid-order receipts minus successful refund events."""
from collections import Counter, defaultdict
from datetime import date, datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path
import json
import re
from portable_reports import read_workbook_sheets
from handoff_protection import is_excluded_order_path
from order_line_protection import classify_order_line

CASH_FIELDS = ('支付金额', '支付订单数', '退款金额', '退款订单数')
CASH_DERIVED = ('客单价', '净成交额', '推广ROI', '推广费比')
PAY_DATES = ('订单支付时间', '支付时间', '订单付款时间', '付款时间')
POLICY_NOTE = ('用户2026-11-09确认：支付金额按订单支付日汇总商家实收金额(元)，退款只计退款成功，'
               '按同意退款时间归日（估算），净成交额=订单实收-当日成功退款。后台金额仅复核；'
               '指定订单剔除保留，商品ID销售排除不扩展到店铺现金流。')


def text(value):
    return str(value).strip() if value is not None else ''


def day(value):
    if isinstance(value, (datetime, date)):
        return value.strftime('%Y-%m-%d')
    raw = text(value)
    match = re.match(r'^(\d{4})[-/](\d{1,2})[-/](\d{1,2})(?:$|[ T])', raw)
    if match:
        try:
            return date(*(int(p) for p in match.groups())).isoformat()
        except ValueError:
            return None
    return None


def amount(value, context):
    try:
        result = Decimal(text(value).replace(',', ''))
    except InvalidOperation as error:
        raise ValueError(f'{context}：金额为空或无效') from error
    if not result.is_finite() or result < 0:
        raise ValueError(f'{context}：金额须为有限非负数')
    return result


def money(value):
    return float(value.quantize(Decimal('.01'), rounding=ROUND_HALF_UP))


def discover(root, legacy, current):
    from jd_promotion import is_jd_archive
    from tmall_v3 import is_tmall_path
    folder = root/legacy if (root/legacy).exists() else root/current
    return sorted((p for p in folder.rglob('*') if p.is_file() and p.suffix.lower() in {'.xlsx', '.xls', '.csv'}
                   and not p.name.startswith('~$') and not is_excluded_order_path(p) and not is_jd_archive(p)
                   and not is_tmall_path(p)),
                  key=lambda p: (p.stat().st_mtime_ns, str(p)), reverse=True) if folder.exists() else []


def canonical_shop(row, path, config, order_shop=None):
    aliases = config.get('shop_aliases', {})
    known = {s['店铺'] for s in config.get('shops', [])} | set(aliases.values())
    raw = next((text(row.get(h)) for h in ['汇总_店铺名称', '店铺名称', '店铺'] if text(row.get(h))), '')
    explicit = aliases.get(raw, raw) if raw else None
    inherited = next((aliases.get(part, part) for part in reversed(path.parts) if aliases.get(part, part) in known), None)
    if not inherited:
        for index, part in enumerate(path.parts[:-2]):
            if part in ('2.订单数据库', '3.售后数据库'):
                folder_shop = path.parts[index + 1]
                if path.name.startswith(folder_shop + '—'):
                    inherited = aliases.get(folder_shop, folder_shop)
                break
    shop = explicit or inherited or order_shop
    if order_shop and shop and order_shop != shop:
        raise ValueError(f'售后店铺与订单归属冲突: {path.name}')
    if not shop:
        raise ValueError(f'现金流记录缺少明确店铺来源: {path}')
    return shop


def records(path, required):
    for sheet in read_workbook_sheets(path):
        for index, values in enumerate(sheet.rows[:20]):
            headers = [text(v) for v in values]
            if not required(set(headers)):
                continue
            for physical, raw in enumerate(sheet.rows[index+1:], index+2):
                if not any(text(v) for v in raw):
                    continue
                record = {}
                for h, v in zip(headers, raw):
                    if h and text(v):
                        if h in record and text(record[h]) != text(v):
                            raise ValueError(f'现金流重复表头值冲突: {path.name}/{sheet.sheet_name}/{physical}/{h}')
                        record[h] = v
                yield record, f'{path.name}/{sheet.sheet_name}/{physical}'
            break


def exclusions(root, config, platform='拼多多'):
    path = root/'_系统/日报配置/订单剔除.json'
    if not path.exists():
        return set(), None
    data = json.loads(path.read_text(encoding='utf-8-sig'))
    aliases = config.get('shop_aliases', {})
    if platform == '天猫':
        from tmall_v3 import shop_name
        aliases = {**aliases, **{r['店铺']: shop_name(r['店铺']) for r in data['groups'] if r.get('平台') == platform}}
    return {(aliases.get(r['店铺'], r['店铺']), text(order)) for r in data['groups']
            for order in r['订单号'] if r.get('平台', '拼多多') == platform}, str(path)


def load(root, dates, config):
    root = Path(root); dates = set(dates)
    excluded, exclusion_source = exclusions(root, config)
    orders = {}; refunds = {}; logs = []; counts = Counter()
    order_sources = []; refund_sources = []
    order_files = discover(root, '订单数据', '2.订单数据库')
    for path in order_files:
        accepted = False
        for row, source in records(path, lambda h: '样式ID' in h and '商家实收金额(元)' in h and bool(h & {'订单号', '订单编号'})):
            accepted = True
            oid = text(row.get('订单号') or row.get('订单编号'))
            if not oid:
                raise ValueError(f'支付记录缺订单号: {source}')
            shop = canonical_shop(row, path, config)
            paid = next((d for h in PAY_DATES if (d := day(row.get(h)))), None)
            normalized = dict(row)
            normalized.setdefault('商品名称', row.get('商品'))
            non_sales = classify_order_line(normalized).sales_excluded
            value = amount(row.get('商家实收金额(元)'), source) if paid else None
            if not paid and text(row.get('商家实收金额(元)')) and amount(row.get('商家实收金额(元)'), source) != 0:
                raise ValueError(f'有实收金额但无有效支付时间，不能改用成交/创建时间: {source}')
            record = {'shop': shop, 'order': oid, 'day': paid, 'amount': value, 'non_sales': non_sales,
                      'product': text(row.get('商品id') or row.get('商品ID')), 'sku': text(row.get('样式ID')),
                      'source': source, 'file': str(path)}
            key = shop, oid
            if key in orders:
                previous = orders[key]
                signature = ('day', 'amount', 'non_sales', 'product', 'sku')
                if previous['file'] == str(path) and any(previous[k] != record[k] for k in signature):
                    raise ValueError(f'同一订单出现冲突支付行，不能重复累加: {source}')
                counts['ignored_order_duplicates'] += 1
                continue
            orders[key] = record
        if accepted:
            order_sources.append(str(path))
    order_shop = defaultdict(set)
    for shop, oid in orders:
        order_shop[oid].add(shop)
    for path in discover(root, '售后数据', '3.售后数据库'):
        accepted = False
        for row, source in records(path, lambda h: {'售后编号', '订单编号', '售后状态', '退款金额'}.issubset(h)):
            accepted = True
            rid = text(row.get('售后编号')); oid = text(row.get('订单编号'))
            if not rid or not oid:
                raise ValueError(f'售后缺售后编号或订单编号，无法核对退款事件: {source}')
            shops = order_shop.get(oid, set())
            shop = canonical_shop(row, path, config, next(iter(shops)) if len(shops) == 1 else None)
            status = text(row.get('售后状态'))
            successful = status == '退款成功'
            record = {'shop': shop, 'order': oid, 'refund': rid, 'status': status,
                      'day': day(row.get('同意退款时间')) if successful else None,
                      'applied': day(row.get('申请时间')),
                      'amount': amount(row.get('退款金额'), source) if successful else None,
                      'source': source, 'file': str(path)}
            key = shop, rid
            if key in refunds:
                previous = refunds[key]
                if previous['file'] == str(path) and any(previous[k] != record[k] for k in ('order','status','day','amount')):
                    raise ValueError(f'同一售后编号出现冲突记录: {source}')
                counts['ignored_refund_snapshots'] += 1
                continue
            refunds[key] = record
        if accepted:
            refund_sources.append(str(path))
    payment_days = defaultdict(list); refund_days = defaultdict(list)
    paid_totals = defaultdict(lambda: Decimal(0)); refund_totals = defaultdict(lambda: Decimal(0))
    paid_ids = defaultdict(set); refund_order_ids = defaultdict(set)
    unknown_refund_days = defaultdict(set)
    payment_trace = defaultdict(list); refund_trace = defaultdict(list)
    for key, record in orders.items():
        dt, shop = record['day'], record['shop']
        if dt:
            payment_days[shop].append(dt)
        if record['non_sales'] or key in excluded:
            if dt in dates:
                counts['excluded_paid_orders'] += 1
            continue
        if dt in dates:
            paid_totals[dt, shop] += record['amount']; paid_ids[dt, shop].add(record['order'])
            payment_trace[dt, shop].append({'order': record['order'], 'amount': str(record['amount']), 'source': record['source']})
    for record in refunds.values():
        shop, dt = record['shop'], record['day']
        if record['status'] != '退款成功':
            counts['non_successful_refunds'] += 1; continue
        if dt:
            refund_days[shop].append(dt)
        original = orders.get((shop, record['order']), {})
        if (shop, record['order']) in excluded or original.get('non_sales'):
            if dt in dates:
                counts['excluded_refund_events'] += 1
            continue
        if not dt:
            affected = {d for d in dates if not record['applied'] or d >= record['applied']}
            unknown_refund_days[shop].update(affected)
            logs.append({'source': record['source'], 'message': '退款成功但无同意退款时间，不能分配日期；可能受影响日期退款保持空白'})
        elif dt in dates:
            refund_totals[dt, shop] += record['amount']; refund_order_ids[dt, shop].add(record['order'])
            refund_trace[dt, shop].append({'refund': record['refund'], 'order': record['order'], 'amount': str(record['amount']), 'source': record['source']})
    result = {}
    shops = {s['店铺'] for s in config.get('shops', []) if s['平台'] == '拼多多'} | set(payment_days) | set(refund_days)
    for shop in shops:
        pc = payment_days.get(shop, []); rc = refund_days.get(shop, [])
        for dt in sorted(dates):
            payment_known = bool(pc) and min(pc) <= dt <= max(pc)
            refund_known = bool(rc) and min(rc) <= dt <= max(rc) and dt not in unknown_refund_days[shop]
            result[dt, '拼多多', shop] = {'支付金额': money(paid_totals[dt,shop]) if payment_known else None,
                '支付订单数': len(paid_ids[dt,shop]) if payment_known else None,
                '退款金额': money(refund_totals[dt,shop]) if refund_known else None,
                '退款订单数': len(refund_order_ids[dt,shop]) if refund_known else None,
                '_现金流明细': {'payments': payment_trace[dt,shop], 'refunds': refund_trace[dt,shop]},
                '_现金流规则': POLICY_NOTE}
    return {'rows': result, 'logs': logs, 'counts': dict(counts), 'order_sources': order_sources,
            'refund_sources': refund_sources, 'exclusions_source': exclusion_source}


def apply(rows, ledger, logs, evidence, backend=None, *, platform='拼多多', policy=POLICY_NOTE):
    for row in rows:
        key = row['日期'], row['平台'], row['店铺']
        if row['平台'] != platform:
            continue
        values = ledger['rows'].get(key, {field: None for field in CASH_FIELDS})
        reference = dict((backend or {}).get(key, row))
        checks = []
        for field in CASH_FIELDS:
            original = reference.get(field)
            value = values.get(field)
            checks.append({'field': field, '订单售后计算': value, '后台复核': original,
                           '差异': round(value-original, 2) if value is not None and original is not None else None})
            row[field] = value
            evidence['|'.join((*key, field))] = policy if platform == '天猫' else (
                '订单支付日商家实收汇总' if field.startswith('支付') else '退款成功，按同意退款时间归日（估算）')
        paid, refunded, count = row['支付金额'], row['退款金额'], row['支付订单数']
        net = money(Decimal(str(paid))-Decimal(str(refunded))) if paid is not None and refunded is not None else None
        row['净成交额'] = net
        from profit_rates import net_margin
        row['净利率'] = net_margin(row.get('店铺净利润'), net)
        row['客单价'] = paid/count if paid is not None and count is not None and count>0 else None
        spend = row.get('推广花费')
        row['推广ROI'] = net/spend if net is not None and net>=0 and spend is not None and spend>0 else None
        row['推广费比'] = spend/net if net is not None and net>0 and spend is not None and spend>=0 else None
        reference_net = (reference['支付金额']-reference['退款金额']) if reference.get('支付金额') is not None and reference.get('退款金额') is not None else None
        checks.append({'field':'净成交额','订单售后计算':net,'后台复核':round(reference_net,2) if reference_net is not None else None,
                       '差异':round(net-reference_net,2) if net is not None and reference_net is not None else None})
        for field in CASH_DERIVED:
            evidence['|'.join((*key,field))] = '订单/售后现金流派生；'+policy
        previous = {k:v for k,v in row.get('_原值保留', {}).items() if k in CASH_FIELDS+CASH_DERIVED}
        if previous:
            row.setdefault('_历史现金流保留', {}).update(previous)
            row['_原值保留'] = {k:v for k,v in row['_原值保留'].items() if k not in CASH_FIELDS+CASH_DERIVED}
        old_note = row.get('_口径说明')
        if old_note and policy not in old_note:
            row.setdefault('_历史口径说明', old_note)
        row['_口径说明'] = policy
        row['_现金流规则'] = policy
        row['_后台金额复核'] = checks
        row['_现金流明细'] = values.get('_现金流明细', {})
        logs.append({'date': key[0], 'shop': key[2], 'message': policy, 'checks': checks})
    logs.extend(ledger['logs'])
