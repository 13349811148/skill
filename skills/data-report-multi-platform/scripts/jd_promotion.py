"""京准通现金、红包流水和SKU推广明细的分日核对。"""

from collections import defaultdict
from datetime import date, datetime
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from pathlib import Path
import json
import os
import sys

from portable_reports import read_workbook_sheets

CASH = '京准通现金消费'
VOUCHER = '京准通红包消费'
DETAIL = '京准通推广明细'
FINANCE_HEADERS = {'流水单号', '投放账户ID', '投放日期', '交易类型', '支出'}
DETAIL_HEADERS = {'日期', '产品线', '单元ID', '创意ID', '智能投放推广SKU ID', '展现数', '点击数', '花费'}
CENT = Decimal('0.01')
DEPENDENCY_DIRECTORY = Path(os.environ.get('LOCALAPPDATA', str(Path.home()))) / 'Codex' / 'daily-report' / 'python-deps'


def as_day(value):
    if isinstance(value, (date, datetime)):
        return value.strftime('%Y-%m-%d')
    text = str(value or '').strip()
    if len(text) == 8 and text.isdigit():
        return datetime.strptime(text, '%Y%m%d').date().isoformat()
    return date.fromisoformat(text[:10].replace('/', '-')).isoformat()


def as_money(value, context):
    try:
        amount = Decimal(str(value).strip().replace(',', ''))
    except InvalidOperation as error:
        raise ValueError(f'{context}金额无效') from error
    if not amount.is_finite() or amount < 0:
        raise ValueError(f'{context}金额必须为有限非负数')
    return amount


def identifier(value):
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value or '').strip().strip("'")


def archives(root, folder, kinds):
    base = Path(root) / folder
    for sidecar in sorted(base.rglob('*.archive.json')):
        metadata = json.loads(sidecar.read_text(encoding='utf-8-sig'))
        if metadata.get('platform') != '京东' or metadata.get('archive_category') not in kinds:
            continue
        path = sidecar.with_name(sidecar.name.removesuffix('.archive.json'))
        shop = metadata.get('shop')
        if not shop or not path.is_file():
            raise ValueError(f'京东归档缺少店铺或原表：{sidecar}')
        yield path, shop, metadata


def is_jd_archive(path):
    sidecar = path.with_suffix(path.suffix + '.archive.json')
    return sidecar.is_file() and json.loads(sidecar.read_text(encoding='utf-8-sig')).get('platform') == '京东'


def records(path, required, *, stream=None):
    found = False
    try:
        sheets = read_workbook_sheets(path, stream=stream, merge_duplicate_headers=False)
    except ModuleNotFoundError as error:
        if error.name == 'xlrd':
            raise ValueError('京准通 .xls 财务原表需要 Python xlrd 读取组件') from error
        raise
    for sheet in sheets:
        for index, row in enumerate(sheet.rows[:20]):
            headers = [str(value or '').strip() for value in row]
            if not required.issubset(headers):
                continue
            if any(headers.count(field) > 1 for field in required):
                raise ValueError(f'京东原表必要字段重复：{path}/{sheet.sheet_name}')
            found = True
            for number, values in enumerate(sheet.rows[index+1:], index+2):
                if any(value not in (None, '') for value in values):
                    yield dict(zip(headers, values)), f'{path.name}/{sheet.sheet_name}/{number}'
            break
    if not found:
        raise ValueError(f'京东原表缺少必需表头：{path}（{", ".join(sorted(required))}）')


def enable_local_dependencies():
    if DEPENDENCY_DIRECTORY.is_dir() and str(DEPENDENCY_DIRECTORY) not in sys.path:
        sys.path.append(str(DEPENDENCY_DIRECTORY))


def load(root, dates):
    enable_local_dependencies()
    transactions = {}
    detail_days = {}
    coverage = defaultdict(set)
    paths = []
    for path, shop, metadata in archives(root, '4.推广数据库', {CASH, VOUCHER, DETAIL}):
        kind = metadata['archive_category']
        paths.append(str(path))
        if kind == DETAIL:
            batch = defaultdict(list)
            for row, source in records(path, DETAIL_HEADERS):
                day = as_day(row['日期'])
                if day not in dates:
                    continue
                channel = str(row['产品线']).strip()
                sku = identifier(row['智能投放推广SKU ID'])
                if not channel or not sku:
                    raise ValueError(f'京准通推广明细缺少产品线或SKU：{source}')
                values = (as_money(row['花费'], source), int(row['展现数']), int(row['点击数']))
                if min(values[1:]) < 0:
                    raise ValueError(f'京准通推广明细曝光或点击为负：{source}')
                batch[day].append((channel, sku, *values))
            rank = (path.stat().st_mtime_ns, str(path))
            for day, entries in batch.items():
                key = shop, day
                if key not in detail_days or rank > detail_days[key][0]:
                    detail_days[key] = (rank, entries)
            continue

        required = FINANCE_HEADERS | ({'结算时间', '结余', '开票方式'} if kind == CASH else set())
        for row, source in records(path, required):
            day = as_day(row['投放日期'])
            flow = identifier(row['流水单号'])
            category = str(row['交易类型']).strip()
            if not flow or not category.endswith('扣费'):
                raise ValueError(f'京准通财务原表含未识别流水：{source}')
            key = shop, kind, flow
            transaction = day, category.removesuffix('扣费'), as_money(row['支出'], source)
            previous = transactions.setdefault(key, transaction)
            if previous != transaction:
                raise ValueError(f'京准通流水重复且字段冲突：{source}')
        period = metadata.get('period', '').split('至')
        if len(period) == 2:
            start, end = map(as_day, period)
            coverage[shop, kind].update(day for day in dates if start <= day <= end)

    ledger = defaultdict(lambda: defaultdict(Decimal))
    shops = {shop for shop, _ in coverage} | {shop for shop, _ in detail_days}
    for (shop, kind, _), (day, channel, amount) in transactions.items():
        shops.add(shop)
        if day in dates:
            coverage[shop, kind].add(day)
            ledger[shop, day, channel][kind] += amount
    for (shop, day), (_, rows) in detail_days.items():
        for channel, *_ in rows:
            ledger[shop, day, channel]

    result = {}
    for shop in sorted(shops):
        for day in sorted(dates):
            cash_known = day in coverage[shop, CASH]
            detail = detail_days.get((shop, day), (None, []))[1]
            sku_weights = defaultdict(lambda: defaultdict(Decimal))
            for channel, sku, amount, _, _ in detail:
                sku_weights[channel][sku] += amount
            allocated = defaultdict(Decimal)
            unassigned = Decimal(0)
            cash_total = Decimal(0)
            voucher_total = Decimal(0)
            channel_audit = []
            for (ledger_shop, ledger_day, channel), amounts in sorted(ledger.items()):
                if (ledger_shop, ledger_day) != (shop, day):
                    continue
                cash, voucher = amounts[CASH], amounts[VOUCHER]
                cash_total += cash
                voucher_total += voucher
                weights = sku_weights[channel]
                detail_total = sum(weights.values(), Decimal(0))
                matched = cash_known and detail_total > 0 and detail_total == cash + voucher
                if matched:
                    shares = {sku: (cash * weight / detail_total).quantize(CENT, rounding=ROUND_HALF_UP)
                              for sku, weight in weights.items() if weight}
                    tail = cash - sum(shares.values(), Decimal(0))
                    shares[max(shares, key=lambda sku: (weights[sku], sku))] += tail
                    for sku, share in shares.items():
                        allocated[sku] += share
                else:
                    unassigned += cash
                channel_audit.append({'渠道': channel, '现金': str(cash), '红包流水合计': str(voucher),
                                      'SKU明细': str(detail_total), '已分摊': matched})
            assert sum(allocated.values(), Decimal(0)) + unassigned == cash_total
            result[shop, day] = {
                'cash': cash_total if cash_known else None,
                'voucher': voucher_total if day in coverage[shop, VOUCHER] else None,
                'allocated': dict(allocated), 'unassigned': unassigned if cash_known else None,
                'impressions': sum(row[3] for row in detail) if detail else None,
                'clicks': sum(row[4] for row in detail) if detail else None,
                'channels': channel_audit,
            }
    return result, paths
