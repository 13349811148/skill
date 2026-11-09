"""Attribute SKU money once to leaf goods, keeping unknowns and empty burn visible."""
from collections import defaultdict
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from data_quality_protection import normalize_match_code

METRICS = ('销售额', '毛利', '推广费用', '净利润')
SOURCE_FIELDS = {'销售额': '实际成交金额（去退款去补单后）', '毛利': '毛利', '推广费用': '推广费用', '净利润': '净利润'}
PRICE = '客单价（货品均价）'
CENT = Decimal('0.01')


def numeric(value, field):
    if value is None or value == '':
        return Decimal(0) if field == '推广费用' else None
    try:
        result = Decimal(str(value))
    except InvalidOperation as error:
        raise ValueError(f'货品分摊收到无效金额: {field}={value}') from error
    if not result.is_finite():
        raise ValueError(f'货品分摊收到非有限金额: {field}={value}')
    return result


def unresolved_code(source):
    identity = str(source.get('商品ID') or '') + '/' + str(source.get('SKU ID') or source['SKU编码'])
    return '待拆分:' + identity


def allocate_financials(quantity_rows, sources, catalog, audit=None):
    audit = audit if audit is not None else {}
    audit.update({'allocations': [], 'unallocated': [], 'daily_reconciliation': []})
    rows = {(r['日期'], r['店铺名称'], normalize_match_code(r['货品编码'])): dict(r) for r in quantity_rows}
    sums = defaultdict(lambda: {m: Decimal(0) for m in METRICS})
    missing = defaultdict(set)
    notes = defaultdict(set)
    day_inputs = defaultdict(lambda: {m: [] for m in METRICS})
    day_incomplete = defaultdict(set)
    for source in sources:
        date, shop = source['日期'], source['店铺名称']
        origin = {k: source.get(k) for k in ['日期', '店铺名称', '商品ID', 'SKU ID', 'SKU编码']}
        money = {m: numeric(source.get(SOURCE_FIELDS[m]), m) for m in METRICS}
        for metric, value in money.items():
            day_inputs[date, shop][metric].append(value)
        quantity = Decimal(str(source.get('实际成交数量（去退款去补单后）') or 0))
        _, parts, error = catalog.annotate(source)
        mode = '单一货品全额归属'
        weights = None
        if quantity == 0:
            if not any(v not in (None, 0) for v in money.values()):
                continue
            identity = '/'.join(str(source.get(k) or '') for k in ['商品ID', 'SKU ID', 'SKU编码'])
            code = '空烧推广费:' + identity
            key = date, shop, normalize_match_code(code)
            rows[key] = {'日期': date, '店铺名称': shop, '货品编码': code,
                         '货品名称': '空烧推广费（无有效销售）', '基础单位': '', '货品数量': 0,
                         '来源SKU及货品数量': str(source.get('SKU编码') or '') + ' × 0',
                         '备注': '无有效销量的推广费用单列，未转摊给销售货品；' + str(source.get('备注') or '')}
            weights = {key[2]: Decimal(1)}
            mode = '空烧费用独立保留'
        elif error:
            key = date, shop, normalize_match_code(unresolved_code(source))
            if key not in rows:
                raise ValueError(f'待拆分货品记录未保留: {origin}')
            weights = {key[2]: Decimal(1)}
            mode = '待拆分SKU金额暂存，不视为已识别货品'
        elif len(parts) == 1:
            weights = {code: Decimal(1) for code in parts}
        else:
            costs = getattr(catalog, 'unit_costs', {})
            absent = [catalog.records[c]['code'] for c in parts if costs.get(c) is None or costs[c] <= 0]
            if absent:
                for code in parts:
                    key = date, shop, code
                    missing[key].update(METRICS)
                    notes[key].add('包含成本不齐的混合装，金额留空；缺成本货品: ' + '、'.join(absent))
                day_incomplete[date, shop].update(METRICS)
                audit['unallocated'].append({'source': origin, 'reason': '混合装子货品成本不齐',
                    'missing_cost_codes': absent, 'amounts': {m: str(v) if v is not None else None for m,v in money.items()}})
                continue
            values = {code: costs[code] * count for code, count in parts.items()}
            total = sum(values.values())
            weights = {code: value / total for code, value in values.items()}
            mode = '按子货品成本×数量占比分摊'
        allocation = {'source': origin, 'mode': mode, 'weights': {c: str(v) for c,v in weights.items()}}
        audit['allocations'].append(allocation)
        for code, weight in weights.items():
            key = date, shop, code
            if key not in rows:
                raise ValueError(f'金额分摊目标货品不存在: {origin}/{code}')
            for metric, value in money.items():
                if value is None:
                    missing[key].add(metric)
                    notes[key].add(metric + '含未确认金额，保持空白')
                else:
                    sums[key][metric] += value * weight
            notes[key].add(mode)
            if '净利润暂计0' in str(source.get('备注') or ''):
                notes[key].add('包含用户确认的缺成本期间净利润暂计0')
            if mode.startswith('按子货品') and any(getattr(catalog, 'cost_kinds', {}).get(c) == '参考测算' for c in parts):
                notes[key].add('分摊权重包含参考测算成本')
    grouped = defaultdict(list)
    for key in rows:
        grouped[key[:2]].append(key)
    for day, keys in grouped.items():
        keys.sort()
        for metric in METRICS:
            known = [k for k in keys if metric not in missing[k]]
            rounded = {k: sums[k][metric].quantize(CENT, rounding=ROUND_HALF_UP) for k in known}
            complete = (metric not in day_incomplete[day] and len(known) == len(keys)
                        and all(v is not None for v in day_inputs[day][metric]))
            expected = None
            if complete:
                expected = sum(day_inputs[day][metric], Decimal(0)).quantize(CENT, rounding=ROUND_HALF_UP)
                delta = expected - sum(rounded.values(), Decimal(0))
                steps = int(abs(delta / CENT))
                # Preserve the day total after display rounding. Tail cents follow
                # the largest discarded fraction; identifiers give a stable tie break.
                ordered = sorted(known, key=lambda k: (-(sums[k][metric]-rounded[k]) if delta > 0 else sums[k][metric]-rounded[k], k))
                if steps and not ordered:
                    raise ValueError('没有可核对的货品金额行')
                for i in range(steps):
                    rounded[ordered[i % len(ordered)]] += CENT if delta > 0 else -CENT
                assert sum(rounded.values(), Decimal(0)) == expected
            for key in keys:
                rows[key][metric] = float(rounded[key]) if key in rounded else None
            audit['daily_reconciliation'].append({'日期': day[0], '店铺': day[1], 'metric': metric,
                'complete': complete, 'source_total': str(expected) if expected is not None else None,
                'output_known_subtotal': str(sum(rounded.values(), Decimal(0))),
                'missing_goods_rows': len(keys)-len(known)})
    result = []
    for key in sorted(rows, key=lambda k: (k[2].startswith('空烧推广费:'), *k)):
        row = rows[key]
        count, sales = row.get('货品数量'), row.get('销售额')
        row[PRICE] = sales/count if sales is not None and count is not None and count > 0 else None
        from profit_rates import net_margin
        row['净利率'] = net_margin(row.get('净利润'), sales)
        row['备注'] = '；'.join(filter(None, [row.get('备注'), *sorted(notes[key])]))
        result.append(row)
    return result
