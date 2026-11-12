"""Goods quantities and store profit, derived only after a product report succeeds."""
from pathlib import Path
from collections import defaultdict
import argparse
import json
import math
from goods_catalog import GoodsCatalog, GOODS_COLUMNS, goods_rows
from product_schema import PRODUCT_COLUMNS
from report_archive import read_product, product, product_sort, commit, locked, store, day
from report_exclusions import exclude_existing_sales
from portable_reports import read_workbook_sheets


def load_catalog(root):
    path = Path(root)/'1.产品数据库/产品成本.xlsx'
    if not path.exists():
        # Match the existing legacy database adapter; absence stays visibly unresolved.
        candidates = list((Path(root)/'成本表').glob('*.xlsx'))
        if len(candidates) != 1:
            return GoodsCatalog()
        path = candidates[0]
    from decimal import Decimal
    from data_quality_protection import load_cost_table_records
    costs, _ = load_cost_table_records(path, read_workbook_sheets(path))
    catalog = costs.goods_catalog
    catalog.unit_costs = {code: Decimal(str(row['产品成本'])) for code,row in costs.items()}
    catalog.cost_kinds = {code: row.get('成本性质') for code,row in costs.items()}
    return catalog


def daily_profit(rows):
    grouped = defaultdict(list)
    for row in rows:
        grouped[row['日期'], row['店铺名称']].append(row)
    result = {}
    for key, sources in grouped.items():
        missing = [r.get('SKU编码') or r.get('SKU ID') for r in sources
                   if not isinstance(r.get('净利润'), (int, float)) or not math.isfinite(r['净利润'])]
        # Unknown profit is never a hidden zero. Confirmed temporary zeros are numeric.
        result[key] = {'value': None if missing else math.fsum(r['净利润'] for r in sources),
                       'missing': missing, 'rows': len(sources),
                       'zero_estimates': sorted({r['SKU编码'] for r in sources if '净利润暂计0' in str(r.get('备注') or '')})}
    return result


def store_payload(path):
    """Read visible current values while preserving per-date protected metadata."""
    from openpyxl import load_workbook
    state = json.loads(path.with_suffix('.store.json').read_text(encoding='utf-8'))
    metadata = {(r['日期'], r['平台'], r['店铺']): r for r in state['rows']}
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        sheet = workbook['①平台日报']; sheet.reset_dimensions()
        headers = list(next(sheet.iter_rows(min_row=3, max_row=3, values_only=True)))
        from build_store_report import HEADERS
        from store_schema import compatible_store_headers
        if not compatible_store_headers(headers):
            raise ValueError(f'店铺表头未识别: {path}')
        rows = []
        for values in sheet.iter_rows(min_row=4, values_only=True):
            if not values[0]:
                continue
            row = dict(zip(headers, values)); row['日期'] = day(row['日期'])
            old = metadata.get((row['日期'], row['平台'], row['店铺']), {})
            row.update({k: v for k, v in old.items() if k.startswith('_')})
            rows.append(row)
    finally:
        workbook.close()
    return dict(state, rows=rows, headers=HEADERS)


def refresh(report, database_root, platform, dates, archive_dir, explicit_store=None, migrate=False, goods_only=False):
    report = Path(report); archive_dir = Path(archive_dir); dates = set(dates)
    catalog = load_catalog(database_root)
    months = read_product(report, list(PRODUCT_COLUMNS))
    original = [r for month in sorted(months) for r in months[month]]
    retained, removed = exclude_existing_sales(original, platform)
    # Old final reports may predate an exact product+style code export. Use it
    # solely for goods expansion; preserve the archived financial business key.
    if platform == 'pdd' and any(str(r.get('SKU编码') or '').strip().casefold() not in catalog.records
                                 and float(r.get('实际成交数量（去退款去补单后）') or 0) > 0 for r in retained):
        import build_pdd_report as pdd
        from portable_reports import install
        install(pdd)
        index, _ = pdd.load_product_exports(Path(database_root))
        for row in retained:
            if str(row.get('SKU编码') or '').strip().casefold() in catalog.records or not row.get('SKU ID'):
                continue
            match = index.resolve(row.get('商品ID'), row.get('SKU ID'))
            if match.get('SKU ID') == row.get('SKU ID') and match.get('商品SKU'):
                row['_货品匹配编码'] = match['商品SKU']
                row['_货品编码来源'] = match.get('来源')
    annotated, issues = [], []
    from product_cost_rules import apply_temporary_profit, apply_name_profit_override
    for row in retained:
        apply_temporary_profit(row)
        row, _, error = catalog.annotate(row)
        apply_name_profit_override(row)
        annotated.append(row)
        if error:
            issues.append({'日期': row['日期'], 'SKU编码': row['SKU编码'], '原因': error})
    if not goods_only and (migrate or annotated != original):
        all_dates = sorted({r['日期'] for r in original})
        state = json.loads(report.with_suffix('.product.json').read_text(encoding='utf-8'))
        notes = dict(state.get('date_notes', {}))
        for date in {r['日期'] for r in removed}:
            note = '用户2026-11-09确认：商品991147507035、1021621380194全日期排除销售，推广费用保留实际值'
            if note not in notes.get(date, ''):
                notes[date] = '；'.join(filter(None, [notes.get(date), note]))
        product(report, list(PRODUCT_COLUMNS), annotated, all_dates, notes)
    financial_audit = {}
    goods, goods_issues = goods_rows(annotated, catalog, financial_audit)
    name = '拼多多' if platform == 'pdd' else '天猫'
    goods_output = archive_dir/f'货品日报_{name}.xlsx'
    parts = [{'name': month, 'rows': [r for r in goods if r['日期'][:7] == month]}
             for month in sorted({r['日期'][:7] for r in annotated})]
    payload = {'columns': GOODS_COLUMNS, 'rows': goods, 'sheets': parts, 'kind': 'goods',
               'issues': goods_issues, 'source_product_report': str(report), 'financial_audit': financial_audit,
               'nativeGoodsSource': str(goods_output) if goods_output.exists() else None}
    with locked(goods_output):
        commit(goods_output, payload, 'write_goods_report.mjs')
    profit = daily_profit(annotated)
    config = json.loads((Path(__file__).parent.parent/'references/store_report_config.json').read_text(encoding='utf-8'))
    aliases = config.get('shop_aliases', {})
    profits_by_alias = {(d, aliases.get(s, s)): v for (d, s), v in profit.items()}
    stores = [Path(explicit_store)] if explicit_store else sorted(archive_dir.glob('店铺日报_????-??.xlsx'))
    updated = []
    for path in ([] if goods_only else stores):
        data = store_payload(path)
        changed = False
        for row in data['rows']:
            if row['平台'] != name or (not migrate and row['日期'] not in dates):
                continue
            key = row['日期'], aliases.get(row['店铺'], row['店铺'])
            summary = profits_by_alias.get(key)
            unallocated = row.get('_万相台未归属商品花费') if platform == 'tmall' else None
            if summary and unallocated is not None:
                value = summary['value']
                summary = {**summary, 'unallocated_wanxiang_spend': unallocated,
                           'value': math.fsum((value, -unallocated)) if value is not None else None}
            row['店铺净利润'] = summary['value'] if summary else None
            from profit_rates import net_margin
            row['净利率'] = net_margin(row.get('店铺净利润'), row.get('净成交额'))
            row['_净利润核对'] = summary or {'value': None, 'missing': ['无当日商品日报'], 'rows': 0}
            evidence_key = '|'.join((row['日期'], row['平台'], row['店铺'], '店铺净利润'))
            data.setdefault('evidence', {})[evidence_key] = str(report)+'；同日店铺SKU净利润之和（含空烧推广），不使用拆分货品重复汇总'
            if unallocated:
                data['evidence'][evidence_key] += f'，扣除未归属商品的万相台花费 {unallocated:.2f} 元'
            if not summary or summary['missing'] or summary['zero_estimates'] or unallocated:
                data.setdefault('logs', []).append({'date': row['日期'], 'shop': row['店铺'], 'message':
                    '店铺净利润口径: '+json.dumps(row['_净利润核对'], ensure_ascii=False)})
            changed = True
        if changed:
            store(data, archive_dir, path); updated.append(str(path))
    audit = {'product_report': str(report), 'goods_report': str(goods_output), 'product_rows': len(annotated),
             'goods_rows': len(goods), 'goods_issues': goods_issues, 'financial_audit': financial_audit, 'removed_sales_rows': removed,
             'daily_profit': [{'日期': d, '店铺': s, **v} for (d, s), v in sorted(profit.items())], 'stores': updated}
    audit_path = goods_output.with_name(goods_output.stem+'_财务分摊核对.json') if goods_only else report.with_name(report.stem+'_货品与利润核对.json')
    # Preserve prior exclusion evidence after an idempotent rerun.
    if audit_path.exists():
        previous = json.loads(audit_path.read_text(encoding='utf-8')).get('removed_sales_rows', [])
        audit['removed_sales_rows'] = previous + [r for r in removed if r not in previous]
    audit_path.write_text(json.dumps(audit, ensure_ascii=False, indent=2), encoding='utf-8')
    print(f'货品日报: {goods_output}; 行数={len(goods)}; 待拆分={len(goods_issues)}；店铺净利润已更新{len(updated)}份月报')
    return audit


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--product', required=True, type=Path)
    parser.add_argument('--database-root', required=True, type=Path)
    parser.add_argument('--archive-dir', required=True, type=Path)
    parser.add_argument('--platform', choices=['pdd', 'tmall'], required=True)
    parser.add_argument('--migrate-existing', action='store_true')
    args = parser.parse_args()
    refresh(args.product, args.database_root, args.platform, set(), args.archive_dir, migrate=args.migrate_existing)
