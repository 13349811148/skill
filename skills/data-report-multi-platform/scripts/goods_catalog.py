"""Exact code BOM expansion, independent of whether a monetary cost exists."""
from collections import Counter
from decimal import Decimal, InvalidOperation
import re
from data_quality_protection import normalize_match_code, text

SALES_QUANTITY = '实际成交数量（去退款去补单后）'
GOODS_COLUMNS = ['日期', '店铺名称', '货品编码', '货品名称', '基础单位', '货品数量', '来源SKU及货品数量', '备注']


class GoodsCatalog:
    def __init__(self, records=None):
        self.records = records or {}

    @classmethod
    def from_sheets(cls, sheets):
        records = {}
        for sheet in sheets:
            for pos, raw in enumerate(sheet.rows[:20]):
                headers = [text(v) for v in raw]
                if not {'编码类型', '组合装编码', '货品编码（组合内）', '产品名称'}.issubset(headers):
                    continue
                for values in sheet.rows[pos + 1:]:
                    row = dict(zip(headers, values))
                    kind = text(row.get('编码类型'))
                    if kind not in {'组合装', '货品'}:
                        continue
                    code = text(row.get('组合装编码') if kind == '组合装' else row.get('货品编码（组合内）'))
                    key = normalize_match_code(code)
                    if not key:
                        continue
                    record = {'code': code, 'name': text(row.get('产品名称')), 'unit': text(row.get('基础单位')),
                              'kind': kind, 'bom': text(row.get('货品编码（组合内）')) if kind == '组合装' else ''}
                    if key in records and records[key] != record:
                        raise ValueError(f'货品编码或组合装定义冲突: {code}')
                    records[key] = record
                break
        return cls(records)

    def expand(self, code, ancestors=()):
        key = normalize_match_code(code)
        if key in ancestors:
            raise ValueError(f'组合装循环引用: {" -> ".join((*ancestors, key))}')
        record = self.records.get(key)
        if not record or not record['name']:
            raise ValueError(f'缺少货品定义: {code}')
        if record['kind'] == '货品':
            return {key: Decimal(1)}
        result = Counter()
        for token in re.split('[；;]', record['bom']):
            match = re.fullmatch(r'\s*(.+?)\s*[×*]\s*(\d+(?:\.\d+)?)\s*', token)
            if not match:
                raise ValueError(f'组合装子件格式不完整: {record["code"]}/{token}')
            count = Decimal(match[2])
            if count <= 0:
                raise ValueError(f'组合装子件数量必须为正: {record["code"]}')
            for child, quantity in self.expand(match[1], (*ancestors, key)).items():
                result[child] += quantity * count
        return dict(sorted(result.items()))

    def annotate(self, row):
        result = dict(row)
        result.update({'货品名称': '', '货品数量': ''})
        try:
            quantity = Decimal(str(row.get(SALES_QUANTITY) or 0))
            if not quantity.is_finite() or quantity < 0:
                raise ValueError('有效销售件数必须为非负数')
            if quantity == 0:
                return result, {}, None
            parts = self.expand(row.get('_货品匹配编码') or row.get('SKU编码') or row.get('商品SKU') or '')
            names, counts = [], []
            for key, factor in parts.items():
                names.append(self.records[key]['name'])
                counts.append(format(quantity * factor, 'f').rstrip('0').rstrip('.') if quantity * factor % 1 else str(int(quantity * factor)))
            result['货品名称'] = '\n'.join(names)
            result['货品数量'] = float(quantity * next(iter(parts.values()))) if len(parts) == 1 else '\n'.join(counts)
            return result, {key: quantity * factor for key, factor in parts.items()}, None
        except (ValueError, InvalidOperation) as error:
            return result, {}, str(error)


def goods_rows(product_rows, catalog):
    grouped, issues = {}, []
    for source in product_rows:
        _, parts, error = catalog.annotate(source)
        if error:
            issues.append({'日期': source['日期'], '店铺名称': source['店铺名称'], '商品ID': source.get('商品ID'),
                           'SKU ID': source.get('SKU ID'), 'SKU编码': source['SKU编码'], '原因': error})
            # Keep the unresolved SKU visible; never pretend its SKU count is a goods count.
            identity = str(source.get('商品ID') or '') + '/' + str(source.get('SKU ID') or source['SKU编码'])
            key = (source['日期'], source['店铺名称'], '待拆分:' + identity)
            target = grouped.setdefault(key, {'日期': key[0], '店铺名称': key[1], '货品编码': key[2], '货品名称': '',
                '基础单位': '', '货品数量': None, '_sources': Counter(), '备注': error+'；来源列为未拆分SKU销量'})
            target['_sources'][identity] += Decimal(str(source.get(SALES_QUANTITY) or 0))
        for code, count in parts.items():
            record = catalog.records[code]
            key = (source['日期'], source['店铺名称'], code)
            target = grouped.setdefault(key, {'日期': key[0], '店铺名称': key[1], '货品编码': record['code'],
                '货品名称': record['name'], '基础单位': record['unit'], '货品数量': Decimal(0), '_sources': Counter(),
                '备注': '已扣退款与补单；按货品基础单位计数'})
            target['货品数量'] += count
            target['_sources'][str(source.get('_货品匹配编码') or source['SKU编码'])] += count
    rows = []
    for key in sorted(grouped):
        row = grouped[key]
        row['来源SKU及货品数量'] = '\n'.join(f'{code} × {count:g}' for code, count in sorted(row.pop('_sources').items()))
        if row['货品数量'] is not None:
            row['货品数量'] = float(row['货品数量'])
        rows.append(row)
    return rows, issues
