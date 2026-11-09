"""User-confirmed all-date sales exclusions; promotion is deliberately separate."""
PDD_EXCLUDED_SALES_PRODUCTS = frozenset({'991147507035', '1021621380194'})

def excluded_pdd_product(value):
    return str(value or '').strip() in PDD_EXCLUDED_SALES_PRODUCTS

def exclude_existing_sales(rows, platform):
    retained, removed = [], []
    for source in rows:
        if platform != 'pdd' or not excluded_pdd_product(source.get('商品ID')):
            retained.append(dict(source)); continue
        qty = float(source.get('实际成交数量（去退款去补单后）') or 0)
        if qty == 0 and '空烧推广费' in str(source.get('备注') or ''):
            retained.append(dict(source)); continue
        removed.append(dict(source))
        promotion = float(source.get('推广费用') or 0)
        if promotion:
            row = dict(source)
            row.update({'实际成交数量（去退款去补单后）': 0, '实际成交金额（去退款去补单后）': 0,
                        '到手价': '不适用（无有效销量）', '综合成本': '', '快递成本': 0,
                        '其他均摊成本': 0, '平台费用': 0, '毛利': 0, '毛利率': '',
                        '净利润': -promotion, '定价是否合理': '建议缩减推广',
                        '备注': str(source.get('备注') or '')+'；空烧推广费：商品ID销售已按用户要求剔除，推广按实际保留'})
            retained.append(row)
    return retained, removed
