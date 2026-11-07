"""Public product columns and recognized legacy archive projection."""
PRODUCT_COLUMNS = ('日期', '店铺名称', '商品ID', 'SKU ID', 'SKU编码', 'SKU名称', '综合成本', '产品成本', '快递成本', '其他均摊成本', '平台费用', '到手价', '实际成交数量（去退款去补单后）', '实际成交金额（去退款去补单后）', '毛利', '毛利率', '推广费用', '净利润', '定价是否合理', '备注')
LEGACY_COLUMNS_32 = ('日期', '品牌', '所属运营', '平台', '店铺名称', '大类', '商品ID', 'SKU编码', 'SKU名称', '综合成本', '产品成本', '快递成本', '其他均摊成本', '平台费用', '促销机制        (天猫参加活动填万人团\n淘宝参加活动填百亿补贴)', '到手价', '实际成交数量（去退款去补单后）', '实际成交金额（去退款去补单后）', '毛利', '毛利率', '推广费用', '净利润', '定价是否合理', '备注', '', '大类辅助列', '辅助', '项目组', '管理类型', '品种', '产线', 'SKU ID')
HIDDEN_OUTPUT_COLUMNS = tuple(c for c in LEGACY_COLUMNS_32 if c not in PRODUCT_COLUMNS)

def normalize_headers(headers):
    return [str(v or '').replace('\r\n','\n').replace('\r','\n') for v in headers]

def validate_product_columns(columns):
    if tuple(columns) != PRODUCT_COLUMNS:
        raise ValueError('商品日报必须为新版20列：日期格式yyyy-mm-dd，SKU ID紧随商品ID，备注为最后一列。')

def compatible_headers(headers, columns):
    actual=tuple(normalize_headers(headers));expected=tuple(columns)
    return actual==expected or (expected==PRODUCT_COLUMNS and actual==LEGACY_COLUMNS_32)
