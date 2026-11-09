"""Known store archive headers, including the original 26-column template."""
STORE_HEADERS = ['日期', '平台', '店铺', '填报人', '访客数', '曝光次数', '点击次数', '支付金额', '支付订单数', '支付买家数', '退款金额', '退款订单数', '新客数', '老客数', '老客数占比', '客单价', '支付转化率', '净成交额', '推广花费', '推广成交额', '推广ROI', '推广费比', '店铺评分', '投诉差评数', '今日主要动作', '异常说明', '店铺净利润', '净利率']

def compatible_store_headers(headers):
    return tuple(headers) in (tuple(STORE_HEADERS),tuple(STORE_HEADERS[:27]),tuple(STORE_HEADERS[:26]))
