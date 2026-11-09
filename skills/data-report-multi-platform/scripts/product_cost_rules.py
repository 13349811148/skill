"""Per-unit comprehensive cost and whole-row gross/net profit."""
from typing import Any
from product_schema import HIDDEN_OUTPUT_COLUMNS

from data_quality_protection import optional_number, normalize_match_code, text

OTHER_COST_RATE = 0.167
PLATFORM_COST_RATE = 0.007
NEW_COST_COLUMNS = ("综合成本", "快递成本", "其他均摊成本", "平台费用")
REMOVED_COLUMNS = ("店铺费用", "平摊管理费用")
REMOVED_REPORT_COLUMNS = ("活动价", "报名价", "每单补贴金额", "总补贴金额",
                          "净销售额（实际成交+总补贴金额）") + HIDDEN_OUTPUT_COLUMNS


TEMPORARY_ZERO_PROFIT_CODES = frozenset({'zb03', 'zb27', 'cb10'})

def apply_temporary_profit(row):
    code = normalize_match_code(row.get('SKU编码') or row.get('商品SKU'))
    cost = optional_number(row.get('产品成本'))
    quantity = optional_number(row.get('实际成交数量（去退款去补单后）'))
    if code in TEMPORARY_ZERO_PROFIT_CODES and (cost is None or cost <= 0) and quantity and quantity > 0:
        row['净利润'] = 0.0
        note = '用户2026-11-09确认：缺产品成本期间净利润暂计0，成本与毛利留空，补齐成本后恢复正常计算'
        if note not in str(row.get('备注') or ''):
            row['备注'] = '；'.join(filter(None, [row.get('备注'), note]))


def apply_product_cost_rules(row: dict[str, Any]) -> None:
    """Comprehensive cost is per unit; gross deducts it times quantity.

    Gross = actual row receipt minus total cost; net = gross minus promotion.
    Subsidy is not part of the new profit definition.
    Unknown product or shipping cost never becomes zero. Shipping is supplied
    by the freight matcher and survives a second partial-reference pass.
    """
    for field in REMOVED_COLUMNS:
        row.pop(field, None)
    quantity = optional_number(row.get("实际成交数量（去退款去补单后）"))
    promotion = optional_number(row.get("推广费用")) or 0.0
    if quantity == 0 and "空烧推广费" in str(row.get("备注") or ""):
        row.update({"快递成本": 0.0, "其他均摊成本": 0.0, "平台费用": 0.0,
                    "综合成本": "", "毛利": 0.0, "毛利率": "",
                    "净利润": -promotion, "定价是否合理": "建议缩减推广"})
        return
    arrival = optional_number(row.get("到手价"))
    cost = optional_number(row.get("产品成本"))
    valid_sales = quantity is not None and quantity > 0 and arrival is not None and arrival >= 0
    shipping = optional_number(row.get("快递成本"))
    row["快递成本"] = shipping if valid_sales and shipping is not None and shipping > 0 else ""
    row["其他均摊成本"] = arrival * OTHER_COST_RATE if valid_sales else ""
    row["平台费用"] = arrival * PLATFORM_COST_RATE if valid_sales else ""
    row.update({"综合成本": "", "毛利": "", "毛利率": "", "净利润": "", "定价是否合理": ""})
    if not valid_sales or cost is None or cost <= 0 or row["快递成本"] == "":
        apply_temporary_profit(row)
        return
    total = cost + row["快递成本"] + row["其他均摊成本"] + row["平台费用"]
    row["综合成本"] = total
    amount = optional_number(row.get("实际成交金额（去退款去补单后）"))
    if amount is None:
        return
    gross = amount - row["综合成本"] * quantity
    row["毛利"] = gross
    row["毛利率"] = gross / amount if amount else ""
    row["净利润"] = gross - promotion
    row["定价是否合理"] = "建议改价" if gross < 0 else "建议缩减推广" if row["净利润"] < 0 else "合理"


def product_output_row(row: dict[str, Any], columns: list[str], costs) -> dict[str, Any]:
    """Rename the report field without changing internal export/matching aliases."""
    code = text(row.get("SKU编码") or row.get("商品SKU"))
    key = normalize_match_code(code)
    names = getattr(costs, "sku_names", {})
    name = names.get(key) or costs.get(key, {}).get("SKU名称", "")
    output = {column: row.get(column, "") for column in columns
              if column not in REMOVED_REPORT_COLUMNS}
    output["SKU编码"] = code
    output["SKU名称"] = name
    catalog = getattr(costs, 'goods_catalog', None)
    if catalog is not None and {'货品名称', '货品数量'}.issubset(columns):
        output, _, _ = catalog.annotate(output)
    apply_temporary_profit(output)
    return output
