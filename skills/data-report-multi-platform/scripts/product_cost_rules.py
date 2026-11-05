"""User-confirmed unit cost and margin rules for product daily reports."""
from typing import Any

from data_quality_protection import optional_number

SHIPPING_PER_UNIT = 2.1
OTHER_COST_RATE = 0.167
PLATFORM_COST_RATE = 0.007
NEW_COST_COLUMNS = ("综合成本", "快递成本", "其他均摊成本", "平台费用")
REMOVED_COLUMNS = ("店铺费用", "平摊管理费用")


def apply_product_cost_rules(row: dict[str, Any]) -> None:
    """Keep cost and gross margin per unit; net profit remains a row total.

    Unknown product cost never becomes zero. Net profit still requires known
    net sales, including any subsidy; missing marketing data cannot imply zero
    subsidy. Empty-burn rows have no fulfillment costs or sales margin.
    """
    for field in REMOVED_COLUMNS:
        row.pop(field, None)
    quantity = optional_number(row.get("实际成交数量（去退款去补单后）"))
    promotion = optional_number(row.get("推广费用")) or 0.0
    if quantity == 0 and "空烧推广费" in str(row.get("备注") or ""):
        row.update({"快递成本": 0.0, "其他均摊成本": 0.0, "平台费用": 0.0,
                    "综合成本": "", "毛利": 0.0, "毛利率": "",
                    "净销售额（实际成交+总补贴金额）": 0.0,
                    "净利润": -promotion, "定价是否合理": "建议缩减推广"})
        return
    arrival = optional_number(row.get("到手价"))
    cost = optional_number(row.get("产品成本"))
    valid_sales = quantity is not None and quantity > 0 and arrival is not None and arrival >= 0
    row["快递成本"] = SHIPPING_PER_UNIT if valid_sales else ""
    row["其他均摊成本"] = arrival * OTHER_COST_RATE if valid_sales else ""
    row["平台费用"] = arrival * PLATFORM_COST_RATE if valid_sales else ""
    row.update({"综合成本": "", "毛利": "", "毛利率": "", "净利润": "", "定价是否合理": ""})
    if not valid_sales or cost is None or cost <= 0:
        return
    total = cost + row["快递成本"] + row["其他均摊成本"] + row["平台费用"]
    gross = arrival - total
    row["综合成本"] = total
    row["毛利"] = gross
    row["毛利率"] = gross / arrival if arrival else ""
    net_sales = optional_number(row.get("净销售额（实际成交+总补贴金额）"))
    if net_sales is None:
        return
    base_profit = net_sales - total * quantity
    row["净利润"] = base_profit - promotion
    row["定价是否合理"] = "建议改价" if base_profit < 0 else "建议缩减推广" if row["净利润"] < 0 else "合理"
