"""User-confirmed unit cost and margin rules for product daily reports."""
from typing import Any

from data_quality_protection import optional_number, normalize_match_code, text

OTHER_COST_RATE = 0.167
PLATFORM_COST_RATE = 0.007
NEW_COST_COLUMNS = ("综合成本", "快递成本", "其他均摊成本", "平台费用")
REMOVED_COLUMNS = ("店铺费用", "平摊管理费用")


def apply_product_cost_rules(row: dict[str, Any]) -> None:
    """Apply the requested net-profit definition: gross minus promotion spend.

    Do not multiply gross by quantity or depend on unknown subsidy/net sales.
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
                    "净销售额（实际成交+总补贴金额）": 0.0,
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
        return
    total = cost + row["快递成本"] + row["其他均摊成本"] + row["平台费用"]
    gross = arrival - total
    row["综合成本"] = total
    row["毛利"] = gross
    row["毛利率"] = gross / arrival if arrival else ""
    row["净利润"] = gross - promotion
    row["定价是否合理"] = "建议改价" if gross < 0 else "建议缩减推广" if row["净利润"] < 0 else "合理"


def product_output_row(row: dict[str, Any], columns: list[str], costs) -> dict[str, Any]:
    """Rename the report field without changing internal export/matching aliases."""
    code = text(row.get("SKU编码") or row.get("商品SKU"))
    key = normalize_match_code(code)
    names = getattr(costs, "sku_names", {})
    name = names.get(key) or costs.get(key, {}).get("SKU名称", "")
    output = {column: row.get(column, "") for column in columns}
    output["SKU编码"] = code
    output["SKU名称"] = name
    return output
