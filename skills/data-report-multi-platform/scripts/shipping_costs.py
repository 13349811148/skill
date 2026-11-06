"""Read SKU shipping estimates from the maintained freight workbook.

Use saved values, not a second implementation of the carrier's pricing rules.
Heilongjiang, Jilin and Liaoning use YTO; other regions use the cheaper valid
YTO/China Post quote. This is a per-unit estimate, not an actual parcel bill.
"""
from dataclasses import dataclass, field
from pathlib import Path
import math

from data_quality_protection import DataQualityError, normalize_match_code, optional_number, text


CARRIER_SHEETS = {"圆通": "圆通各地运费", "邮政": "邮政各地运费"}
NORTHEAST = {"黑龙江", "吉林", "辽宁"}
PROVINCES = set("湖北 浙江 江苏 安徽 江西 广东 山东 湖南 河南 福建 天津 河北 陕西 山西 广西 重庆 四川 贵州 云南 辽宁 吉林 黑龙江 上海 北京 内蒙古 甘肃 海南 青海 宁夏 新疆 西藏".split())
PROVINCE_HEADERS = ("收货省份", "收货省", "收件省份", "收件省", "收货人省份", "省份", "省")
CITY_HEADERS = ("收货城市", "收货市", "收件城市", "收件市", "收货人城市", "城市", "市")
QUANTITY = "实际成交数量（去退款去补单后）"
POLICY = "黑吉辽固定圆通；其他地区取圆通、邮政有效报价较低者"


def normalize_province(value) -> str:
    value = normalize_match_code(value)
    aliases = {"黑": "黑龙江", "吉": "吉林", "辽": "辽宁",
               "广西壮族自治区": "广西", "宁夏回族自治区": "宁夏",
               "新疆维吾尔自治区": "新疆", "内蒙古自治区": "内蒙古", "西藏自治区": "西藏"}
    value = aliases.get(value, value)
    if value.endswith(("省", "市")):
        value = value[:-1]
    return value if value in PROVINCES else ""


def destination_region(province, city) -> str:
    normalized = normalize_province(province)
    city = normalize_match_code(city).removesuffix("市")
    if city == "深圳":
        return "深圳" if normalized in ("", "广东") and (normalized or not text(province)) else ""
    if normalized == "广东" and not city:
        return ""  # Shenzhen has a separate quote; the city is needed to distinguish it.
    if not normalized and not text(province) and city in {"北京", "上海", "天津", "重庆"}:
        return city
    return normalized


@dataclass(frozen=True)
class ShippingQuote:
    amount: float | None
    source: str
    reason: str = ""


@dataclass
class ShippingTable:
    prices: dict = field(default_factory=dict)
    selected_region: str = ""

    def lookup(self, code: str, carrier: str, region: str) -> ShippingQuote:
        key = normalize_match_code(code)
        if not key:
            return ShippingQuote(None, "", "SKU编码为空")
        if carrier not in CARRIER_SHEETS:
            return ShippingQuote(None, "", "快递公司未明确或报价表不支持")
        if not region:
            return ShippingQuote(None, "", "收货地区未明确")
        return self.prices.get((key, carrier, region),
                               ShippingQuote(None, "", "运费表未匹配编码或地区"))

    def select(self, code: str, province: str, city: str = "") -> ShippingQuote:
        region = destination_region(province, city)
        if not region:
            return ShippingQuote(None, "", "收货省市缺失、冲突或无法识别")
        yto = self.lookup(code, "圆通", region)
        if region in NORTHEAST:
            return ShippingQuote(yto.amount, "黑吉辽固定圆通；" + yto.source, yto.reason)
        postal = self.lookup(code, "邮政", region)
        valid = [(q.amount, carrier, q) for carrier, q in (("圆通", yto), ("邮政", postal))
                 if q.amount is not None]
        if not valid:
            return ShippingQuote(None, f"{yto.source}；{postal.source}",
                                 f"圆通：{yto.reason}；邮政：{postal.reason}")
        # Stable carrier order makes equal-price selections reproducible.
        amount, carrier, selected = min(valid, key=lambda item: item[0])
        comparison = f"圆通={yto.amount if yto.amount is not None else '无有效报价'}，邮政={postal.amount if postal.amount is not None else '无有效报价'}"
        return ShippingQuote(amount, f"其他地区取低价（{comparison}，选{carrier}）；{selected.source}")


def capture_shipping_order(aggregate: dict, order: dict, quantity: float,
                           code: str, source: str) -> None:
    """Called only after the normal exclusion/review/dedup gates accept a line."""
    if quantity <= 0:
        return
    province = next((text(order.get(h)) for h in PROVINCE_HEADERS if text(order.get(h))), "")
    city = next((text(order.get(h)) for h in CITY_HEADERS if text(order.get(h))), "")
    key = (normalize_match_code(code), province, city)
    contexts = aggregate.setdefault("_运费订单", {})
    context = contexts.setdefault(key, {"编码": code, "省": province, "市": city,
                                        "数量": 0.0, "来源示例": source})
    context["数量"] += quantity


def apply_shipping_cost(row: dict, aggregate: dict, table: ShippingTable,
                        warnings: list[str], cost_log: dict) -> None:
    """Quantity-weighted unit cost; any unmatched effective units stay unknown."""
    row["快递成本"] = ""
    quantity = optional_number(row.get(QUANTITY))
    if quantity is None or quantity <= 0:
        cost_log.update({"快递成本": "", "运费来源": "", "运费匹配状态": "无有效销量"})
        return
    code = text(row.get("SKU编码") or row.get("商品SKU"))
    contexts = aggregate.get("_运费订单", {}).values()
    total_cost = matched = captured = 0.0
    sources = []
    missing = []
    for context in contexts:
        count = context["数量"]
        captured += count
        context_code = context["编码"] or code
        if normalize_match_code(context_code) != normalize_match_code(code):
            quote = ShippingQuote(None, "", "同一报表行的订单编码与最终SKU编码冲突")
        else:
            quote = table.select(context_code, context["省"], context["市"])
        sources.append(f"{context['省']}/{context['市']}×{count:g}件：{quote.source or quote.reason}")
        if quote.amount is None:
            missing.append(f"{context_code or '编码为空'}/{context['省']}/{context['市']}×{count:g}件；"
                           f"{quote.reason}；订单来源={context['来源示例']}")
        else:
            matched += count
            total_cost += count * quote.amount
    if not math.isclose(captured, quantity, rel_tol=1e-9, abs_tol=1e-6):
        missing.append(f"运费上下文数量{captured:g}与报表有效数量{quantity:g}不一致")
    status = f"已匹配{matched:g}/{quantity:g}件；{POLICY}"
    if missing:
        status = "未完整匹配，整行快递成本留空；" + status
        warnings.append("运费未完整匹配，快递成本及相关利润留空；"
                        f"日期={text(row.get('日期'))}，店铺={text(row.get('店铺名称'))}，"
                        f"SKU ID={text(row.get('SKU ID'))}，SKU编码={code}；" + "；".join(missing))
        remark = "运费未完整匹配，综合成本及利润留空"
    else:
        row["快递成本"] = total_cost / quantity
        remark = "快递成本按收货地区报价预估，黑吉辽圆通、其他取低价"
    row["备注"] = "；".join(filter(None, (text(row.get("备注")), remark)))
    cost_log.update({"快递成本": row["快递成本"], "运费来源": "；".join(sources), "运费匹配状态": status})


def load_shipping_table(path: Path) -> tuple[ShippingTable, list[str]]:
    """Load both region matrices, retaining invalid quotes as unknown values."""
    table = ShippingTable()
    warnings = []
    if not path.is_file():
        return table, [f"运费成本表不存在，快递成本及相关利润留空：{path}"]
    from openpyxl import load_workbook
    from openpyxl.utils import get_column_letter
    try:
        wb = load_workbook(path, read_only=True, data_only=True)
    except Exception as exc:
        raise DataQualityError(f"运费成本表读取失败：{path}；{exc}") from exc
    try:
        if "编码匹配" in wb.sheetnames:
            table.selected_region = text(wb["编码匹配"]["B4"].value)
        for carrier, name in CARRIER_SHEETS.items():
            if name not in wb.sheetnames:
                raise DataQualityError(f"运费成本表缺少工作表：{path}/{name}")
            sheet = wb[name]
            sheet.reset_dimensions()
            iterator = sheet.iter_rows(values_only=True)
            headers = None
            for row_number, row in enumerate(iterator, 1):
                if row_number > 20:
                    break
                labels = [text(v) for v in row]
                if "编码" in labels and "计费重(kg)" in labels:
                    headers = labels
                    break
            if headers is None:
                raise DataQualityError(f"运费成本表缺少编码及地区表头：{path}/{name}")
            code_col = headers.index("编码")
            weight_col = headers.index("计费重(kg)")
            regions = [(i, label) for i, label in enumerate(headers)
                       if i > weight_col and label]
            if not regions or len({v for _, v in regions}) != len(regions):
                raise DataQualityError(f"运费成本表地区表头为空或重复：{path}/{name}")
            for physical_row, row in enumerate(iterator, row_number + 1):
                raw_code = text(row[code_col]) if code_col < len(row) else ""
                code = normalize_match_code(raw_code)
                if not code:
                    continue
                for index, region in regions:
                    raw = row[index] if index < len(row) else None
                    amount = optional_number(raw)
                    valid = amount is not None and amount > 0 and not isinstance(raw, bool)
                    source = (f"运费表:{path.name}/{name}/{get_column_letter(index + 1)}"
                              f"{physical_row}/编码={raw_code}/地区={region}（预估）")
                    quote = ShippingQuote(amount if valid else None, source,
                                          "" if valid else f"费用无有效正数或公式缓存缺失：{text(raw) or '空白'}")
                    key = (code, carrier, region)
                    previous = table.prices.get(key)
                    if previous:
                        if previous.amount != quote.amount or previous.reason != quote.reason:
                            raise DataQualityError(f"运费编码价格冲突：{previous.source}；{quote.source}")
                        if index == regions[0][0]:
                            warnings.append(f"运费编码重复且价格一致，已去重：{previous.source}；{source}")
                        continue
                    table.prices[key] = quote
    finally:
        wb.close()
    codes = {key[0] for key in table.prices}
    usable = {key[0] for key, quote in table.prices.items() if quote.amount is not None}
    warnings.append(f"运费成本表按已保存的地区报价读取：编码{len(codes)}个，可用报价编码{len(usable)}个；"
                    f"其余{len(codes - usable)}个编码暂无有效运费，按实际匹配记录留空。")
    return table, warnings


def load_database_shipping(database_root: Path) -> tuple[ShippingTable, list[str]]:
    return load_shipping_table(database_root / "1.产品数据库" / "运费成本.xlsx")
