"""Identify equal nominal sale-unit weights; never invent measured dimensions."""
from pathlib import Path
from decimal import Decimal, ROUND_HALF_UP
import re
import unicodedata

from data_quality_protection import normalize_match_code, optional_number, text

# These are the only products the user authorized for nearest-weight fallback.
NEAREST_WEIGHT_CODES = {"cb10", "lmb01", "sqm05", "wh03"}


def name_grams(name):
    value = unicodedata.normalize("NFKC", text(name)).lower()
    if any(x in value for x in ("纸箱", "空袋", "空盒", "辅材", "【托盒】", "【托盘】")):
        return None
    # A gift with no declared weight makes the total unknown, even if the main item has grams.
    if re.search(r"(?:袋|箱|盒|桶|包)[】\])）]*\s*\+", value):
        return None
    matches = list(re.finditer(r"(\d+(?:\.\d+)?)\s*(kg|千克|公斤|g|克)", value))
    if len(matches) != 1:
        return None
    m = matches[0]
    grams = float(m[1]) * (1000 if m[2] in {"kg", "千克", "公斤"} else 1)
    tail = value[m.end():]
    multiple = re.search(r"[*×]\s*(\d+(?:\.\d+)?)\s*(?:袋|箱|盒|桶|包)", tail)
    if not multiple:
        multiple = re.search(r"【\s*(\d+(?:\.\d+)?)\s*(?:袋|箱|盒|桶|包)", tail)
    if multiple:
        grams *= float(multiple[1])
    return round(grams, 6) if grams > 0 else None


def package(name):
    return next((p for p in ("箱", "盒", "桶", "袋") if p in text(name)), "")


def nominal_weights(path: Path):
    """Use explicit product names or a fully resolved BOM, independently of cost amounts."""
    if not path.is_file():
        return {}
    from openpyxl import load_workbook
    wb = load_workbook(path, read_only=True, data_only=True)
    records = {}
    try:
        if "编码成本汇总" not in wb.sheetnames:
            return {}
        headers = None
        for rn, values in enumerate(wb["编码成本汇总"].iter_rows(values_only=True), 1):
            if headers is None:
                labels = [text(v) for v in values]
                if "编码类型" in labels and "组合装编码" in labels:
                    headers = labels
                elif rn > 20:
                    break
                continue
            row = dict(zip(headers, values))
            code = normalize_match_code(row.get("组合装编码") if row.get("编码类型") == "组合装" else row.get("货品编码（组合内）"))
            if code:
                if code in records and records[code] != row:
                    # No first-match fallback for ambiguous weight identities.
                    records[code] = {}
                else:
                    records[code] = row
    finally:
        wb.close()
    cache = {}
    def resolve(code, seen=()):
        if code in cache:
            return cache[code]
        if code in seen:
            return None
        row = records.get(code, {})
        if row.get("编码类型") == "组合装":
            parts = re.split("[；;]", text(row.get("货品编码（组合内）")))
            total = 0
            for part in parts:
                m = re.fullmatch(r"\s*(.+?)\s*[×*]\s*(\d+(?:\.\d+)?)\s*", part)
                unit = resolve(normalize_match_code(m[1]), (*seen, code)) if m else None
                if unit is None:
                    total = None
                    break
                total += unit * float(m[2])
            result = round(total, 6) if total else None
        else:
            result = name_grams(row.get("产品名称") or row.get("商品名称"))
        cache[code] = result
        return result
    return {code: resolve(code) for code in records}


def reference_products(code, measurements, weights):
    """Exact nominal grams first; the four authorized codes may use nearest grams."""
    def grams(code, row):
        if "辅材" in text(row.get("分类")):
            return None
        return weights.get(code) or name_grams(row.get("货品名称"))
    row = measurements.get(code, {})
    nominal = grams(code, row)
    if nominal is None:
        return None
    candidates = [(other, data, grams(other, data)) for other, data in measurements.items()
                  if other != code and (optional_number(data.get("重量(g)")) or 0) > 0
                  and grams(other, data) is not None]
    donors = [(other, data, g) for other, data, g in candidates if g == nominal]
    if not donors and normalize_match_code(code) in NEAREST_WEIGHT_CODES and candidates:
        distance = min(abs(g - nominal) for _, _, g in candidates)
        reference_g = max(g for _, _, g in candidates if abs(g - nominal) == distance)
        donors = [(other, data, g) for other, data, g in candidates if g == reference_g]
    if not donors:
        return None
    complete = [(other, data, g) for other, data, g in donors if (optional_number(data.get("体积(cm³)")) or 0) > 0]
    if complete:
        donors = complete
    same_package = [(other, data, g) for other, data, g in donors if package(data.get("货品名称")) == package(row.get("货品名称"))]
    if same_package:
        donors = same_package
    reference_g = donors[0][2]
    return {"target_g":nominal, "reference_g":reference_g, "difference_g":abs(reference_g - nominal),
            "donors":[other for other, _, _ in donors], "kind":"同克重估算" if reference_g == nominal else "最接近克重估算"}


def fill_same_weight_quotes(table, measurements, weights):
    """Only absent weight AND volume qualify. Donors always have measured weight."""
    from shipping_costs import ShippingQuote
    from statistics import median
    original = dict(table.prices)
    for code, row in measurements.items():
        if optional_number(row.get("重量(g)")) is not None or optional_number(row.get("体积(cm³)")) is not None:
            continue
        group = reference_products(code, measurements, weights)
        if group is None:
            continue
        for key, current in original.items():
            refresh_estimate = normalize_match_code(code) in NEAREST_WEIGHT_CODES and "克重估算" in current.estimate
            if key[0] != code or (current.amount is not None and not refresh_estimate):
                continue
            refs = [(other, original.get((other, key[1], key[2]))) for other in group['donors']]
            refs = [(other, q) for other, q in refs if q is not None and q.amount is not None]
            if not refs:
                continue
            amount = float(Decimal(str(median(q.amount for _, q in refs))).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))
            detail = (f"{group['target_g']:g}g" if group['difference_g'] == 0 else
                      f"目标{group['target_g']:g}g，参考{group['reference_g']:g}g，差{group['difference_g']:g}g")
            estimate = f"{group['kind']}（{detail}，参考{','.join(other for other, _ in refs)}的地区报价中位数）"
            table.prices[key] = ShippingQuote(amount, estimate + "；" + "；".join(q.source for _, q in refs), estimate=estimate)
