"""Explicit partial-reference mode; never suppress order confirmation or bad data."""
import sys
from pathlib import Path
from portable_reports import install as portable_install
from pdd_promotion_daily import load as load_promotion

def install(module):
    portable_install(module)
    module.load_promo_costs=load_promotion
    original_marketing=module.load_marketing_rows;original_costs=module.load_costs_by_merchant_code;original_build=module.build_report
    mode='--allow-missing-reference' in sys.argv
    missing=[]
    original_resolve_marketing=module.resolve_marketing_path;original_resolve_cost=module.resolve_cost_table_path
    def resolve_marketing(argument,root):
        if mode and not argument and not (root/'营销活动监控').exists():return root/'营销活动监控/营销活动.xlsx'
        return original_resolve_marketing(argument,root)
    def resolve_cost(argument,root,marketing_path):
        if mode and not argument and not (root/'营销活动监控').exists():return root/'营销活动监控/成本表.xlsx'
        return original_resolve_cost(argument,root,marketing_path)
    module.resolve_marketing_path=resolve_marketing;module.resolve_cost_table_path=resolve_cost
    def marketing(path):
        if not Path(path).exists() and mode:
            missing.append('营销活动表');return []
        return original_marketing(path)
    def costs(path):
        if not Path(path).exists() and mode:
            missing.append('成本表');return {},['用户授权先填写可用字段：缺少成本表，成本与利润未计算']
        return original_costs(path)
    def build(*args,**kwargs):
        columns,rows,warnings,logs=original_build(*args,**kwargs)
        if missing:
            fields=['产品成本','活动价','报名价',module.PROMOTION_MECHANISM_COLUMN,'每单补贴金额','总补贴金额','净销售额（实际成交+总补贴金额）','毛利','毛利率','店铺费用','平摊管理费用','净利润','定价是否合理']
            for row in rows:
                for field in fields:row[field]=''
                for field in ['品牌','所属运营','大类']:
                    if row.get(field)=='营销活动表未匹配/未维护':row[field]=''
                row['备注']='；'.join(filter(None,[row.get('备注',''),'缺营销/成本，补贴和利润待补；商品推广按实际，托管按成交额分摊']))
            warnings.insert(0,'缺少'+','.join(sorted(set(missing)))+'；按用户授权输出已知商品字段，成本、活动、补贴与利润相关字段留空。')
        return columns,rows,warnings,logs
    module.load_marketing_rows=marketing;module.load_costs_by_merchant_code=costs;module.build_report=build
