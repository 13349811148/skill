from collections import defaultdict
from decimal import Decimal,ROUND_HALF_UP
from pathlib import Path
import json
from portable_reports import read_workbook_sheets

def value(v,required=False):
    if v is None or str(v).strip() in ('','-','n.a.'):
        if required:raise ValueError('缺少推广数字')
        return Decimal(0)
    d=Decimal(str(v).replace(',',''))
    if not d.is_finite() or d<0:raise ValueError(f'无效推广数字 {v}')
    return d

def load(root,target_dates):
    root=Path(root);folder=root/'4.推广数据库'
    if not folder.exists():folder=root/'推广数据'
    # Latest complete dated product snapshot, never blend an old allocated snapshot into a new export.
    products={};accounts={};logs=[]
    for path in sorted(folder.rglob('*.xlsx')):
        if path.name.startswith('~$'):continue
        if any(t in str(path).lower() for t in ['备份','副本','backup']):raise ValueError(f'推广目录含备份 {path}')
        meta=path.with_suffix(path.suffix+'.archive.json')
        md=json.loads(meta.read_text(encoding='utf-8-sig')) if meta.exists() else {}
        relative=path.relative_to(folder).parts
        shop=md.get('shop') or relative[0] if len(relative)>1 else md.get('shop')
        if not shop:raise ValueError(f'推广来源店铺未知 {path}')
        rank=(path.stat().st_mtime_ns,str(path));batches=defaultdict(list)
        for sheet in read_workbook_sheets(path):
            if not sheet.rows:continue
            h=[str(v).strip() if v is not None else '' for v in sheet.rows[0]]
            for raw in sheet.rows[1:]:
                r=dict(zip(h,raw));dt=str(r.get('归档日期') or r.get('日期') or '')[:10]
                if dt not in target_dates:continue
                pid=r.get('归档商品ID') or r.get('商品ID')
                if pid is None and ('账户每日汇总' in str(path)):
                    key=(dt,shop);fee=value(r.get('总花费(元)') if str(r.get('总花费(元)') or '').strip() not in ('','-') else r.get('成交花费(元)'),True)
                    if key not in accounts or rank>accounts[key][0]:accounts[key]=(rank,fee,path)
                elif pid is not None and str(pid).isdigit():
                    batches[(dt,shop)].append(r)
                    if '当日成交花费(元)' in r:
                        key=(dt,shop);hint_rank=(-1,str(path));fee=value(r['当日成交花费(元)'],True)
                        if key not in accounts or hint_rank>accounts[key][0]:accounts[key]=(hint_rank,fee,path)
        for key,rows in batches.items():
            if key not in products or rank>products[key][0]:
                if key in products:logs.append(f'推广日快照更新:{key},保留={path},忽略={products[key][2]}')
                products[key]=(rank,rows,path)
    costs=defaultdict(Decimal)
    for key,(_,rows,path) in sorted(products.items()):
        if key not in accounts:raise ValueError(f'缺少店铺每日推广总额 {key}')
        total=accounts[key][1];managed=[];regular=defaultdict(Decimal)
        for r in rows:
            pid=str(r.get('归档商品ID') or r['商品ID']);scene=r.get('推广场景')
            if scene=='全店托管':managed.append((pid,value(r.get('交易额(元)'),True)))
            elif scene in ('稳定成本推广','商品推广','标准推广','全站推广'):
                fee=r.get('总花费(元)');fee=r.get('成交花费(元)') if fee is None or str(fee).strip() in ('','-') else fee
                regular[pid]+=value(fee,True)
            else:raise ValueError(f'未知推广类型 {key}/{pid}/{scene}')
        actual=sum(regular.values(),Decimal(0));pool=total-actual;denominator=sum((a for _,a in managed),Decimal(0))
        if pool<0:raise ValueError(f'托管花费负差 {key}:总额{total},商品推广{actual}')
        if pool>0 and denominator<=0:raise ValueError(f'托管有花费但成交额为0，无法分摊 {key}')
        amounts=defaultdict(Decimal)
        for pid,amount in managed:amounts[pid]+=amount
        allocated={pid:(pool*amount/denominator).quantize(Decimal('.01'),rounding=ROUND_HALF_UP) if denominator else Decimal(0) for pid,amount in amounts.items()}
        tail=pool-sum(allocated.values(),Decimal(0))
        if tail:
            pid=max(amounts,key=amounts.get);allocated[pid]+=tail
        for pid,fee in allocated.items():costs[(*key,pid)]+=fee
        for pid,fee in regular.items():costs[(*key,pid)]+=fee
        assert sum(allocated.values(),Decimal(0))+actual==total
        logs.append(f'推广核对:{key},账户总额={total},商品推广实际={actual},托管分摊={pool},尾差={tail},逐日核对=0')
    for key in accounts:
        if key not in products and accounts[key][1]>0:raise ValueError(f'账户有推广花费但无商品明细 {key}')
    return {k:float(v) for k,v in costs.items()},logs
