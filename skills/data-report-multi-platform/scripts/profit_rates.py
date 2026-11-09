"""Net profit rate uses net sales, never receipts before refunds or a cost base."""
import math


def net_margin(profit, net_sales):
    if profit is None or profit == '' or net_sales is None or net_sales == '':
        return None
    profit, net_sales = float(profit), float(net_sales)
    if not math.isfinite(profit) or not math.isfinite(net_sales):
        raise ValueError('净利率输入必须为有限数值')
    return profit/net_sales if net_sales > 0 else None


def refresh_net_margins(directory):
    """Add rates without moving source columns or rebuilding native user pivots."""
    from pathlib import Path
    import json,re
    from openpyxl import load_workbook
    from report_archive import commit,locked,store,day
    from report_extensions import store_payload
    from goods_catalog import GOODS_COLUMNS
    def letter(number):
        value=''
        while number:
            number,remainder=divmod(number-1,26);value=chr(65+remainder)+value
        return value
    root=Path(directory);outputs=[]
    for path in sorted(root.glob('货品日报_*.xlsx')):
        state_path=path.with_suffix('.goods.json')
        state=json.loads(state_path.read_text(encoding='utf-8')) if state_path.exists() else {}
        workbook=load_workbook(path,read_only=True,data_only=True)
        sheets=[];all_rows=[];seen=set();edits=[];summaries=[]
        try:
            for sheet in workbook:
                sheet.reset_dimensions();raw_rows=list(sheet.iter_rows(values_only=True))
                monthly=bool(re.fullmatch(r'\d{4}-\d{2}',sheet.title))
                if monthly:
                    headers=list(raw_rows[1]);legacy=[c for c in GOODS_COLUMNS if c!='净利率']
                    if headers not in [GOODS_COLUMNS,legacy]:raise ValueError(f'货品表头未识别: {sheet.title}')
                    header_row=2;column='N';profit_column='K';sales_column='G';rows=[]
                    dependencies=[]
                    for physical,raw in enumerate(raw_rows[2:],3):
                        if not raw[0]:continue
                        row=dict(zip(headers,raw));row['日期']=day(row['日期'])
                        key=row['日期'],row['店铺名称'],row['货品编码']
                        if key in seen:raise ValueError(f'货品业务键重复: {key}')
                        seen.add(key)
                        if row['日期'][:7]!=sheet.title:raise ValueError('货品日期与月份子表不一致')
                        row['净利率']=net_margin(row.get('净利润'),row.get('销售额'))
                        dependencies.append({'row':physical,'profit':row.get('净利润'),'sales':row.get('销售额')})
                        rows.append(row)
                    sheets.append({'name':sheet.title,'rows':rows});all_rows.extend(rows)
                else:
                    headers=None
                    for i,raw in enumerate(raw_rows[:20],1):
                        names=[str(v or '').replace('求和项:','').replace('求和项：','').strip() for v in raw]
                        if '净利润' in names and ('净销售额' in names or '销售额' in names):
                            headers=names;header_row=i;break
                    if headers is None:continue
                    profit_column=letter(headers.index('净利润')+1)
                    sales_column=letter(headers.index('净销售额' if '净销售额' in headers else '销售额')+1)
                    column=letter(headers.index('净利率')+1 if '净利率' in headers else len(headers)+1)
                    dependencies=[]
                    for physical,raw in enumerate(raw_rows[header_row:],header_row+1):
                        if not any(v is not None for v in raw):continue
                        pi=headers.index('净利润');profit=raw[pi] if pi<len(raw) else None
                        si=headers.index('净销售额' if '净销售额' in headers else '销售额');sales=raw[si] if si<len(raw) else None
                        net_margin(profit,sales)
                        dependencies.append({'row':physical,'profit':profit,'sales':sales})
                    summaries.append(sheet.title)
                targets=[f'{column}{header_row}',*[f'{column}{d["row"]}' for d in dependencies]]
                if monthly:targets.insert(0,f'{column}1')
                edits.append({'name':sheet.title,'monthly':monthly,'headerRow':header_row,'column':column,
                              'profitColumn':profit_column,'salesColumn':sales_column,'dependencies':dependencies,'targets':targets})
        finally:workbook.close()
        payload=dict(state,columns=GOODS_COLUMNS,rows=all_rows,sheets=sheets,kind='goods',marginOnly=True,
                     nativeGoodsSource=str(path),rateEdits=edits,summary_rate_sheets=summaries,
                     net_margin_rule='净利润÷净销售额；现有销售额为去退款去补单的实际成交金额；分母不大于0或缺失则留空')
        with locked(path):commit(path,payload,'write_profit_rates.mjs')
        outputs.append(path)
    for path in sorted(root.glob('店铺日报_????-??.xlsx')):
        payload=store_payload(path)
        for row in payload['rows']:row['净利率']=net_margin(row.get('店铺净利润'),row.get('净成交额'))
        payload['net_margin_rule']='店铺净利润÷净成交额；分母不大于0或缺失则留空'
        store(payload,root,path);outputs.append(path)
    if not outputs:raise ValueError('没有可更新的货品或店铺日报')
    print('已增加净利率: '+', '.join(str(p) for p in outputs))
    return outputs
