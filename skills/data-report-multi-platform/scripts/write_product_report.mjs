import fs from 'node:fs/promises';
import {Workbook,SpreadsheetFile,FileBlob} from './artifact_runtime.mjs';
const [input,output]=process.argv.slice(2),d=JSON.parse(await fs.readFile(input,'utf8'));
const wb=d.existingWorkbook?await SpreadsheetFile.importXlsx(await FileBlob.load(d.existingWorkbook)):Workbook.create();
const letter=(n)=>{let value='';for(;n>0;n=Math.floor((n-1)/26))value=String.fromCharCode(65+(n-1)%26)+value;return value;};
const column=(name)=>{const index=d.columns.indexOf(name);return index<0?null:letter(index+1);};
const last=letter(d.columns.length);
const money=['综合成本','产品成本','快递成本','其他均摊成本','平台费用','到手价','实际成交金额（去退款去补单后）','毛利','推广费用','净利润'];
const parts=d.sheets||[{name:'每日销售数据',rows:d.rows,oldRows:0}];
for(const part of parts){
let s;
try{s=wb.worksheets.getItem(part.name);}catch{s=wb.worksheets.add(part.name);}
if(!s)s=wb.worksheets.add(part.name);
const n=part.rows.length;
s.getRange(`A1:${letter(Math.max(d.columns.length,part.oldColumnCount||0))}${Math.max(n,part.oldRows||0)+1}`).clear({applyTo:'contents'});
s.showGridLines=false;
s.getRange(`A1:${last}${n+1}`).values=[d.columns,...part.rows.map(r=>d.columns.map(h=>h==='日期'?new Date(r[h]+'T00:00:00Z'):r[h]===''||r[h]===undefined?null:r[h]))];
s.getRange(`A1:${last}${n+1}`).format.font={name:'Arial',size:10,color:'#202830'};
s.getRange(`A1:${last}${n+1}`).format.columnWidth=18;
s.getRange(`A1:${last}${n+1}`).format.rowHeight=24;
s.getRange('A:A').format.columnWidth=14;s.getRange('E:E').format.columnWidth=35;
for(const name of ['商品ID','SKU编码','SKU ID']){const c=column(name);if(c)s.getRange(`${c}:${c}`).format.columnWidth=23;}
const skuName=column('SKU名称');
if(skuName){s.getRange(`${skuName}:${skuName}`).format.columnWidth=42;if(n)s.getRange(`${skuName}2:${skuName}${n+1}`).format.wrapText=true;}
const remark=column('备注');
if(remark){s.getRange(`${remark}:${remark}`).format.columnWidth=75;
if(n){s.getRange(`${remark}2:${remark}${n+1}`).format.wrapText=true;s.getRange(`${remark}2:${remark}${n+1}`).format.rowHeight=48;}}
s.getRange(`A1:${last}1`).format={fill:'#25476A',font:{name:'Arial',size:10,bold:true,color:'#FFFFFF'},wrapText:true,rowHeight:64,horizontalAlignment:'center',verticalAlignment:'center'};
if(n){s.getRange(`A2:A${n+1}`).setNumberFormat('yyyy-mm-dd');
for(const name of ['商品ID','SKU编码','SKU ID']){const col=column(name);if(col)s.getRange(`${col}2:${col}${n+1}`).setNumberFormat('@');}
for(const name of money){const col=column(name);if(col)s.getRange(`${col}2:${col}${n+1}`).setNumberFormat('#,##0.00');}
const qty=column('实际成交数量（去退款去补单后）'),margin=column('毛利率');
if(qty)s.getRange(`${qty}2:${qty}${n+1}`).setNumberFormat('#,##0');if(margin)s.getRange(`${margin}2:${margin}${n+1}`).setNumberFormat('0.00%');}
// Historical schema migrations can carry the retained columns' original formats.
if(n&&part.sourceNumberFormats)s.getRange(`A2:${last}${n+1}`).format.numberFormat=part.sourceNumberFormats;
if(n)for(const name of part.sourceWrappedColumns||[]){const c=column(name);if(c)s.getRange(`${c}2:${c}${n+1}`).format.wrapText=true;}
s.freezePanes.freezeRows(1);wb.recalculate();
for(const [area,name] of process.env.DAILY_REPORT_PREVIEW==='0'?[]:[['A1:I7','identity'],['J1:R7','costs'],['S1:AC7','amounts']]){
 const img=await wb.render({sheetName:part.name,range:area,scale:1.2,format:'png'});await fs.writeFile(output+`.${d.sheets?part.name+'.':''}${name}.png`,new Uint8Array(await img.arrayBuffer()));
}
}
wb.recalculate();
await(await SpreadsheetFile.exportXlsx(wb)).save(output);
console.log(`商品日报已导出:${output}; rows=${d.rows.length}; 更新子表=${parts.map(s=>s.name).join(',')}`);
