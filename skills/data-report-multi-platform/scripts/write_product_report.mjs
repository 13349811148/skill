import fs from 'node:fs/promises';
import {Workbook,SpreadsheetFile,FileBlob} from './artifact_runtime.mjs';
const [input,output]=process.argv.slice(2),d=JSON.parse(await fs.readFile(input,'utf8'));
const wb=d.existingWorkbook?await SpreadsheetFile.importXlsx(await FileBlob.load(d.existingWorkbook)):Workbook.create();
const parts=d.sheets||[{name:'每日销售数据',rows:d.rows,oldRows:0}];
for(const part of parts){
let s;
try{s=wb.worksheets.getItem(part.name);}catch{s=wb.worksheets.add(part.name);}
if(!s)s=wb.worksheets.add(part.name);
const n=part.rows.length;
s.getRange(`A1:AH${Math.max(n,part.oldRows||0)+1}`).clear({applyTo:'contents'});
s.showGridLines=false;
s.getRange(`A1:AH${n+1}`).values=[d.columns,...part.rows.map(r=>d.columns.map(h=>h==='日期'?new Date(r[h]+'T00:00:00Z'):r[h]===''||r[h]===undefined?null:r[h]))];
s.getRange(`A1:AH${n+1}`).format.font={name:'Arial',size:10,color:'#202830'};
s.getRange(`A1:AH${n+1}`).format.columnWidth=18;
s.getRange(`A1:AH${n+1}`).format.rowHeight=24;
s.getRange('A:A').format.columnWidth=14;s.getRange('E:E').format.columnWidth=35;
s.getRange('G:H').format.columnWidth=23;s.getRange('AH:AH').format.columnWidth=23;
s.getRange('Z:Z').format.columnWidth=75;
if(n){s.getRange(`Z2:Z${n+1}`).format.wrapText=true;
s.getRange(`Z2:Z${n+1}`).format.rowHeight=48;}
s.getRange('A1:AH1').format={fill:'#25476A',font:{name:'Arial',size:10,bold:true,color:'#FFFFFF'},wrapText:true,rowHeight:64,horizontalAlignment:'center',verticalAlignment:'center'};
if(n){s.getRange(`A2:A${n+1}`).setNumberFormat('yyyy-mm-dd');
for(const col of ['G','H','AH'])s.getRange(`${col}2:${col}${n+1}`).setNumberFormat('@');
s.getRange(`I2:Y${n+1}`).setNumberFormat('#,##0.00');s.getRange(`N2:N${n+1}`).setNumberFormat('#,##0');s.getRange(`T2:T${n+1}`).setNumberFormat('0.00%');}
s.freezePanes.freezeRows(1);wb.recalculate();
for(const [area,name] of process.env.DAILY_REPORT_PREVIEW==='0'?[]:[['A1:H7','identity'],['M1:Z7','amounts']]){
 const img=await wb.render({sheetName:part.name,range:area,scale:1.2,format:'png'});await fs.writeFile(output+`.${d.sheets?part.name+'.':''}${name}.png`,new Uint8Array(await img.arrayBuffer()));
}
}
wb.recalculate();
await(await SpreadsheetFile.exportXlsx(wb)).save(output);
console.log(`商品日报已导出:${output}; rows=${d.rows.length}; 更新子表=${parts.map(s=>s.name).join(',')}`);
