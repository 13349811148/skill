import fs from 'node:fs/promises';
import {Workbook,SpreadsheetFile,FileBlob} from './artifact_runtime.mjs';
import {col,excelDate,detailStyle,previews} from './report_layout.mjs';
const [input,output]=process.argv.slice(2),d=JSON.parse(await fs.readFile(input,'utf8'));
const wb=d.existingWorkbook?await SpreadsheetFile.importXlsx(await FileBlob.load(d.existingWorkbook)):Workbook.create();
for(const part of d.sheets||[{name:'每日销售数据',rows:d.rows}]){
 let s;try{s=wb.worksheets.getItem(part.name);}catch{}
 if(!s)s=wb.worksheets.add(part.name);
 const end=part.rows.length+2,last=col(d.columns.length);
 s.getRange(`A1:${col(Math.max(d.columns.length,part.oldColumnCount||0))}${Math.max(end,(part.oldRows||0)+2)}`).clear({applyTo:'contents'});
 if(part.rows.length)s.getRange(`A3:${last}${end}`).values=part.rows.map(r=>d.columns.map(h=>h==='日期'?excelDate(r[h]):r[h]??null));
 detailStyle(s,d.columns,part.rows,2,[[1,2,'基本信息','#304A68'],[3,6,'商品 / SKU','#304A68'],[7,8,'货品拆分','#2A7080'],[9,13,'单件成本','#9A722F'],[14,18,'销售与毛利','#3D6788'],[19,21,'推广与净利润','#3D7665'],[22,22,'说明','#66758A']]);
 wb.recalculate();
 if(part.rows.length)for(const c of ['C','D'])s.getRange(`${c}3:${c}${end}`).setNumberFormat('0');
 await previews(wb,part.name,output,[['A1:H6','identity'],['I1:R6','economics'],['S1:V6','profit']]);
 if(part.rows.length)for(const c of ['C','D'])s.getRange(`${c}3:${c}${end}`).setNumberFormat('@');
}
wb.recalculate();await(await SpreadsheetFile.exportXlsx(wb)).save(output);
console.log(`商品日报已导出:${output}; rows=${d.rows.length}`);
