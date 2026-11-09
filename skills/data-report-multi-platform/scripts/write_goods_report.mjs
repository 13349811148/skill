import fs from 'node:fs/promises';
import {Workbook,SpreadsheetFile} from './artifact_runtime.mjs';
import {col,excelDate,detailStyle,previews} from './report_layout.mjs';
const [input,output]=process.argv.slice(2),d=JSON.parse(await fs.readFile(input,'utf8')),wb=Workbook.create();
for(const part of d.sheets){
 const s=wb.worksheets.add(part.name),n=part.rows.length;
 if(n)s.getRange(`A3:${col(d.columns.length)}${n+2}`).values=part.rows.map(r=>d.columns.map(h=>h==='日期'?excelDate(r[h]):r[h]??null));
 detailStyle(s,d.columns,part.rows,2,[[1,2,'基本信息','#304A68'],[3,5,'单个货品','#2A7080'],[6,6,'有效货品数量','#3D7665'],[7,8,'来源与口径','#66758A']]);
 wb.recalculate();await previews(wb,part.name,output,[['A1:H7','goods']]);
}
await(await SpreadsheetFile.exportXlsx(wb)).save(output);
console.log(`货品日报已导出:${output}; rows=${d.rows.length}`);
