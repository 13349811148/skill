import fs from 'node:fs/promises';
import {Workbook,SpreadsheetFile} from './artifact_runtime.mjs';
import {col,excelDate,detailStyle,previews} from './report_layout.mjs';
const [input,output]=process.argv.slice(2),d=JSON.parse(await fs.readFile(input,'utf8')),wb=Workbook.create();
for(const part of d.sheets){
 const s=wb.worksheets.add(part.name),n=part.rows.length;
 if(n)s.getRange(`A3:${col(d.columns.length)}${n+2}`).values=part.rows.map(r=>d.columns.map(h=>h==='日期'?excelDate(r[h]):r[h]??null));
 detailStyle(s,d.columns,part.rows,2,[[1,2,'基本信息','#304A68'],[3,5,'单个货品','#2A7080'],[6,8,'销量与销售额','#3D6788'],[9,11,'毛利、推广与净利润','#3D7665'],[12,13,'来源与口径','#66758A']]);
 if(n)for(let row=3;row<=n+2;row++)s.getRange(`H${row}`).formulas=[[`=IF(COUNT(G${row},F${row})=2,IF(F${row}>0,G${row}/F${row},""),"")`]];
 wb.recalculate();await previews(wb,part.name,output,[['A1:H7','goods'],['I1:M7','profit']]);
}
await(await SpreadsheetFile.exportXlsx(wb)).save(output);
console.log(`货品日报已导出:${output}; rows=${d.rows.length}`);
