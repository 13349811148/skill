import fs from 'node:fs/promises';
import {Workbook,SpreadsheetFile} from './artifact_runtime.mjs';
const [input,output]=process.argv.slice(2),d=JSON.parse(await fs.readFile(input,'utf8'));
const wb=Workbook.create(),s=wb.worksheets.add('每日销售数据'),n=d.rows.length;
s.showGridLines=false;
s.getRange(`A1:AH${n+1}`).values=[d.columns,...d.rows.map(r=>d.columns.map(h=>h==='日期'?new Date(r[h]+'T00:00:00Z'):r[h]===''||r[h]===undefined?null:r[h]))];
s.getRange(`A1:AH${n+1}`).format.font={name:'Arial',size:10,color:'#202830'};
s.getRange(`A1:AH${n+1}`).format.columnWidth=18;
s.getRange(`A1:AH${n+1}`).format.rowHeight=24;
s.getRange('A:A').format.columnWidth=14;s.getRange('E:E').format.columnWidth=35;
s.getRange('G:H').format.columnWidth=23;s.getRange('AH:AH').format.columnWidth=23;
s.getRange('Z:Z').format.columnWidth=75;
s.getRange(`Z2:Z${n+1}`).format.wrapText=true;
s.getRange(`Z2:Z${n+1}`).format.rowHeight=48;
s.getRange('A1:AH1').format={fill:'#25476A',font:{name:'Arial',size:10,bold:true,color:'#FFFFFF'},wrapText:true,rowHeight:64,horizontalAlignment:'center',verticalAlignment:'center'};
s.getRange(`A2:A${n+1}`).setNumberFormat('yyyy-mm-dd');
for(const col of ['G','H','AH'])s.getRange(`${col}2:${col}${n+1}`).setNumberFormat('@');
s.getRange(`I2:Y${n+1}`).setNumberFormat('#,##0.00');s.getRange(`N2:N${n+1}`).setNumberFormat('#,##0');s.getRange(`T2:T${n+1}`).setNumberFormat('0.00%');
s.freezePanes.freezeRows(1);wb.recalculate();
for(const [area,name] of [['A1:H7','identity'],['M1:Z7','amounts']]){
 const img=await wb.render({sheetName:'每日销售数据',range:area,scale:1.2,format:'png'});await fs.writeFile(output+`.${name}.png`,new Uint8Array(await img.arrayBuffer()));
}
await(await SpreadsheetFile.exportXlsx(wb)).save(output);
console.log(`商品日报已导出:${output}; rows=${n}`);
