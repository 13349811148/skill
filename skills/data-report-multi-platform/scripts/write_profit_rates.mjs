import fs from 'node:fs/promises';
import {Workbook,SpreadsheetFile} from './artifact_runtime.mjs';
const [input,output]=process.argv.slice(2),payload=JSON.parse(await fs.readFile(input,'utf8')),wb=Workbook.create();
// Author only the new cells in a small dependency model. Native packaging keeps
// the user's pivot tables, caches and every existing cell in the final workbook.
for(const part of payload.rateEdits){
 const sheet=wb.worksheets.add(part.name);
 for(const row of part.dependencies){
  sheet.getRange(`${part.salesColumn}${row.row}`).values=[[row.sales??null]];
  sheet.getRange(`${part.profitColumn}${row.row}`).values=[[row.profit??null]];
  sheet.getRange(`${part.column}${row.row}`).formulas=[[`=IF(COUNT(${part.profitColumn}${row.row},${part.salesColumn}${row.row})=2,IF(${part.salesColumn}${row.row}>0,${part.profitColumn}${row.row}/${part.salesColumn}${row.row},""),"")`]];
  sheet.getRange(`${part.column}${row.row}`).format={font:{name:'Microsoft YaHei',size:10,color:'#253449'},verticalAlignment:'center'};
  sheet.getRange(`${part.column}${row.row}`).setNumberFormat('0.00%;[Red](0.00%);0.00%');
 }
 sheet.getRange(`${part.column}:${part.column}`).format.columnWidth=18;
 sheet.getRange(`${part.column}${part.headerRow}`).values=[['净利率']];
 sheet.getRange(`${part.column}${part.headerRow}`).format={fill:'#E9EEF5',font:{name:'Microsoft YaHei',size:10,bold:true,color:'#263A52'},wrapText:true,horizontalAlignment:'center',verticalAlignment:'center'};
 if(part.monthly){sheet.getRange(`${part.column}1`).values=[['经营净利率']];sheet.getRange(`${part.column}1`).format={fill:'#3D7665',font:{name:'Microsoft YaHei',size:11,bold:true,color:'#FFFFFF'},horizontalAlignment:'center',verticalAlignment:'center'};}
}
wb.recalculate();await(await SpreadsheetFile.exportXlsx(wb)).save(output);
