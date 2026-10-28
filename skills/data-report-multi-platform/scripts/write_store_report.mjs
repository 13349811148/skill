import fs from 'node:fs/promises';
import path from 'node:path';
import {FileBlob,SpreadsheetFile} from './artifact_runtime.mjs';
const [input,output]=process.argv.slice(2);
const payload=JSON.parse(await fs.readFile(input,'utf8'));
const wb=await SpreadsheetFile.importXlsx(await FileBlob.load(payload.existingWorkbook||path.join(import.meta.dirname,'../assets/store_daily_template.xlsx')));
const sheet=wb.worksheets.getItem('①平台日报');
// The bundled asset contains only the requested sheet. Never export sample rows.
const n=payload.rows.length,end=Math.max(33,n+3,(payload.oldRows||0)+3);
sheet.getRange(`A4:Z${end}`).clear({applyTo:'contents'});
if(n){
 for(let i=5;i<=n+3;i++)sheet.getRange(`A${i}:Z${i}`).copyFrom(sheet.getRange('A4:Z4'),'all');
 sheet.getRange(`A4:Z${n+3}`).values=payload.rows.map(r=>payload.headers.map(h=>h==='日期'?new Date(r[h]+'T00:00:00Z'):r[h]??null));
 for(let i=4;i<=n+3;i++){
  const rawRate=payload.rows[i-4]._后台支付转化率;
  const form={
   M:`=IF(COUNT(J${i},O${i})=2,IF(AND(J${i}>=0,O${i}>=0,O${i}<=1),J${i}-ROUND(J${i}*O${i},0),""),"")`,
   N:`=IF(COUNT(J${i},O${i})=2,IF(AND(J${i}>=0,O${i}>=0,O${i}<=1),ROUND(J${i}*O${i},0),""),"")`,
   P:`=IF(COUNT(H${i},I${i})=2,IF(I${i}>0,H${i}/I${i},""),"")`,
   Q:`=IF(COUNT(J${i},E${i})=2,IF(E${i}>0,J${i}/E${i},""),${typeof rawRate==='number'?rawRate:'""'})`,
   R:`=IF(COUNT(H${i},K${i})=2,H${i}-K${i},"")`,
   U:`=IF(COUNT(R${i},S${i})=2,IF(AND(R${i}>=0,S${i}>0),R${i}/S${i},""),"")`,
   V:`=IF(COUNT(R${i},S${i})=2,IF(AND(R${i}>0,S${i}>=0),S${i}/R${i},""),"")`
  };
  for(const [c,f] of Object.entries(form))sheet.getRange(`${c}${i}`).formulas=[[f]];
 }
 sheet.getRange(`A4:A${n+3}`).setNumberFormat('yyyy-mm-dd');
 for(const c of ['O','Q','V'])sheet.getRange(`${c}4:${c}${n+3}`).setNumberFormat('0.00%');
 sheet.getRange(`M4:N${n+3}`).setNumberFormat('#,##0');
 for(const c of ['H','K','P','R','S','T','U'])sheet.getRange(`${c}4:${c}${n+3}`).setNumberFormat('#,##0.00');
}
sheet.getRange('A2').values=[[payload.report_note||'口径：支付日原始支付金额；退款按实际退款日；曝光/点击仅推广；ROI=全店净成交额÷推广花费；费比=推广花费÷全店净成交额。缺失及负净成交比率留空，详见独立日志。']];
sheet.getRange('A2:Z2').format.wrapText=true;
sheet.getRange('A2:Z2').format.rowHeight=42;
// The configured shop's full name must remain readable in the template column.
sheet.getRange('C:C').format.columnWidth=36;
sheet.freezePanes.freezeRows(3);
wb.recalculate();
const check=await wb.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#NUM!',options:{useRegex:true,maxResults:20},maxChars:2000});
await fs.writeFile(output+'.inspect.ndjson',check.ndjson);
if(process.env.STORE_REPORT_PREVIEW){
 for(const [area,name] of [['A1:J8','identity'],['K1:Z8','metrics']]){
  const image=await wb.render({sheetName:'①平台日报',range:area,scale:1.2,format:'png'});
  await fs.writeFile(output+`.${name}.png`,new Uint8Array(await image.arrayBuffer()));
 }
}
await(await SpreadsheetFile.exportXlsx(wb)).save(output);
console.log(`Saved ${output}`);
