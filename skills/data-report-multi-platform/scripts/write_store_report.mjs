import fs from 'node:fs/promises';
import path from 'node:path';
import {FileBlob,SpreadsheetFile} from './artifact_runtime.mjs';
import {band,excelDate} from './report_layout.mjs';
const [input,output]=process.argv.slice(2);
const payload=JSON.parse(await fs.readFile(input,'utf8'));
const wb=await SpreadsheetFile.importXlsx(await FileBlob.load(payload.existingWorkbook||path.join(import.meta.dirname,'../assets/store_daily_template.xlsx')));
const sheet=wb.worksheets.getItem('①平台日报');
// The bundled asset contains only the requested sheet. Never export sample rows.
const n=payload.rows.length,end=Math.max(33,n+3,(payload.oldRows||0)+3);
sheet.getRange(`A4:AA${end}`).clear({applyTo:'contents'});
if(n){
 for(let i=5;i<=n+3;i++)sheet.getRange(`A${i}:AA${i}`).copyFrom(sheet.getRange('A4:AA4'),'all');
 sheet.getRange(`A4:AA${n+3}`).values=payload.rows.map(r=>payload.headers.map(h=>h==='日期'?excelDate(r[h]):r[h]??null));
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
  const retained=payload.rows[i-4]._原值保留||{};
  const labels={M:'新客数',N:'老客数',P:'客单价',Q:'支付转化率',R:'净成交额',U:'推广ROI',V:'推广费比'};
  for(const [c,f] of Object.entries(form)){
   if(Object.hasOwn(retained,labels[c]))sheet.getRange(`${c}${i}`).values=[[retained[labels[c]]]];
   else sheet.getRange(`${c}${i}`).formulas=[[f]];
  }
 }
 sheet.getRange(`A4:A${n+3}`).setNumberFormat('yyyy-mm-dd');
 for(const c of ['O','Q','V'])sheet.getRange(`${c}4:${c}${n+3}`).setNumberFormat('0.00%');
 sheet.getRange(`M4:N${n+3}`).setNumberFormat('#,##0');
 for(const c of ['H','K','P','R','S','T','U','AA'])sheet.getRange(`${c}4:${c}${n+3}`).setNumberFormat('#,##0.00');
}
sheet.getRange('A1:AA1').unmerge();sheet.getRange('A1:AA1').merge();
sheet.getRange('A1').values=[['店铺每日经营日报 · 净利润按商品日报汇总（含空烧推广；暂计0见日志）']];
sheet.getRange('A1:AA1').format={fill:'#1E334B',font:{name:'Microsoft YaHei',size:16,bold:true,color:'#FFFFFF'},rowHeight:40,verticalAlignment:'center'};
sheet.getRange('A2:AA2').unmerge();sheet.getRange('A2:AA2').clear({applyTo:'contents'});
for(const [range,label,color] of [['A2:D2','基本信息','#304A68'],['E2:G2','流量（曝光 / 点击为推广口径）','#2A7080'],['H2:L2','支付与退款','#3D6788'],['M2:Q2','客户与转化','#3D6788'],['R2:V2','成交与推广','#9A722F'],['W2:Z2','店铺记录','#66758A'],['AA2:AA2','经营利润','#3D7665']])band(sheet,range,label,color);
sheet.getRange('A3:AA3').values=[payload.headers];
sheet.getRange('A3:AA3').format={fill:'#E9EEF5',font:{name:'Microsoft YaHei',size:10,bold:true,color:'#263A52'},rowHeight:54,wrapText:true,horizontalAlignment:'center',verticalAlignment:'center'};
if(n){
 sheet.getRange(`A4:AA${n+3}`).format.font={name:'Microsoft YaHei',size:10,color:'#253449'};
 sheet.getRange(`A4:AA${n+3}`).format.rowHeight=40;
 sheet.getRange(`A4:AA${n+3}`).format.verticalAlignment='center';
 sheet.getRange(`AA4:AA${n+3}`).setNumberFormat('#,##0.00;[Red](#,##0.00);0.00');
 for(let r=4;r<=n+3;r++)sheet.getRange(`A${r}:AA${r}`).format.fill=r%2?'#F2F6FA':'#FFFFFF';
 sheet.getRange(`AA4:AA${n+3}`).format.font.bold=true;
}
sheet.getRange('AA:AA').format.columnWidth=20;
sheet.showGridLines=false;
// The configured shop's full name must remain readable in the template column.
sheet.getRange('C:C').format.columnWidth=36;
sheet.freezePanes.freezeRows(3);
wb.recalculate();
const check=await wb.inspect({kind:'match',searchTerm:'#REF!|#DIV/0!|#VALUE!|#NAME\\?|#NUM!',options:{useRegex:true,maxResults:20},maxChars:2000});
await fs.writeFile(output+'.inspect.ndjson',check.ndjson);
if(process.env.STORE_REPORT_PREVIEW){
 for(const [area,name] of [['A1:J8','identity'],['R1:AA8','profit']]){
  const image=await wb.render({sheetName:'①平台日报',range:area,scale:1.2,format:'png'});
  await fs.writeFile(output+`.${name}.png`,new Uint8Array(await image.arrayBuffer()));
 }
}
await(await SpreadsheetFile.exportXlsx(wb)).save(output);
console.log(`Saved ${output}`);
