import fs from 'node:fs/promises';
export const col=n=>{let v='';for(;n>0;n=Math.floor((n-1)/26))v=String.fromCharCode(65+(n-1)%26)+v;return v;};
export function excelDate(value){
 const raw=String(value??'').slice(0,10),m=/^(\d{4})-(\d{2})-(\d{2})$/.exec(raw);
 if(!m)throw new Error(`无效日报日期: ${value}`);
 const t=Date.UTC(...[+m[1],+m[2]-1,+m[3]]);
 if(new Date(t).toISOString().slice(0,10)!==raw)throw new Error(`无效日报日期: ${value}`);
 return (t-Date.UTC(1899,11,30))/86400000;
}
export function band(sheet,address,label,color){
 const range=sheet.getRange(address);range.merge();range.values=[[label]];
 range.format={fill:color,font:{name:'Microsoft YaHei',size:11,bold:true,color:'#FFFFFF'},horizontalAlignment:'center',verticalAlignment:'center',rowHeight:30};
}
export function detailStyle(sheet,headers,rows,headerRow,groups){
 const last=col(headers.length),first=headerRow+1,end=headerRow+rows.length;
 sheet.showGridLines=false;
 sheet.getRange(`A${headerRow}:${last}${Math.max(headerRow,end)}`).format={font:{name:'Microsoft YaHei',size:10,color:'#253449'},columnWidth:16,rowHeight:38,verticalAlignment:'center'};
 sheet.getRange(`A${headerRow}:${last}${headerRow}`).values=[headers];
 sheet.getRange(`A${headerRow}:${last}${headerRow}`).format={fill:'#E9EEF5',font:{name:'Microsoft YaHei',size:10,bold:true,color:'#263A52'},wrapText:true,rowHeight:58,horizontalAlignment:'center',verticalAlignment:'center'};
 let previous='',stripe=-1;
 rows.forEach((row,i)=>{
  const r=first+i;if(row['日期']!==previous){stripe++;previous=row['日期'];}
  sheet.getRange(`A${r}:${last}${r}`).format.fill=stripe%2?'#F2F6FA':'#FFFFFF';
  const lines=Math.max(1,...['货品名称','货品数量','来源SKU及货品数量'].map(h=>String(row[h]??'').split('\n').length));
  sheet.getRange(`A${r}:${last}${r}`).format.rowHeight=Math.max(52,lines*19+8);
  if(i===0||rows[i-1]['日期']!==row['日期'])sheet.getRange(`A${r}:${last}${r}`).format.borders={top:{style:'thin',color:'#A8B9CA'}};
 });
 for(const [start,finish,label,color] of groups){
  band(sheet,`${col(start)}${headerRow-1}:${col(finish)}${headerRow-1}`,label,color);
  sheet.getRange(`${col(start)}${headerRow}:${col(start)}${Math.max(headerRow,end)}`).format.borders={left:{style:'medium',color:'#CFD9E5'}};
 }
 const widths={'日期':14,'店铺名称':34,'商品ID':23,'SKU ID':23,'SKU编码':15,'SKU名称':42,'货品编码':17,'货品名称':48,'货品数量':14,'来源SKU及货品数量':28,'基础单位':12,'销售额':18,'客单价（货品均价）':19,'毛利':18,'推广费用':18,'净利润':18,'备注':70,'实际成交数量（去退款去补单后）':18,'实际成交金额（去退款去补单后）':21};
 headers.forEach((h,i)=>{
  const c=col(i+1);sheet.getRange(`${c}:${c}`).format.columnWidth=widths[h]||16;
  if(rows.length){
   const range=sheet.getRange(`${c}${first}:${c}${end}`);
   if(h==='日期'){range.setNumberFormat('yyyy-mm-dd');range.format.horizontalAlignment='center';}
   else if(['商品ID','SKU ID','SKU编码','货品编码'].includes(h))range.setNumberFormat('@');
   else if(h==='毛利率')range.setNumberFormat('0.00%');
   else if(['货品数量','实际成交数量（去退款去补单后）'].includes(h))range.setNumberFormat(rows.some(r=>typeof r[h]==='number'&&!Number.isInteger(r[h]))?'#,##0.00':'#,##0');
   else if(['综合成本','产品成本','快递成本','其他均摊成本','平台费用','到手价','销售额','客单价（货品均价）','实际成交金额（去退款去补单后）','毛利','推广费用','净利润'].includes(h))range.setNumberFormat('#,##0.00;[Red](#,##0.00);0.00');
   if(['SKU名称','货品名称','货品数量','来源SKU及货品数量','备注'].includes(h))range.format.wrapText=true;
  }
 });
 const net=headers.indexOf('净利润');
 if(net>=0&&rows.length){const c=col(net+1);sheet.getRange(`${c}${first}:${c}${end}`).format.font.bold=true;}
 sheet.freezePanes.freezeRows(headerRow);
}
export async function previews(wb,sheet,output,parts){
 if(process.env.DAILY_REPORT_PREVIEW==='0')return;
 for(const [range,name] of parts){const img=await wb.render({sheetName:sheet,range,scale:1.15,format:'png'});await fs.writeFile(output+`.${sheet}.${name}.png`,new Uint8Array(await img.arrayBuffer()));}
}
