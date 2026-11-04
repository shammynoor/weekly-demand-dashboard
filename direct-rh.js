/* Direct RH imports only update the existing OneDrive source CSV. */
(function(root){
 'use strict';
 const required=['tenant_name','request_date','request_time','request_type','admin_type','age_between_created_and_processed'];
 function columns(text,papa){
  return (papa.parse(text,{preview:1,skipEmptyLines:'greedy'}).data[0]||[]).map(String);
 }
 function recognises(text,papa){const fields=columns(text,papa);return required.every(field=>fields.includes(field))}
 function parse(text,papa){
  const result=papa.parse(text,{skipEmptyLines:'greedy'});
  if(result.errors?.length)throw Error('Direct RH CSV could not be read: '+result.errors[0].message);
  const [fields,...rows]=result.data;
  if(!fields||!required.every(field=>fields.includes(field))||new Set(fields).size!==fields.length)
   throw Error('Direct RH CSV is missing required columns or has duplicate headings.');
  if(!rows.length)throw Error('Direct RH CSV has no records.');
  if(rows.some(row=>row.length!==fields.length))throw Error('Direct RH CSV has rows with the wrong number of columns.');
  return {fields,rows};
 }
 function timestamp(date,time){
  const m=/^(\d{1,2})-([a-z]{3})-(\d{2}|\d{4})$/i.exec(String(date).trim());
  const t=/^(\d{2}):(\d{2})$/.exec(String(time).trim());
  if(!m||!t)throw Error('Direct RH CSV has an invalid request date or time. Expected DD-MMM-YY and HH:mm.');
  const year=m[3].length===2?2000+Number(m[3]):Number(m[3]);
  const month=['jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec'].indexOf(m[2].toLowerCase());
  const day=Number(m[1]),hour=Number(t[1]),minute=Number(t[2]);
  const value=Date.UTC(year,month,day,hour,minute),check=new Date(value);
  if(month<0||hour>23||minute>59||check.getUTCFullYear()!==year||check.getUTCMonth()!==month||check.getUTCDate()!==day)
   throw Error('Direct RH CSV has an invalid request date or time.');
  // Numeric wall-clock ordering: do not apply browser timezone or daylight-saving offsets.
  return value;
 }
 function plan(original,incoming){
  const fields=original.fields;
  if(fields.length!==incoming.fields.length||fields.some(field=>!incoming.fields.includes(field)))
   throw Error('Direct RH CSV columns do not match the existing file.');
  const order=fields.map(field=>incoming.fields.indexOf(field));
  const date=fields.indexOf('request_date'),time=fields.indexOf('request_time'),tenant=fields.indexOf('tenant_name');
  const tenants=new Set(original.rows.map(row=>row[tenant]));
  let cutoff=-Infinity;
  for(const row of original.rows)cutoff=Math.max(cutoff,timestamp(row[date],row[time]));
  if(!Number.isFinite(cutoff))throw Error('The existing Direct RH CSV has no request timestamp.');
  const additions=[];
  for(const source of incoming.rows){
   const row=order.map(index=>source[index]);
   if(!tenants.has(row[tenant]))throw Error('Direct RH CSV tenant does not match the existing file.');
   const value=timestamp(row[date],row[time]);
   if(value>cutoff)additions.push({row,value});
  }
  additions.sort((a,b)=>a.value-b.value);
  return {rows:additions.map(entry=>entry.row),skipped:incoming.rows.length-additions.length,cutoff};
 }
 function merge(originalText,incomingText,papa){
  const original=parse(originalText,papa),incoming=parse(incomingText,papa);
  const result=plan(original,incoming);
  if(!result.rows.length)return {...result,text:originalText};
  const newline=originalText.includes('\r\n')?'\r\n':'\n';
  // Preserve every existing byte; serialize only the new records, with no extra header or metadata.
  const separator=/[\r\n]$/.test(originalText)?'':newline;
  return {...result,text:originalText+separator+papa.unparse(result.rows,{newline})+newline};
 }
 const api={recognises,parse,timestamp,plan,merge};
 if(typeof module==='object'&&module.exports)module.exports=api;
 else root.DirectRHCSV=api;
})(typeof globalThis==='object'?globalThis:this);
