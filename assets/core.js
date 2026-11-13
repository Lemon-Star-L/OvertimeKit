/* Author: YiLong Li */
(function(root){
'use strict';
const fields=['加班用于调休','加点时数','周末加班时数','节假日加班时数'];
const sourceFields=['员工编码','申请加班小时数','加班补偿方式','加班类型'];
const clean=v=>String(v??'').trim();
function number(v,blank=false){if(clean(v)==='')return blank?0:null;const s=clean(v);if(!/^(?:\d+|\d{1,3}(?:,\d{3})+)(?:\.\d+)?$/.test(s))return null;const n=Number(s.replaceAll(',',''));return Number.isFinite(n)&&n>=0?n:null;}
const round=n=>Math.round((n+Number.EPSILON)*1e8)/1e8;
function reconcile(src,dst,sm,tm){
 const sourceCounts=new Map(),validCounts=new Map(),sums=new Map(),issues=[],results=[],targetIds=new Map();
 const issue=o=>{o.id='I'+issues.length;issues.push(o);};
 for(const r of src){
  const id=clean(r.values[sm[0]]),h=number(r.values[sm[1]]),pay=clean(r.values[sm[2]]),type=clean(r.values[sm[3]]);
  if(id)sourceCounts.set(id,(sourceCounts.get(id)||0)+1);
  const category=pay==='调休'?0:pay==='加班费'?({'工作日加班':1,'休息日加班':2,'法定节假日加班':3}[type]??-1):-1;
  if(!id||h===null||category<0){issue({kind:'source',row:r.row,employee:id,field:'基础数据',reason:!id?'员工编码为空':h===null?'小时数为空、负数或无效':'补偿方式或加班类型不识别',original:JSON.stringify(r.values),computed:null});continue;}
  validCounts.set(id,(validCounts.get(id)||0)+1);
  if(!sums.has(id))sums.set(id,[0,0,0,0]);sums.get(id)[category]=round(sums.get(id)[category]+h);
 }
 for(const r of dst){const id=clean(r.values[tm[0]]);if(id)targetIds.set(id,(targetIds.get(id)||0)+1);}
 for(const r of dst){
  const id=clean(r.values[tm[0]]),duplicate=(targetIds.get(id)||0)>1,missing=!sums.has(id),calc=sums.get(id)||[0,0,0,0];
  const result={row:r.row,employee:id,values:r.values.slice(),existing:tm.slice(1).map(c=>r.values[c]??''),sourceRecords:sourceCounts.get(id)||0,validRecords:validCounts.get(id)||0,computed:calc.slice()};results.push(result);
  if(!id||duplicate){issue({kind:'target',row:r.row,employee:id,field:'员工编码',reason:!id?'汇总表员工编码为空，保留整行':'汇总表员工编码重复，四类时数需逐项人工确认',original:id,computed:null});}
  if(id&&missing)issue({kind:'missing',row:r.row,employee:id,field:'员工编码',reason:'基础数据无此员工记录，请确认四类汇总应为零或保留原值',original:id,computed:null});
  fields.forEach((field,i)=>{
   const old=r.values[tm[i+1]],n=number(old,true);
   if(!id||duplicate||n===null||Math.abs(n-calc[i])>1e-7){issue({kind:'cell',row:r.row,employee:id,field,column:tm[i+1],reason:duplicate?'重复员工编码，禁止自动分配':!id?'员工编码为空':n===null?'原时数无效':'时数不一致',original:old??'',computed:(!id||duplicate)?null:calc[i]});}
  });
 }
 for(const [id,v] of sums)if(!targetIds.has(id))issue({kind:'unmatched',row:null,employee:id,field:'四类时数',reason:'汇总表无此员工，记录至未匹配员工表，不自动新增',original:'',computed:v});
 return {sums:[...sums],issues,results,sourceCount:src.length,targetCount:dst.length,invalidCount:issues.filter(i=>i.kind==='source').length};
}
function finalize(run,decisions,allowPending=false){
 const rows=run.results.map(r=>({...r,values:r.values.slice()}));
 for(const i of run.issues){const d=decisions[i.id];if(!d){if(allowPending)continue;throw Error('仍有未确认项目');}
  if(i.kind==='cell'){if(!['keep','computed','manual'].includes(d.action))throw Error('处理方式无效');const r=rows.find(r=>r.row===i.row);if(d.action==='computed'){if(typeof i.computed!=='number')throw Error('该项目不能采用计算值');r.values[i.column]=i.computed;}if(d.action==='manual'){const n=number(d.value);if(n===null)throw Error('人工值必须为非负有限数字');r.values[i.column]=n;}}
  else if(d.action!=='ack')throw Error('异常尚未确认');
 }
 return rows;
}
const api={fields,sourceFields,clean,number,reconcile,finalize};if(typeof module!=='undefined')module.exports=api;root.OT=api;
})(typeof globalThis!=='undefined'?globalThis:this);
