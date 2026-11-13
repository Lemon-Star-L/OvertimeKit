/* Author: YiLong Li */
'use strict';
const $=id=>document.getElementById(id),esc=v=>String(v??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
const state={s:null,t:null};let run=null,decisions={},page=0,issuePage=0,busy=false;
const names={s:'基础数据表',t:'汇总表'},required={s:OT.sourceFields,t:['员工编码',...OT.fields]};
let noticeTimer=null;
function msg(t,type,duration){
 type=type||(/失败|超出范围/.test(t)?'error':/^正在/.test(t)?'loading':/完成|成功|已导入|已确认|已恢复|已清空|已发起/.test(t)?'success':'info');
 $('message').textContent=t;$('message').dataset.type=type;
 const toast=$('operationToast');clearTimeout(noticeTimer);toast.hidden=false;toast.dataset.type=type;
 $('toastTitle').textContent=({error:'操作未完成',loading:'正在处理',success:'操作成功',info:'操作提示'})[type];
 $('toastText').textContent=t;
 if(type!=='loading'&&type!=='error')noticeTimer=setTimeout(()=>{toast.hidden=true;},duration??6500);
}
$('dismissToast').onclick=()=>{clearTimeout(noticeTimer);$('operationToast').hidden=true;};
const paint=()=>new Promise(resolve=>setTimeout(resolve,30));
function reviewSummary(){
 $('matchSummary').textContent=run?`已完成匹配：基础数据 ${run.sourceCount} 行 / 汇总表 ${run.targetCount} 行；待确认 ${run.issues.length-Object.keys(decisions).length} 项，已确认 ${Object.keys(decisions).length} 项。`:'尚未执行匹配。';
}

function invalidate(){selectedIssues.clear();run=null;decisions={};issuePage=0;update();msg('数据或映射已更改，请重新执行匹配。');}
function guard(fn){return async(...args)=>{
 if(busy){msg('正在处理，请稍候再操作。');return;}
 busy=true;$('workspace').inert=true;$('workspace').setAttribute('aria-busy','true');$('clearWorkspace').disabled=true;
 const button=args[0]?.currentTarget?.tagName==='BUTTON'?args[0].currentTarget:null;
 const oldLabel=button?.textContent;
 if(button)button.textContent='正在处理…';
 try{await fn(...args);}catch(e){msg('操作失败：'+e.message,'error');}

 finally{await saveBrowserCache();if(button){button.textContent=oldLabel;if(button.id==='batchApply')renderIssueList(visibleIssues());}busy=false;$('workspace').inert=false;$('workspace').removeAttribute('aria-busy');$('clearWorkspace').disabled=false;}
};}
function fileStatus(side,text,kind='empty'){
 const label=$('filename-'+side);label.textContent=text;label.title=text;label.dataset.state=kind;
 $('pick-'+side).textContent=state[side]?'重新选择':'选择文件';
}
function clearSide(side){
 state[side]=null;$('file-'+side).value='';$('sheet-'+side).innerHTML='';
 $('header-'+side).value=1;$('map-'+side).innerHTML='';fileStatus(side,'尚未导入文件');
}
async function clearWorkspace(){
 if(busy){msg('正在处理，请完成后再清空工作区。');return;}
 if((state.s||state.t||run)&&!window.confirm('清空两张表、匹配结果及全部人工确认记录？\n尚未下载的进度将丢失，请先保存核对进度。\n原始文件和已下载文件不会删除。')){msg('已取消清空，当前数据和确认记录均已保留。','info');return;}
 selectedIssues.clear();clearSide('s');clearSide('t');run=null;decisions={};page=0;issuePage=0;
 $('previewSide').value='s';$('issueFilter').value='pending';$('search').value='';$('restoreProgress').value='';
 resetFilters();update();preview();await saveBrowserCache();msg('工作区已清空，可以重新导入文件。页面已解除对业务数据的引用，内存由浏览器自动回收；原始文件和已下载文件未删除。');
}
$('clearWorkspace').onclick=clearWorkspace;
for(const side of ['s','t']){
 const box=document.createElement('div');box.innerHTML=`<h3>${names[side]}</h3><div class="file-picker"><input type="file" id="file-${side}" accept=".xlsx,.xls" aria-label="${names[side]}" hidden><button type="button" id="pick-${side}">选择文件</button><span id="filename-${side}" class="file-status" role="status" aria-live="polite">尚未导入文件</span></div><label>工作表 <select id="sheet-${side}"></select></label><label>表头行 <input id="header-${side}" type="number" min="1" value="1"></label><div class="mapping" id="map-${side}"></div>`;$('imports').append(box);
 $('pick-'+side).onclick=()=>{if(!busy)$('file-'+side).click();};
 $('file-'+side).onchange=guard(async e=>{
 const file=e.target.files[0];if(!file)return;
 clearSide(side);invalidate();preview();fileStatus(side,'正在读取：'+file.name,'loading');
 msg('正在读取 '+file.name+'…');
 try{
  const wb=XLSX.read(await file.arrayBuffer(),{type:'array',cellText:true,cellDates:false});
  if(!wb.SheetNames.length)throw Error('文件中没有可用的工作表');
  state[side]={wb,name:file.name,map:[]};
  $('sheet-'+side).innerHTML=wb.SheetNames.map(n=>`<option>${esc(n)}</option>`).join('');
  $('header-'+side).value=1;loadSheet(side,true);fileStatus(side,'已导入：'+file.name,'loaded');
  msg('导入成功：'+file.name+'，共 '+state[side].rows.length+' 行。请检查字段映射。','success');
 }catch(e){clearSide(side);invalidate();preview();fileStatus(side,'导入失败：'+file.name+'（请重新选择）','error');throw e;}
 });
 $('sheet-'+side).onchange=guard(()=>loadSheet(side,true));$('header-'+side).onchange=guard(()=>loadSheet(side,false));
}
function loadSheet(side,auto){
 const st=state[side];if(!st)return;const ws=st.wb.Sheets[$('sheet-'+side).value];
 st.sheet=$('sheet-'+side).value;st.ws=ws;st.raw=XLSX.utils.sheet_to_json(ws,{header:1,defval:'',raw:true,blankrows:true,range:0});
 if(auto){let best=0,score=-1;st.raw.slice(0,50).forEach((r,i)=>{const n=required[side].filter(f=>r.some(v=>OT.clean(v)===f)).length;if(n>score){best=i;score=n;}});$('header-'+side).value=best+1;}
 const h=Number($('header-'+side).value);invalidate();if(!Number.isInteger(h)||h<1||h>st.raw.length){st.headers=[];st.rows=[];st.map=[];$('map-'+side).innerHTML='';update();preview();msg('表头行超出范围，请填写有效的 Excel 行号。');return;}
 st.header=h;st.headers=st.raw[h-1].map(OT.clean);st.rows=st.raw.slice(h).map((values,i)=>({row:h+i+1,values:values.slice()})).filter(r=>r.values.some(v=>OT.clean(v)!==''));
 st.map=required[side].map(f=>st.headers.indexOf(f));$('map-'+side).innerHTML=required[side].map((f,i)=>`<label>${esc(f)}<select id="m-${side}-${i}"><option value="-1">请选择字段</option>${st.headers.map((v,j)=>`<option value="${j}" ${st.map[i]===j?'selected':''}>${XLSX.utils.encode_col(j)} · ${esc(v||'未命名')}</option>`).join('')}</select></label>`).join('');
 required[side].forEach((f,i)=>$('m-'+side+'-'+i).onchange=e=>{st.map[i]=Number(e.target.value);invalidate();saveBrowserCache();});page=0;preview();update();msg(names[side]+'已切换至“'+st.sheet+'”，表头第 '+st.header+' 行，共 '+st.rows.length+' 行。请确认映射后重新执行匹配。','info');
}
function rowsFor(side){const st=state[side];return st.rows.map(r=>{const values=r.values.slice(),col=st.map[0],cell=st.ws[XLSX.utils.encode_cell({r:r.row-1,c:col})];if(cell){if(cell.t==='n'&&cell.z&&/^0+$/.test(cell.z))values[col]=String(cell.v).padStart(cell.z.length,'0');else if(cell.t==='n'&&/^0\d+$/.test(cell.w||''))values[col]=cell.w;else values[col]=OT.clean(cell.v);}for(const c of st.map){const cell=st.ws[XLSX.utils.encode_cell({r:r.row-1,c})];if(cell?.f&&(cell.v===undefined||cell.v===null))values[c]='[公式无缓存]';if(cell?.t==='e')values[c]='[Excel错误]';}return {...r,values};});}
function preview(){const st=state[$('previewSide').value];if(!st?.rows){$('previewTable').textContent='等待导入';$('pageInfo').textContent='';$('prev').disabled=true;$('next').disabled=true;return;}const pages=Math.max(1,Math.ceil(st.rows.length/30));page=Math.max(0,Math.min(page,pages-1));$('pageInfo').textContent=`${page+1} / ${pages} 页 · ${st.rows.length} 行`;$('previewTable').innerHTML='<table><thead><tr><th>Excel 行号</th>'+st.headers.map(h=>'<th>'+esc(h)+'</th>').join('')+'</tr></thead><tbody>'+st.rows.slice(page*30,page*30+30).map(r=>'<tr><td>'+r.row+'</td>'+st.headers.map((_,i)=>'<td title="'+esc(r.values[i])+'">'+esc(r.values[i])+'</td>').join('')+'</tr>').join('')+'</tbody></table>';$('prev').disabled=page===0;$('next').disabled=page>=pages-1;}
function previewFeedback(){msg(names[$('previewSide').value]+'：'+($('pageInfo').textContent||'尚未导入数据'),'info',2200);}
$('previewSide').onchange=()=>{page=0;preview();previewFeedback();};$('prev').onclick=()=>{page--;preview();previewFeedback();};$('next').onclick=()=>{page++;preview();previewFeedback();};
$('run').onclick=guard(async()=>{for(const s of ['s','t']){const st=state[s];if(!st?.rows?.length)throw Error('请导入有效的'+names[s]);if(st.map.some(i=>i<0)||new Set(st.map).size!==st.map.length)throw Error(names[s]+'字段未完整匹配或重复映射');}msg('正在累计四类加班时数并核对差异，请稍候…','loading');await paint();run=OT.reconcile(rowsFor('s'),rowsFor('t'),state.s.map,state.t.map);decisions={};selectedIssues.clear();for(const [k] of filterColumns)$('filter-'+k).value='';issuePage=0;$('search').value='';$('issueFilter').value='pending';update();$('review').scrollIntoView({behavior:'smooth',block:'start'});msg(`匹配完成：基础数据 ${run.sourceCount} 行，汇总表 ${run.targetCount} 行，共 ${run.issues.length} 项需人工确认。其中 ${run.invalidCount} 条无效基础数据暂未参与汇总。${run.issues.length?'请在下方逐项确认。':'无需人工确认，可以直接导出最终结果。'}`,'success',10000);});
const selectedIssues=new Set();
const filterColumns=[['status','状态'],['id','编号'],['employee','员工编码'],['row','Excel 行号'],['field','字段'],['reason','异常原因'],['original','汇总表现有值（小时）'],['computed','基础表计算值（小时）'],['delta','差值（计算－现有）']];
function targetResult(i){return i.kind==='source'||i.kind==='unmatched'?null:run?.results.find(r=>r.row===i.row);}
function comparisonRows(i){
 const r=targetResult(i);
 if(i.kind==='cell')return [{field:i.field,existing:i.original,computed:i.computed}];
 if(r){const old=r.existing||state.t?.map?.slice(1).map(c=>r.values[c])||[];return OT.fields.map((field,k)=>({field,existing:old[k]??'',computed:i.kind==='target'?null:r.computed[k]}));}
 if(i.kind==='unmatched')return OT.fields.map((field,k)=>({field,existing:null,computed:i.computed[k]}));
 return [];
}
function diffValue(row){const old=OT.number(row.existing,true);return row.existing===null||old===null||typeof row.computed!=='number'?null:Math.round((row.computed-old)*1e8)/1e8;}
function hourText(v){return v===null||v===undefined?'无法比较':v===''?'0（原为空）':String(v);}
function reasonText(i){
 if(i.kind==='missing'){
  const r=targetResult(i);let count=r?.sourceRecords;
  if(count===undefined&&state.s?.rows)count=state.s.rows.filter(row=>OT.clean(row.values[state.s.map[0]])===i.employee).length;
  return count>0?'基础表有记录，但没有有效加班时数（明细被排除）':'基础表未找到该员工的加班记录（按 0 汇总）';
 }
 if(i.kind==='cell'&&i.reason==='时数不一致')return '汇总表现有时数与基础表计算时数不一致';
 return i.reason;
}
function displayField(i){return i.kind==='missing'||i.kind==='target'||i.kind==='unmatched'?'四类加班时数':i.field;}
function issueValue(i,key){
 if(key==='status')return decisions[i.id]?'已确认':'待确认';
 if(key==='reason')return reasonText(i);if(key==='field')return displayField(i);
 if(['original','computed','delta'].includes(key)){
  const rows=comparisonRows(i);if(!rows.length)return '不适用（明细数据异常）';
  return rows.map(r=>(rows.length>1?r.field+'：':'')+hourText(key==='original'?r.existing:key==='computed'?r.computed:diffValue(r))).join('\n');
 }
 const v=i[key];return v==null?'不适用':String(v);
}
function comparisonCard(i){
 const rows=comparisonRows(i);if(!rows.length)return '<p>该项为基础明细异常，未计入汇总。原始记录：'+esc(i.original)+'</p>';
 const r=targetResult(i),counts=r?.sourceRecords===undefined?'':`基础表记录 ${r.sourceRecords} 条，其中有效 ${r.validRecords} 条。`;
 return `<p>${esc(counts)}${i.kind==='missing'?'计算值为当前有效明细汇总，0 不代表已确认该员工没有加班，请检查是否漏导数据。':''}</p><div class="scroll"><table class="comparison-table"><thead><tr><th>加班项目</th><th>汇总表现有值（小时）</th><th>基础表计算值（小时）</th><th>差值（计算－现有）</th></tr></thead><tbody>${rows.map(x=>`<tr><td>${esc(x.field)}</td><td>${esc(hourText(x.existing))}</td><td>${esc(hourText(x.computed))}</td><td class="${diffValue(x)?'warn':''}">${esc(hourText(diffValue(x)))}</td></tr>`).join('')}</tbody></table></div>`;
}
function visibleIssues(){return run?run.issues.filter(i=>($('issueFilter').value==='all'||($('issueFilter').value==='done'?!!decisions[i.id]:!decisions[i.id]))&&(JSON.stringify(i)+' '+reasonText(i)).includes($('search').value.trim())&&filterColumns.every(([k])=>{const v=$('filter-'+k).value.trim();return !v||(['reason','field'].includes(k)?issueValue(i,k)===v:issueValue(i,k).toLowerCase().includes(v.toLowerCase()));})):[];}
function resetFilters(){ $('batchScope').value='selected';for(const [k] of filterColumns)$('filter-'+k).value='';$('search').value='';$('issueFilter').value='all';selectedIssues.clear();issuePage=0;renderIssue();}
$('columnFilters').innerHTML=filterColumns.map(([k,label])=>`<label>${label}${['reason','field'].includes(k)?`<select id="filter-${k}" aria-label="筛选${label}"><option value="">全部${label}</option></select>`:`<input id="filter-${k}" placeholder="包含…" aria-label="筛选${label}">`}</label>`).join('');
for(const [k] of filterColumns)$('filter-'+k).oninput=()=>{selectedIssues.clear();issuePage=0;renderIssue();};
$('resetFilters').onclick=resetFilters;

function refreshFilterChoices(){
 for(const k of ['reason','field']){
  const select=$('filter-'+k),previous=select.value,values=[...new Set((run?.issues||[]).map(i=>issueValue(i,k)))].sort((a,b)=>a.localeCompare(b,'zh-CN'));
  select.innerHTML='<option value="">全部'+(k==='reason'?'异常原因':'字段')+'</option>'+values.map(v=>`<option value="${esc(v)}">${esc(v)}</option>`).join('');
  select.value=values.includes(previous)?previous:'';
 }
}
function update(){refreshFilterChoices();reviewSummary(); $('ks').textContent=state.s?.rows?.length||0;$('kt').textContent=state.t?.rows?.length||0;const count=Object.keys(decisions).length;$('ki').textContent=run?run.issues.length-count:0;$('kd').textContent=count;$('exportFinal').disabled=!run;for(const id of ['exportAudit','saveProgress'])$(id).disabled=!run;renderIssue();}
const ISSUE_LIST_SIZE=20;
function renderIssueList(list){
 const active=document.activeElement,focusId=active?.id,start=active?.selectionStart,end=active?.selectionEnd;
 const scroll=$('issueList').scrollLeft,vertical=$('issueList').scrollTop;
 // Move the same input nodes rather than replacing them; preserve values and event handlers.
 for(const [k] of filterColumns)$('columnFilters').appendChild($('filter-'+k));
 const listPage=Math.floor(issuePage/ISSUE_LIST_SIZE),pages=Math.max(1,Math.ceil(list.length/ISSUE_LIST_SIZE));
 const pageIds=new Set(list.slice(listPage*ISSUE_LIST_SIZE,(listPage+1)*ISSUE_LIST_SIZE).map(i=>i.id));
 for(const id of selectedIssues)if(!pageIds.has(id))selectedIssues.delete(id);
 $('batchCount').textContent='已选 '+selectedIssues.size+' 项 · 筛选结果 '+list.length+' 项';
 $('batchApply').textContent=$('batchScope').value==='filtered'?'处理全部筛选结果（'+list.length+' 项）':'处理勾选项';
 $('issueListInfo').textContent=`共 ${list.length} 项 · 第 ${listPage+1} / ${pages} 页 · 每页 ${ISSUE_LIST_SIZE} 项`;
 $('listPrev').disabled=!list.length||listPage===0;$('listNext').disabled=!list.length||listPage>=pages-1;
 $('issueList').innerHTML='<table><thead><tr><th><input type="checkbox" id="selectPage" aria-label="全选当前页"></th><th>操作</th>'+filterColumns.map(([k,label])=>`<th><span>${esc(label)}</span><div id="headfilter-${k}" class="header-filter"></div></th>`).join('')+'</tr></thead><tbody>'+list.slice(listPage*ISSUE_LIST_SIZE,(listPage+1)*ISSUE_LIST_SIZE).map((i,index)=>{
 const pos=listPage*ISSUE_LIST_SIZE+index;
 return `<tr class="${pos===issuePage?'selected':''}"><td><input type="checkbox" data-select-id="${esc(i.id)}" aria-label="选择${esc(i.id)}" ${selectedIssues.has(i.id)?'checked':''}></td><td><button class="issue-open" type="button" data-issue-index="${pos}">${pos===issuePage?'当前项':'查看 / 核对'}</button></td>${filterColumns.map(([k])=>`<td class="${['original','computed','delta'].includes(k)?'comparison-cell':k==='reason'?'reason':''}" title="${esc(issueValue(i,k))}">${esc(issueValue(i,k))}</td>`).join('')}</tr>`;
 }).join('')+(!list.length?'<tr><td colspan="11">'+(run?'没有符合条件的项目，请调整表头筛选。':'执行匹配后展示异常对比。')+'</td></tr>':'')+'</tbody></table>';
 for(const [k] of filterColumns)$('headfilter-'+k).appendChild($('filter-'+k));
 $('selectPage').disabled=!list.length;$('selectPage').checked=pageIds.size>0&&selectedIssues.size===pageIds.size;$('selectPage').indeterminate=selectedIssues.size>0&&selectedIssues.size<pageIds.size;
 if(focusId?.startsWith('filter-')){const node=$(focusId);node.focus({preventScroll:true});if(node.setSelectionRange&&typeof start==='number')node.setSelectionRange(start,end);}
 $('issueList').scrollLeft=scroll;$('issueList').scrollTop=vertical;
}
$('issueList').onchange=e=>{const ids=visibleIssues().slice(Math.floor(issuePage/ISSUE_LIST_SIZE)*ISSUE_LIST_SIZE,(Math.floor(issuePage/ISSUE_LIST_SIZE)+1)*ISSUE_LIST_SIZE).map(i=>i.id);if(e.target.id==='selectPage'){for(const id of ids)e.target.checked?selectedIssues.add(id):selectedIssues.delete(id);}else if(e.target.dataset.selectId){e.target.checked?selectedIssues.add(e.target.dataset.selectId):selectedIssues.delete(e.target.dataset.selectId);}if(e.target.id==='selectPage'||e.target.dataset.selectId)renderIssueList(visibleIssues());};
$('batchScope').onchange=()=>{renderIssueList(visibleIssues());msg($('batchScope').value==='filtered'?'批处理范围已切换为全部筛选结果，包含其他页。':'批处理范围已切换为当前页勾选项。','info');};
$('batchAction').onchange=()=>{$('batchValue').hidden=$('batchAction').value!=='manual';};
$('batchApply').onclick=guard(()=>{
 const allFiltered=$('batchScope').value==='filtered',selected=allFiltered?visibleIssues():visibleIssues().filter(i=>selectedIssues.has(i.id));if(!selected.length)throw Error(allFiltered?'当前筛选结果为空':'请先勾选当前页需要处理的项目');
 const mode=$('batchAction').value,value=$('batchValue').value,note=$('batchNote').value.trim(),next={};
 for(const i of selected){let action=mode;if(i.kind!=='cell'){if(mode!=='keep'&&mode!=='ack')throw Error('所选项含非时数异常，请使用“保留原值 / 确认已核查”，或按字段筛选后操作');action='ack';}else if(mode==='ack')throw Error('时数差异不能只确认已核查，请选择保留、计算值或人工值');next[i.id]={action,value,note,time:new Date().toISOString()};}
 OT.finalize({...run,issues:selected},next);
 if(!window.confirm('处理范围：'+(allFiltered?'全部筛选结果（跨页）':'当前页勾选项')+'\n异常原因：'+($('filter-reason').value||'全部')+'\n字段：'+($('filter-field').value||'全部')+'\n共 '+selected.length+' 项，是否继续批量处理？')){msg('已取消批量处理。','info');return;}
 Object.assign(decisions,next);selectedIssues.clear();update();msg('批量处理完成：已确认 '+selected.length+' 项。','success');
});

$('issueList').onclick=e=>{const button=e.target.closest('button[data-issue-index]');if(!button)return;issuePage=Number(button.dataset.issueIndex);renderIssue();$('reviewBox').scrollIntoView({behavior:'smooth',block:'start'});msg('已定位到 '+visibleIssues()[issuePage].id+'，请在核对卡片中处理。','info',2400);};
$('listPrev').onclick=()=>{issuePage=Math.max(0,(Math.floor(issuePage/ISSUE_LIST_SIZE)-1)*ISSUE_LIST_SIZE);renderIssue();};
$('listNext').onclick=()=>{issuePage=Math.min(visibleIssues().length-1,(Math.floor(issuePage/ISSUE_LIST_SIZE)+1)*ISSUE_LIST_SIZE);renderIssue();};
function renderIssue(){const list=visibleIssues();issuePage=Math.max(0,Math.min(issuePage,list.length-1));renderIssueList(list);$('issueInfo').textContent=list.length?`${issuePage+1} / ${list.length} 项`:'';$('issuePrev').disabled=issuePage===0;$('issueNext').disabled=issuePage>=list.length-1;const i=list[issuePage];if(!i){$('reviewBox').textContent=run?'当前筛选下没有待核对项目。':'尚未执行匹配';return;}const d=decisions[i.id];
 $('reviewBox').innerHTML=`<span class="pill">${d?'已确认':'待确认'} · ${esc(i.id)}</span><h3>员工 ${esc(i.employee||'（空）')} · ${esc(displayField(i))}</h3><p class="warn">${esc(reasonText(i))}${i.row?' · Excel 第 '+i.row+' 行':''}</p>${comparisonCard(i)}<label>处理方式 <select id="action">${i.kind==='cell'?'<option value="keep">保留原值</option>'+(typeof i.computed==='number'?'<option value="computed">采用计算值</option>':'')+'<option value="manual">输入人工值</option>':'<option value="ack">'+(i.kind==='source'?'确认排除此无效明细':i.kind==='unmatched'?'确认记录至未匹配员工表':'确认已核查，四类时数按各差异项处理')+'</option>'}</select></label><label id="manualLabel">人工时数 <input id="manual" type="number" min="0" step="any"></label><label>处理说明（选填）<input id="note" placeholder="填写核实依据或保留、排除原因"></label><button class="primary" id="confirm">确认本项</button>`;
 if(d){$('action').value=d.action;$('manual').value=d.value??'';$('note').value=d.note;}$('action').onchange=()=>{$('manualLabel').hidden=$('action').value!=='manual';};$('action').onchange();
 $('confirm').onclick=guard(()=>{const action=$('action').value,note=$('note').value.trim(),value=$('manual').value;if(action==='manual'&&OT.number(value)===null)throw Error('请输入有效的非负时数');decisions[i.id]={action,note,value,time:new Date().toISOString()};update();const remaining=run.issues.length-Object.keys(decisions).length;msg(`已确认 ${i.id}（员工 ${i.employee||'编码为空'}）。${remaining?'还剩 '+remaining+' 项待确认。':'全部确认完成，可以导出最终 Excel。'}`,'success');});}
for(const id of ['issueFilter','search'])$(id).oninput=()=>{selectedIssues.clear();issuePage=0;renderIssue();};$('issuePrev').onclick=()=>{issuePage--;renderIssue();};$('issueNext').onclick=()=>{issuePage++;renderIssue();};
function audit(){return [['项目编号','类型','Excel行号','员工编码','字段','原因','汇总表现有值（小时）','基础表计算值（小时）','差值（计算－现有）','处理方式','人工值','处理说明','确认时间'],...run.issues.map(i=>{const d=decisions[i.id]||{};return [i.id,i.kind,i.row,i.employee,displayField(i),reasonText(i),issueValue(i,'original'),issueValue(i,'computed'),issueValue(i,'delta'),d.action||'待确认',d.value,d.note,d.time];})];}
function append(wb,name,data){const ws=XLSX.utils.aoa_to_sheet(data);ws['!cols']=(data[0]||[]).map(()=>({wch:20}));XLSX.utils.book_append_sheet(wb,ws,name);}
$('exportAudit').onclick=guard(async()=>{msg('正在生成待核对清单…','loading');await paint();const wb=XLSX.utils.book_new();append(wb,'待核对清单',audit());XLSX.writeFile(wb,'OvertimeKit_待核对清单.xlsx');msg('已发起待核对清单下载，请查看浏览器下载列表。','success');});
$('exportFinal').onclick=guard(async()=>{msg('正在生成最终 Excel…','loading');await paint();if(!run)throw Error('请先执行匹配');const pending=run.issues.filter(i=>!decisions[i.id]).length;if(pending&&!window.confirm('还有 '+pending+' 项未完成核对。未确认项保留原值，并在核对记录中标为待确认。是否仍要导出？')){msg('已取消导出，请继续核对。','info');return;}const rows=OT.finalize(run,decisions,true),data=state.t.raw.map(r=>r.slice());for(const r of rows)data[r.row-1]=r.values;const wb=XLSX.utils.book_new();append(wb,'核对结果',data);append(wb,'核对记录',audit());append(wb,'基础数据汇总',[['员工编码',...OT.fields],...run.sums.map(([id,v])=>[id,...v])]);append(wb,'未匹配员工',[['员工编码',...OT.fields],...run.issues.filter(i=>i.kind==='unmatched').map(i=>[i.employee,...i.computed])]);append(wb,'导出说明',[['项目','内容'],['核对状态',pending?'未全部核对':'已全部核对'],['待确认项数',pending],['未确认处理','保留汇总表原值；无效明细仍未计入汇总，请查看核对记录'],['导出时间',new Date().toISOString()]]);XLSX.writeFile(wb,pending?'OvertimeKit_核对结果_未全部核对.xlsx':'OvertimeKit_最终核对结果.xlsx');msg('已发起 Excel 下载。'+(pending?'仍有 '+pending+' 项未确认，已在文件中标记。':'全部项目已确认。'),'success');});
function fingerprint(){return JSON.stringify(['OvertimeKit-1',state.s.name,state.t.name,state.s.sheet,state.t.sheet,state.s.header,state.t.header,state.s.map,state.t.map,state.s.raw,state.t.raw]);}
async function digest(){const bytes=new TextEncoder().encode(fingerprint());return Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256',bytes)),x=>x.toString(16).padStart(2,'0')).join('');}
$('saveProgress').onclick=guard(async()=>{msg('正在生成核对进度文件…','loading');await paint();const payload={version:1,fingerprint:await digest(),decisions};const a=document.createElement('a'),url=URL.createObjectURL(new Blob([JSON.stringify(payload,null,2)],{type:'application/json'}));a.href=url;a.download='OvertimeKit_核对进度.json';a.click();setTimeout(()=>URL.revokeObjectURL(url),1000);msg('已发起核对进度下载，包含 '+Object.keys(decisions).length+' 项确认记录。请检查浏览器下载列表；恢复时需导入相同数据和配置。');});
$('restoreProgress').onchange=guard(async e=>{const f=e.target.files[0];if(!f)return;try{msg('正在校验数据并恢复进度…','loading');await paint();if(!run)throw Error('先导入原文件并执行匹配');const p=JSON.parse(await f.text());if(p.version!==1||p.fingerprint!==await digest())throw Error('进度与当前数据或配置不一致');const valid={};for(const i of run.issues){const d=p.decisions?.[i.id];if(!d)continue;const partial={...run,issues:[i]};OT.finalize(partial,{[i.id]:d});valid[i.id]=d;}decisions=valid;update();msg('进度恢复成功：已恢复 '+Object.keys(decisions).length+' 项确认记录，剩余 '+(run.issues.length-Object.keys(decisions).length)+' 项。','success');}finally{e.target.value='';}});
let cacheReady=false,cacheQueue=Promise.resolve();
const CACHE_DB='OvertimeKit_Workspace_v1',CACHE_KEY='current';
function cacheStatus(t){$('cacheStatus').textContent=t;}
function openCache(){return new Promise((resolve,reject)=>{if(!globalThis.indexedDB){reject(Error('浏览器不支持本地数据库'));return;}const req=indexedDB.open(CACHE_DB,1);req.onupgradeneeded=()=>{if(!req.result.objectStoreNames.contains('workspace'))req.result.createObjectStore('workspace');};req.onsuccess=()=>resolve(req.result);req.onerror=()=>reject(req.error);req.onblocked=()=>reject(Error('缓存被其他窗口占用，请关闭其他工作台页面'));});}
async function cacheTransaction(mode,action){const db=await openCache();try{return await new Promise((resolve,reject)=>{const tx=db.transaction('workspace',mode);const req=action(tx.objectStore('workspace'));let result;req.onsuccess=()=>{result=req.result;};tx.oncomplete=()=>resolve(result);tx.onerror=()=>reject(tx.error||req.error);tx.onabort=()=>reject(tx.error||Error('缓存事务被取消'));});}finally{db.close();}}
function saveBrowserCache(){
 if(!cacheReady)return Promise.resolve();
 let snapshot;try{snapshot=structuredClone({version:1,s:state.s,t:state.t,run,decisions,savedAt:new Date().toISOString()});}catch(e){cacheStatus('缓存失败：'+e.message+'；请手动保存进度');return Promise.resolve();}
 cacheStatus('正在保存浏览器缓存…');
 cacheQueue=cacheQueue.catch(()=>{}).then(()=>cacheTransaction('readwrite',store=>store.put(snapshot,CACHE_KEY))).then(()=>cacheStatus('已自动缓存 · '+new Date(snapshot.savedAt).toLocaleTimeString()),e=>cacheStatus('缓存失败：'+e.message+'；请使用进度 JSON 备份'));
 return cacheQueue;
}
async function restoreBrowserCache(){
 busy=true;$('workspace').inert=true;$('clearWorkspace').disabled=true;$('clearBrowserCache').disabled=true;cacheStatus('正在检查浏览器缓存…');
 try{const saved=await cacheTransaction('readonly',store=>store.get(CACHE_KEY));
 if(saved){if(saved.version!==1)throw Error('缓存版本不兼容');
 for(const side of ['s','t']){const st=saved[side];if(st&&(!st.wb||!Array.isArray(st.map)||!Array.isArray(st.rows)))throw Error('缓存结构不完整');}
 if(saved.run)OT.finalize(saved.run,saved.decisions||{},true);
 state.s=saved.s;state.t=saved.t;run=saved.run;decisions=saved.decisions||{};
 for(const side of ['s','t']){const st=state[side];if(!st)continue;fileStatus(side,'已从缓存恢复：'+st.name,'cached');$('sheet-'+side).innerHTML=st.wb.SheetNames.map(n=>`<option ${n===st.sheet?'selected':''}>${esc(n)}</option>`).join('');$('header-'+side).value=st.header;
 $('map-'+side).innerHTML=required[side].map((f,i)=>`<label>${esc(f)}<select id="m-${side}-${i}"><option value="-1">请选择字段</option>${st.headers.map((v,j)=>`<option value="${j}" ${st.map[i]===j?'selected':''}>${XLSX.utils.encode_col(j)} · ${esc(v||'未命名')}</option>`).join('')}</select></label>`).join('');
 required[side].forEach((f,i)=>$('m-'+side+'-'+i).onchange=e=>{st.map[i]=Number(e.target.value);invalidate();saveBrowserCache();});}
 update();preview();msg('已恢复浏览器缓存：两张表、字段配置与确认进度已载入。','success');cacheStatus('已恢复缓存 · '+new Date(saved.savedAt).toLocaleString());
 }else cacheStatus('暂无缓存；导入及处理后将自动保存');
 }catch(e){cacheStatus('无法恢复缓存：'+e.message+'；可重新导入文件或清空缓存');}
 finally{cacheReady=true;busy=false;$('workspace').inert=false;$('clearWorkspace').disabled=false;$('clearBrowserCache').disabled=false;}
}
$('clearBrowserCache').onclick=async()=>{
 if(busy){msg('正在处理，请稍候清理缓存。','info');return;}
 if(!window.confirm('仅删除本工具在当前浏览器保存的缓存？当前页面数据和已下载文件保留；后续数据操作将重新自动缓存。'))return;
 busy=true;$('workspace').inert=true;$('clearWorkspace').disabled=true;$('clearBrowserCache').disabled=true;
 try{await cacheQueue;await cacheTransaction('readwrite',store=>store.delete(CACHE_KEY));cacheStatus('浏览器缓存已清空；下一次数据操作后重新保存');msg('浏览器缓存已清空，当前页面和已下载文件保留。','success');}
 catch(e){msg('清理缓存失败：'+e.message,'error');}
 finally{busy=false;$('workspace').inert=false;$('clearWorkspace').disabled=false;$('clearBrowserCache').disabled=false;}
};
update();preview();restoreBrowserCache();
