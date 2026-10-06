'use strict';
let state=null, templates=[], busy=false;
const $=id=>document.getElementById(id);
const VALUE_TYPES=['xs:string','xs:double','xs:float','xs:decimal','xs:integer','xs:int','xs:long','xs:short','xs:byte','xs:unsignedInt','xs:boolean','xs:date','xs:dateTime','xs:time','xs:anyURI'];
const SHORT={Submodel:'Submodel',SubmodelElementCollection:'Collection',SubmodelElementList:'List',Property:'Property',MultiLanguageProperty:'Multi-language',Entity:'Entity'};
const VALUE_KINDS=['Property','MultiLanguageProperty'];
const ROW_ITEMS=['SubmodelElementCollection','Entity'];
const GLOBAL_ASSET='#globalAssetId';

function el(tag,value,cls){const node=document.createElement(tag);if(value!==undefined&&value!==null)node.textContent=String(value);if(cls)node.className=cls;return node;}
function message(text,error=false,link){const node=$('message');node.replaceChildren(el('span',text));if(link)node.append(document.createTextNode(' '),link);node.hidden=false;node.className=error?'error':'success';}
function ready(){return !!(state&&state.excel&&state.aas&&state.aas.shell&&state.rules.length);}
function buttons(){$('checkButton').disabled=busy||!ready();$('importButton').disabled=busy||!ready();for(const id of ['excelFile','aasFile','templateFile','rulesFile'])$(id).disabled=busy;}
async function task(label,fn){if(busy)return;busy=true;$('busyStatus').textContent=label;buttons();try{await fn();}catch(e){message(e.message,true);}finally{busy=false;$('busyStatus').textContent='Ready';buttons();}}
async function request(url,options={}){const resp=await fetch(url,options);let data={};try{data=await resp.json();}catch(e){}if(!resp.ok){if(data.report)renderReport(data.report);throw new Error(data.error||`Request failed (${resp.status})`);}return data;}
function post(url,body={}){return request(url,{method:'POST',headers:{'Content-Type':'application/json','X-CSRF-Token':state.csrf},body:JSON.stringify({...body,revision:state.revision})});}
function upload(url,file,withRevision=true){const form=new FormData();form.append('file',file);if(withRevision)form.append('revision',state.revision);return request(url,{method:'POST',headers:{'X-CSRF-Token':state.csrf},body:form});}
async function apply(promise,text){state=await promise;render();if(text)message(text);}

const key=target=>JSON.stringify(target);
function describe(target){return [target.submodel_idShort||target.submodel,...target.path.map(s=>typeof s==='number'?`[${s}]`:s)].join(' / ');}
function fill(select,options,placeholder){const previous=select.value;select.replaceChildren();if(placeholder!==undefined){const o=el('option',placeholder);o.value='';select.append(o);}for(const [value,label,disabled] of options){const o=el('option',label);o.value=value;o.disabled=!!disabled;select.append(o);}if([...select.options].some(o=>o.value===previous&&!o.disabled))select.value=previous;}
function sanitize(name){let s=String(name).trim().replace(/[^A-Za-z0-9_]/g,'_');if(!/^[A-Za-z]/.test(s))s='C_'+s;return s.slice(0,128);}
function nodes(){return state.aas?.tree||[];}
function sheets(){return (state.excel?.order||[]).map(name=>[name,state.excel.sheets[name]]).filter(([,s])=>!s.problem&&s.columns.length);}
function childOf(node,parent){const a=node.target.path,b=parent.target.path;return node.target.submodel===parent.target.submodel&&a.length>b.length&&b.every((s,i)=>s===a[i]);}

/* Step 1: Excel */
function renderExcel(){
  const excel=state.excel;$('excelBody').hidden=!excel;
  $('excelInfo').textContent=excel?`${excel.filename} · ${Object.keys(excel.sheets).length} worksheets`:'No Excel file loaded. Any worksheet whose first row holds column names can be used.';
  if(!excel)return;
  fill($('sheetSelect'),excel.order.map(name=>[name,excel.sheets[name].problem?`${name} (not usable)`:`${name} (${excel.sheets[name].row_count} ${excel.sheets[name].row_count===1?'row':'rows'})`]));
  renderSheet();
}
function renderSheet(){
  const sheet=state.excel.sheets[$('sheetSelect').value];$('sheetNotes').replaceChildren();$('sheetPreview').replaceChildren();if(!sheet)return;
  if(sheet.problem)$('sheetNotes').append(el('div',sheet.problem,'issue error'));
  for(const note of sheet.notes)$('sheetNotes').append(el('div',note,'issue'));
  if(sheet.problem)return;
  const table=el('table'),thead=el('thead'),head=el('tr'),body=el('tbody');head.append(el('th','Row'));
  for(const column of sheet.columns){const th=el('th',column);th.append(el('span',sheet.types[column],'path-label'));head.append(th);}
  for(const row of sheet.preview){const tr=el('tr');tr.append(el('td',row._row));for(const column of sheet.columns){const value=String(row[column]??'');const td=el('td',value.slice(0,80));td.title=value;tr.append(td);}body.append(tr);}
  thead.append(head);table.append(thead,body);$('sheetPreview').append(table);
  if(sheet.row_count>sheet.preview.length)$('sheetPreview').append(el('p',`Showing ${sheet.preview.length} of ${sheet.row_count} rows.`,'hint'));
}

/* Step 2: target AAS */
function renderAas(){
  const aas=state.aas;$('aasBody').hidden=!aas;
  if(!aas){$('aasInfo').textContent='Open an existing AASX package or create a new AAS.';return;}
  const shell=aas.shells.find(s=>s.id===aas.shell);
  $('aasInfo').textContent=(aas.opened?`Opened ${aas.filename} (${aas.format.toUpperCase()})`:`New AAS, exported as ${aas.filename}`)+(shell?` · ${shell.idShort} · ${shell.id}`:' · select a shell below');
  $('shellPicker').hidden=aas.shells.length<2;
  fill($('shellSelect'),aas.shells.map(s=>[s.id,`${s.idShort} (${s.id})`]),aas.shell?undefined:'Select…');if(aas.shell)$('shellSelect').value=aas.shell;
  $('aasProblems').replaceChildren();
  if(aas.existing_problem_count){const details=el('details'),summary=el('summary',`The opened AAS already has ${aas.existing_problem_count} metamodel problems. They are kept unchanged and do not block the import.`);details.append(summary);for(const p of aas.existing_problems)details.append(el('div',p,'leaf'));$('aasProblems').append(details);}
  renderTree($('aasTree'),nodes(),true);
  fill($('elementForm').elements.parent,nodes().filter(n=>n.container&&(n.modelType!=='SubmodelElementList'||n.listType==='SubmodelElementCollection')).map(n=>[key(n.target),describe(n.target)+(n.modelType==='SubmodelElementList'?' (list)':'')]));
  syncElementForm();
}
function renderTree(container,list,editable){
  container.replaceChildren();
  if(!list.length){container.append(el('div',editable?'No submodels yet. Add one from a template or create your own.':'Nothing to show.','empty-note'));return;}
  for(const n of list){
    if(n.more){const row=el('div',n.label,'node more');row.style.paddingLeft=`${10+n.depth*18}px`;container.append(row);continue;}
    const row=el('div',undefined,'node'+(n.changed?' changed':''));row.style.paddingLeft=`${10+n.depth*18}px`;
    if(n.semanticId)row.title=`semanticId: ${n.semanticId}`;
    row.append(el('span',n.label,'name'),el('span',SHORT[n.modelType]||n.modelType,'badge'+(n.container?' container':'')));
    if(n.valueType)row.append(el('span',n.valueType,'meta'));
    if(n.cardinality)row.append(el('span',n.cardinality,'meta'));
    row.append(el('span',n.value??'','val grow'));
    if(editable){
      if(n.container&&(n.modelType!=='SubmodelElementList'||n.listType==='SubmodelElementCollection')){const add=el('button','Add here');add.type='button';add.addEventListener('click',()=>{const f=$('elementForm');f.elements.parent.value=key(n.target);syncElementForm();(f.elements.idShort.closest('label').hidden?f.elements.modelType:f.elements.idShort).focus();});row.append(add);}
      const remove=el('button','Delete');remove.type='button';remove.addEventListener('click',()=>{if(confirm(`Delete ${describe(n.target)}?`))task('Deleting…',()=>apply(post('/api/mapper/aas/edit',{action:'delete',target:n.target}),'Deleted.'));});row.append(remove);
    }
    container.append(row);
  }
}
function syncElementForm(){
  const f=$('elementForm'),parent=nodes().find(n=>key(n.target)===f.elements.parent.value),inList=parent?.modelType==='SubmodelElementList';
  for(const o of f.elements.modelType.options)o.disabled=inList&&o.value!=='SubmodelElementCollection';
  if(inList)f.elements.modelType.value='SubmodelElementCollection';
  f.elements.idShort.closest('label').hidden=inList;
  f.querySelector('[data-for="Property"]').hidden=f.elements.modelType.value!=='Property';
}
function renderTemplates(){
  const options=[];
  for(const t of templates){
    if(t.error){options.push([`error:${t.file}`,`⚠ ${t.file}: ${t.error}`,true]);continue;}
    for(const s of t.submodels)options.push([JSON.stringify({file:t.file,submodel:s.id}),`${s.idShort} — ${t.file}${s.kind==='Template'?'':' (instance)'}`]);
  }
  fill($('templateForm').elements.template,options,options.some(o=>!o[2])?undefined:'No templates yet: add a template file');
}

/* Step 3: mapping */
function renderRules(){
  const box=$('rulesList');box.replaceChildren();
  if(!state.rules.length){box.append(el('p','No mappings yet. Add single values or table rows below.','hint'));return;}
  const table=el('table',undefined,'rules'),thead=el('thead'),head=el('tr'),body=el('tbody');
  for(const h of ['#','Type','Excel source','AAS target','Details',''])head.append(el('th',h));
  state.rules.forEach((rule,i)=>{
    const tr=el('tr');
    const source=rule.kind==='value'?`${rule.sheet} · ${rule.column} · `+(rule.row?`row ${rule.row}`:`where ${rule.match.column} = ${rule.match.value}`):`${rule.sheet} · all rows`;
    const details=rule.kind==='value'?(rule.language!=='en'?`language ${rule.language}`:''):
      [rule.key_column?`named by ${rule.key_column}`:'',rule.prototype!==null&&rule.prototype!==undefined?`structure like ${rule.prototype}`:'',rule.mode==='replace'?'replace all':''].filter(Boolean).concat(rule.columns.map(c=>`${c.column} → ${c.path.join('/')}`)).join(' · ');
    const td=el('td'),remove=el('button','Remove');remove.type='button';
    remove.addEventListener('click',()=>task('Saving mapping…',()=>apply(post('/api/mapper/rules',{rules:state.rules.filter((_,j)=>j!==i)}),'Mapping removed.')));
    td.append(remove);tr.append(el('td',i+1),el('td',rule.kind==='value'?'Single value':'Table rows'),el('td',source),el('td',describe(rule.target)),el('td',details),td);body.append(tr);
  });
  thead.append(head);table.append(thead,body);box.append(table);
}
function renderValueForm(){
  const f=$('valueForm');fill(f.elements.sheet,sheets().map(([n])=>[n,n]));
  const columns=(state.excel?.sheets[f.elements.sheet.value]?.columns||[]).map(c=>[c,c]);fill(f.elements.column,columns);fill(f.elements.matchColumn,columns);
  fill(f.elements.target,nodes().filter(n=>VALUE_KINDS.includes(n.modelType)||n.modelType==='Entity').map(n=>[key(n.target),`${describe(n.target)} (${n.modelType==='Entity'?'entity global asset ID':n.valueType||'multi-language'})`]));
  syncValueForm();
}
function syncValueForm(){
  const f=$('valueForm');for(const label of f.querySelectorAll('[data-mode]'))label.hidden=label.dataset.mode!==f.elements.rowMode.value;
  f.querySelector('[data-for="MultiLanguageProperty"]').hidden=nodes().find(n=>key(n.target)===f.elements.target.value)?.modelType!=='MultiLanguageProperty';
}
function containerNode(){return nodes().find(n=>key(n.target)===$('rowsForm').elements.target.value);}
function prototypes(parent){return parent?nodes().filter(n=>childOf(n,parent)&&n.target.path.length===parent.target.path.length+1&&ROW_ITEMS.includes(n.modelType)):[];}
function renderRowsForm(){
  const f=$('rowsForm');fill(f.elements.sheet,sheets().map(([n])=>[n,n]));
  fill(f.elements.target,nodes().filter(n=>n.container&&(n.modelType!=='SubmodelElementList'||ROW_ITEMS.includes(n.listType))).map(n=>[key(n.target),describe(n.target)+(n.modelType==='SubmodelElementList'?' (list)':n.modelType==='Entity'?' (entity)':'')]));
  syncRowsTarget(false);
}
function syncRowsTarget(targetChanged){
  const f=$('rowsForm'),parent=containerNode(),isList=parent?.modelType==='SubmodelElementList',choices=prototypes(parent);
  fill(f.elements.prototype,choices.map(n=>[JSON.stringify(n.target.path.at(-1)),`Like ${n.label}`]),'New properties only');
  if(targetChanged&&isList&&choices.length)f.elements.prototype.selectedIndex=1;
  for(const label of f.querySelectorAll('[data-for="named"]'))label.hidden=isList;
  f.querySelector('[data-for="list"]').hidden=!isList;
  fill(f.elements.keyColumn,(state.excel?.sheets[f.elements.sheet.value]?.columns||[]).map(c=>[c,c]),'Excel row number (Row_2, Row_3, …)');
  renderColumnMap();
}
function renderColumnMap(){
  const f=$('rowsForm'),parent=containerNode(),sheet=state.excel?.sheets[f.elements.sheet.value],body=$('columnMap');
  const previous=new Map([...body.querySelectorAll('tr')].map(tr=>[tr.dataset.column,{use:tr.querySelector('[data-role=use]').checked,target:tr.querySelector('[data-role=target]').value,idShort:tr.querySelector('[data-role=idShort]').value,valueType:tr.querySelector('[data-role=valueType]').value}]));
  body.replaceChildren();if(!sheet)return;
  let slots=[];
  if(parent&&f.elements.prototype.value){
    const segment=JSON.parse(f.elements.prototype.value),proto=prototypes(parent).find(n=>n.target.path.at(-1)===segment);
    if(proto)slots=nodes().filter(n=>childOf(n,proto)&&VALUE_KINDS.includes(n.modelType)).map(n=>n.target.path.slice(proto.target.path.length)).filter(p=>p.every(s=>typeof s==='string'));
    if(proto?.modelType==='Entity')slots.unshift([GLOBAL_ASSET]);
  }
  for(const column of sheet.columns){
    const tr=el('tr'),prev=previous.get(column);tr.dataset.column=column;
    const use=el('input');use.type='checkbox';use.dataset.role='use';use.checked=prev?prev.use:true;use.setAttribute('aria-label',`Use column ${column}`);
    const target=el('select');target.dataset.role='target';target.setAttribute('aria-label',`Target for ${column}`);fill(target,slots.map(p=>[JSON.stringify(p),p[0]===GLOBAL_ASSET?'Entity global asset ID':p.join(' / ')]),'New property');
    const match=slots.find(p=>p.at(-1).toLowerCase()===sanitize(column).toLowerCase());
    target.value=prev&&[...target.options].some(o=>o.value===prev.target)?prev.target:(match?JSON.stringify(match):'');
    const idShort=el('input');idShort.dataset.role='idShort';idShort.value=prev?.idShort||sanitize(column);idShort.setAttribute('aria-label',`New property name for ${column}`);
    const valueType=el('select');valueType.dataset.role='valueType';valueType.setAttribute('aria-label',`Value type for ${column}`);fill(valueType,VALUE_TYPES.map(t=>[t,t]));valueType.value=prev?.valueType||sheet.types[column]||'xs:string';
    const sync=()=>{target.disabled=!use.checked;idShort.disabled=valueType.disabled=!use.checked||!!target.value;};
    use.addEventListener('change',sync);target.addEventListener('change',sync);sync();
    for(const item of [use,el('span',column),target,idShort,valueType]){const td=el('td');td.append(item);tr.append(td);}
    body.append(tr);
  }
}

/* Step 4: result */
function renderReport(report,tree,imported){
  const box=$('resultIssues');box.replaceChildren();
  const counts=el('div',undefined,'counts');
  for(const [label,value] of [['Values written',report.values],['Elements created',report.created],['Rows updated',report.updated],['Empty optional elements left out',report.removed_optional??0],['Errors',report.errors.length],['Warnings',report.warnings.length]]){const metric=el('div',undefined,'metric');metric.append(el('span',label),el('b',value));counts.append(metric);}
  const list=el('div',undefined,'issue-list');
  for(const item of [...report.errors.map(x=>({...x,error:true})),...report.warnings]){
    const row=el('div',undefined,'issue'+(item.error?' error':''));const where=[item.rule?`Mapping ${item.rule}`:'',item.sheet,item.row?`row ${item.row}`:'',item.column].filter(Boolean).join(' · ');
    if(where)row.append(el('span',where,'issue-location'));row.append(el('span',item.message));list.append(row);
  }
  box.append(counts,list);
  $('resultInfo').textContent=report.errors.length?'Fix the errors below; nothing is written while errors remain.':imported?'Imported. Written values are highlighted.':'No errors. Values that will be written are highlighted; Import writes them into a new AASX.';
  $('resultTree').hidden=!tree;if(tree)renderTree($('resultTree'),tree,false);
}

function render(){renderExcel();renderAas();renderTemplates();renderRules();renderValueForm();renderRowsForm();buttons();}

$('sheetSelect').addEventListener('change',renderSheet);
$('excelFile').addEventListener('change',()=>task('Reading Excel…',async()=>{const file=$('excelFile').files[0];$('excelFile').value='';if(file)await apply(upload('/api/mapper/excel',file),`Loaded ${file.name}.`);}));
$('aasFile').addEventListener('change',()=>task('Opening AAS…',async()=>{const file=$('aasFile').files[0];$('aasFile').value='';if(file&&(!state.aas||confirm('Replace the current target AAS? Structure changes that were not imported are discarded.')))await apply(upload('/api/mapper/aas/open',file),`Opened ${file.name}. The original file is not changed.`);}));
$('newAasToggle').addEventListener('click',()=>{const form=$('newAasForm');form.hidden=!form.hidden;if(!form.hidden)form.elements.idShort.focus();});
$('newAasForm').addEventListener('submit',e=>{e.preventDefault();const f=e.target.elements;if(state.aas&&!confirm('Replace the current target AAS? Structure changes that were not imported are discarded.'))return;
  task('Creating AAS…',async()=>{await apply(post('/api/mapper/aas/new',{idShort:f.idShort.value.trim(),globalAssetId:f.globalAssetId.value.trim()}),'New AAS created. Add submodels from templates or create your own.');$('newAasForm').hidden=true;});});
$('shellSelect').addEventListener('change',e=>{if(e.target.value)task('Selecting shell…',()=>apply(post('/api/mapper/aas/shell',{shell:e.target.value})));});
$('templateForm').addEventListener('submit',e=>{e.preventDefault();const f=e.target.elements;if(!f.template.value)return message('Add a template file first.',true);const t=JSON.parse(f.template.value);
  task('Adding submodel…',()=>apply(post('/api/mapper/aas/edit',{action:'template',file:t.file,submodel:t.submodel,keepValues:f.keepValues.checked}),'Template submodel added. Add your own elements where the template is not enough.'));});
$('templateFile').addEventListener('change',()=>task('Adding template…',async()=>{const file=$('templateFile').files[0];$('templateFile').value='';if(!file)return;const data=await upload('/api/mapper/templates',file,false);templates=data.templates;renderTemplates();message(`Template file ${data.added} added to the library.`);}));
$('submodelForm').addEventListener('submit',e=>{e.preventDefault();const form=e.target,f=form.elements;task('Adding submodel…',async()=>{await apply(post('/api/mapper/aas/edit',{action:'submodel',idShort:f.idShort.value.trim(),semanticId:f.semanticId.value.trim()}),'Submodel added.');form.reset();});});
$('elementForm').addEventListener('change',syncElementForm);
$('elementForm').addEventListener('submit',e=>{e.preventDefault();const f=e.target.elements;if(!f.parent.value)return message('Add a submodel first.',true);
  task('Adding element…',async()=>{await apply(post('/api/mapper/aas/edit',{action:'element',parent:JSON.parse(f.parent.value),modelType:f.modelType.value,idShort:f.idShort.value.trim(),valueType:f.valueType.value,semanticId:f.semanticId.value.trim()}),'Element added.');f.idShort.value='';f.semanticId.value='';});});
$('valueForm').addEventListener('change',e=>{if(e.target.name==='sheet')renderValueForm();else syncValueForm();});
$('valueForm').addEventListener('submit',e=>{e.preventDefault();const f=e.target.elements;if(!f.target.value)return message('Select a target property. Add one in step 2 if it does not exist yet.',true);
  const rule={kind:'value',sheet:f.sheet.value,column:f.column.value,target:JSON.parse(f.target.value),language:f.language.value.trim()||'en'};
  if(f.rowMode.value==='row')rule.row=Number(f.row.value);else rule.match={column:f.matchColumn.value,value:f.matchValue.value};
  task('Saving mapping…',()=>apply(post('/api/mapper/rules',{rules:[...state.rules,rule]}),'Mapping added.'));});
$('rowsForm').addEventListener('change',e=>{if(e.target.dataset.role)return;if(e.target.name==='sheet')syncRowsTarget(false);else if(e.target.name==='target')syncRowsTarget(true);else if(e.target.name==='prototype')renderColumnMap();});
$('rowsForm').addEventListener('submit',e=>{e.preventDefault();const f=e.target.elements;if(!f.target.value)return message('Select a target container. Add a submodel or collection in step 2 first.',true);
  const isList=containerNode()?.modelType==='SubmodelElementList';
  const columns=[...$('columnMap').querySelectorAll('tr')].filter(tr=>tr.querySelector('[data-role=use]').checked).map(tr=>{const target=tr.querySelector('[data-role=target]').value;
    return target?{column:tr.dataset.column,path:JSON.parse(target)}:{column:tr.dataset.column,path:[tr.querySelector('[data-role=idShort]').value.trim()],valueType:tr.querySelector('[data-role=valueType]').value};});
  if(!columns.length)return message('Select at least one column.',true);
  const rule={kind:'rows',sheet:f.sheet.value,target:JSON.parse(f.target.value),prototype:f.prototype.value?JSON.parse(f.prototype.value):null,key_column:isList?'':f.keyColumn.value,mode:isList?'replace':f.mode.value,columns};
  task('Saving mapping…',async()=>{await apply(post('/api/mapper/rules',{rules:[...state.rules,rule]}),'Mapping added.');$('columnMap').replaceChildren();renderColumnMap();});});
$('rulesFile').addEventListener('change',()=>task('Loading mapping…',async()=>{const file=$('rulesFile').files[0];$('rulesFile').value='';if(!file)return;let data;
  try{data=JSON.parse(await file.text());}catch(e){throw new Error('The mapping file is not valid JSON.');}
  await apply(post('/api/mapper/rules',{rules:Array.isArray(data)?data:data.rules}),'Mapping loaded.');}));
$('checkButton').addEventListener('click',()=>task('Checking…',async()=>{const data=await post('/api/mapper/check',{dropEmptyOptional:$('dropOptional').checked});state=data;render();const report=data.result.report;renderReport(report,data.result.tree,false);
  message(report.errors.length?`Check found ${report.errors.length} ${report.errors.length===1?'error':'errors'}.`:'Check passed. Review the highlighted values, then import.',!!report.errors.length);}));
$('importButton').addEventListener('click',()=>task('Importing…',async()=>{const data=await post('/api/mapper/import',{dropEmptyOptional:$('dropOptional').checked});state=data;render();renderReport(data.result.report,data.result.tree,true);
  const link=el('a',`Download ${data.result.filename}`);link.href=data.result.download_url;link.className='button';link.id='downloadResult';message('Import finished.',false,link);}));
fill($('elementForm').elements.valueType,VALUE_TYPES.map(t=>[t,t]));
task('Loading…',async()=>{state=await request('/api/mapper/state');templates=(await request('/api/mapper/templates')).templates;render();});
