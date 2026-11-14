// Repair helpers: background links remain separate from the fixed SST goal area.
function sstRepairSourceConnections(){
 const ss=sstBook_(),lock=LockService.getDocumentLock();lock.waitLock(30000);
 try{
  const props=PropertiesService.getDocumentProperties(),registry=sstRegistry_(),g=ss.getSheetByName('General');
  const before=sstRepairPreservedState_(ss),people=sstPeople_(ss),patches=[];
  SST_ROUNDS.forEach(t=>{const s=ss.getSheetByName(t);people.forEach(p=>{const row=sstFindRow_(s,p.name);['B','BK'].forEach(c=>{const cell=s.getRange(c+row),f=cell.getFormula();if(!f||!f.includes('XLOOKUP')||!f.includes('General!'))throw Error('Unexpected background cell at '+t+'!'+c+row+'. Nothing was cleared.');patches.push({sheet:t,row,column:c,old:f,next:sstSafeBackgroundFormula_(row,c,c==='B'?'AR':'BO')});});});});
  const ghost=registry.aleyza&&registry.aleyza.records&&registry.aleyza.records[1],partial=ss.getSheetByName('Progress - Aleyza').getRange('A61:H62');
  const ghostMissing=ghost&&!ss.getNamedRanges().some(n=>n.getName()===ghost.rangeName);
  if(ghostMissing){if(ghost.autoVersion||ghost.state||(ghost.blocks||[]).length||partial.getDisplayValues().flat().filter(Boolean).join('|')!=='SST 1 — meeting record|Meeting date')throw Error('Incomplete draft requires manual review; no records were removed.');}
  const backup=Utilities.base64Encode(Utilities.gzip(Utilities.newBlob(JSON.stringify({at:new Date().toISOString(),registry,patches,partial:partial.getValues()}))).getBytes());
  for(let i=0;i<backup.length;i+=7000)props.setProperty('SST_SOURCE_REPAIR_BACKUP_'+(i/7000),backup.slice(i,i+7000));
  patches.forEach(p=>ss.getSheetByName(p.sheet).getRange(p.column+p.row).setFormula(p.next));SpreadsheetApp.flush();
  if(ghostMissing){props.setProperty('SST_QUARANTINED_INCOMPLETE_DRAFT',JSON.stringify({student:'Aleyza',round:1,record:ghost,reason:'Background formula overflow after a General column insertion; no goals or progress saved in this draft.'}));delete registry.aleyza.records[1];sstSaveRegistry_(registry);partial.clearContent();}
  const blank=g.getRange('BL1:BP1000');if(blank.getDisplayValues().flat().every(v=>!v.trim())&&!blank.getFormulas().flat().some(Boolean))g.hideColumns(64,5);
  ss.getSheetByName('Progress - Aleyza').hideRows(60,8);
  sstReconcile_(ss);SpreadsheetApp.flush();
  const after=sstRepairPreservedState_(ss);if(JSON.stringify(before)!==JSON.stringify(after))throw Error('Preservation check failed. Recovery snapshot was kept; inspect before further changes.');
  const checks=sstConnectionAudit_(ss);if(checks.errors.length)throw Error(checks.errors.join(' | '));
  props.setProperty('SST_SOURCE_CONNECTIONS_VERSION','1');console.log('REPAIR PASSED '+JSON.stringify({...checks,preservedStudentData:true,preservedGoalsAndProgress:true,incompleteDraftQuarantined:!!ghostMissing,blankColumnsHidden:true}));
  ss.toast('Connections repaired and checked. Student information, goals and progress were preserved.','SST connections',10);
 }finally{lock.releaseLock();}
}
function sstVerifyConnections(){
 const ss=sstBook_(),before=sstRepairPreservedState_(ss);sstSafeStartup_();SpreadsheetApp.flush();const result=sstConnectionAudit_(ss);result.preserved=JSON.stringify(before)===JSON.stringify(sstRepairPreservedState_(ss));if(result.errors.length||!result.preserved)throw Error(JSON.stringify(result));console.log('CONNECTION CHECK PASSED '+JSON.stringify(result));ss.toast('Automatic connections checked successfully. Saved goals and progress were preserved.','SST connections',9);
}
function sstSafeStartup_(waitMs){const lock=LockService.getDocumentLock();if(!lock.tryLock(waitMs||1000))return false;try{sstReconcile_(sstBook_());return true;}finally{lock.releaseLock();}}
function sstSafeBackgroundFormula_(row,first,last){
 const ss=sstBook_(),general=ss.getSheetByName('General'),source=ss.getSheetByName(SST_ROUNDS[0]),heads=general.getRange(2,1,1,general.getLastColumn()).getDisplayValues()[0],target=source.getRange(2,1,1,source.getLastColumn()).getDisplayValues()[0];
 function norm(v){let k=String(v).replace(/\s*\(\d{4}-\d{4}\)/g,'').trim().replace(/\s+/g,' ').toLowerCase();return({'reading intervention teacher comments':'comments from reading intervention teacher','eld teacher comments':'comments from eld support teacher','eld support frequency':'how often is the support provided?'}[k]||k);}
 function colNumber(t){return t.split('').reduce((n,c)=>n*26+c.charCodeAt(0)-64,0);}
 const a=colNumber(first),b=colNumber(last),ranges=[];
 for(let c=a;c<=b;c++){const key=norm(target[c-1]);let hits=heads.map((h,i)=>norm(h)===key?i+1:0).filter(Boolean),occurrence=a>=63?1:target.slice(0,c).filter(h=>norm(h)===key).length;if(!key||!hits[occurrence-1])throw Error('Cannot match shared field '+target[c-1]+' to General.');const col=sstColumn_(hits[occurrence-1]);ranges.push('General!$'+col+'$3:$'+col+'$1000');}
 const data=ranges.length===1?ranges[0]:'HSTACK('+ranges.join(',')+')';return '=IF(TRIM($A'+row+')="","",ARRAYFORMULA(IFNA(LET(record,XLOOKUP(TRIM($A'+row+'),ARRAYFORMULA(TRIM(General!$A$3:$A$1000)),'+data+'),IF(record="","",record)),"")))';
}
function sstRepairPreservedState_(ss){
 const reg=sstRegistry_();return{general:ss.getSheetByName('General').getRange('A1:BT7').getValues(),students:Object.entries(reg).map(([key,e])=>({key,id:e.studentId,sheetId:e.sheetId,records:Object.entries(e.records||{}).filter(([,r])=>r.autoVersion===2).map(([n,r])=>({n,state:r.state,control:r.controlName,range:r.rangeName,meeting:ss.getSheetById(e.sheetId).getRange(7,1,1,8).getValues(),blocks:(r.blocks||[]).map(b=>({id:b.id,range:b.rangeName,goal:b.goal,plan:b.plan,version:b.version,values:sstNamed_(ss,b.rangeName).getValues(),notes:sstNamed_(ss,b.rangeName).getNotes()}))}))}))};
}
function sstConnectionAudit_(ss){
 const errors=[],registry=sstRegistry_(),people=sstPeople_(ss);let records=0,blocks=0,backgroundCells=0,sourceGoals=0;
 SST_ROUNDS.forEach(t=>{const s=ss.getSheetByName(t);people.forEach(p=>{const r=sstFindRow_(s,p.name);const row=s.getRange(r,2,1,66).getDisplayValues()[0];if(row.some(v=>/^#(REF!|N\/A|VALUE!|ERROR!|NAME\?|NUM!|DIV\/0!)/.test(v)))errors.push(t+' row '+r+' has a formula error.');['B','BK'].forEach(c=>{if(!s.getRange(c+r).getFormula().includes('HSTACK('))errors.push(t+' '+c+r+' still uses an expanding background range.');backgroundCells++;});sourceGoals+=sstGoals_(ss,p.name,SST_ROUNDS.indexOf(t)+1).length;});});
 Object.entries(registry).forEach(([key,e])=>Object.entries(e.records||{}).forEach(([n,r])=>{if(r.autoVersion!==2){errors.push(key+' SST'+n+' is incomplete.');return;}records++;if(!ss.getRangeByName(r.rangeName)||!ss.getRangeByName(r.controlName))errors.push(key+' SST'+n+' marker missing.');(r.blocks||[]).forEach(b=>{blocks++;if(!ss.getRangeByName(b.rangeName))errors.push(key+' goal marker missing.');});}));
 return{students:people.length,records,goalBlocks:blocks,sourceGoals,backgroundFormulaCells:backgroundCells,errors};
}
