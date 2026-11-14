// Writing feedback connections read the rubric spreadsheets configured in Rubric Connections.
const SST_COPY_ID = '1Dr938_RpMMEbaVrpI2J5VMkoOXwxdoShzfl9XLGqdAI';
const SST_INDEX = 'Student Index';
const SST_ROUNDS = ['SST 1', 'SST2', 'SST3'];
const SST_COLORS = {ink:'#223f53', green:'#246962', pale:'#e9f3ef', input:'#fff8e6', blue:'#e9eff8'};

function sstBook_(){const s=SpreadsheetApp.getActiveSpreadsheet();if(!s||s.getId()!==SST_COPY_ID)throw Error('This script is limited to the newest authorized spreadsheet copy.');return s;}
function sstKey_(name){return String(name).trim().toLowerCase();}
function sstRegistry_(){return JSON.parse(PropertiesService.getDocumentProperties().getProperty('SST_PROGRESS_REGISTRY')||'{}');}
function sstSaveRegistry_(r){PropertiesService.getDocumentProperties().setProperty('SST_PROGRESS_REGISTRY',JSON.stringify(r));}
function sstPeople_(ss){
 const g=ss.getSheetByName('General');if(!g)throw Error('General is missing.');
 const rows=g.getRange(3,1,g.getMaxRows()-2,71).getValues();const seen={};
 return rows.map((v,i)=>({name:String(v[0]).trim(),row:i+3,values:v})).filter(p=>p.name).map(p=>{let k=sstKey_(p.name);if(seen[k])throw Error('Duplicate student name in General: '+p.name+'. Give each student a unique name.');seen[k]=true;return p;});
}
function sstRule_(options){return sstSharedRule_(options);}
function sstFormula_(r,first,last){return sstSafeBackgroundFormula_(r,first,last);}
function sstGoals_(ss,name,n){
 const cached=SST_SOURCE_READ_CACHE_&&SST_SOURCE_READ_CACHE_[SST_ROUNDS[n-1]+':'+sstKey_(name)];let values,heads;
 if(cached){values=cached.values;heads=cached.heads;}else{const s=ss.getSheetByName(SST_ROUNDS[n-1]),row=sstFindRow_(s,name);values=s.getRange(row,45,1,18).getDisplayValues()[0];heads=s.getRange(2,45,1,18).getDisplayValues()[0];}
 const goals=[];for(let i=0;i<18;i+=2){if(String(values[i]).trim())goals.push({area:heads[i].replace(/\s*goal\s*$/i,''),goal:values[i],plan:values[i+1]||'',sourceColumn:i+45});}return goals;
}
function sstLink_(sheet,row,label){return '=HYPERLINK("#gid='+sheet.getSheetId()+'&range=A'+row+'","'+String(label).replace(/"/g,'""')+'")';}
function sstLegacyIndex_(ss,people){
 let s=ss.getSheetByName(SST_INDEX),fresh=!s;if(fresh)s=ss.insertSheet(SST_INDEX);
 let name=fresh?(people[0]?people[0].name:''):s.getRange('B3').getDisplayValue(),round=fresh?'SST 3':s.getRange('B4').getDisplayValue();
 s.getRange('A1:H2').breakApart().merge().setValue('SST STUDENT INDEX').setBackground(SST_COLORS.ink).setFontColor('white').setFontSize(22).setFontWeight('bold');
 s.getRange('A3').setValue('Choose student');s.getRange('A4').setValue('Choose meeting');
 s.getRange('B3').clearDataValidations().setValue(people.some(p=>p.name===name)?name:(people[0]?people[0].name:''));if(people.length)s.getRange('B3').setDataValidation(sstRule_(people.map(p=>p.name)));s.getRange('B3').setBackground(SST_COLORS.input);
 s.getRange('B4').setDataValidation(sstRule_(['SST 1','SST 2','SST 3'])).setValue(['SST 1','SST 2','SST 3'].includes(round)?round:'SST 3').setBackground(SST_COLORS.input);
 s.getRange('C3:H4').breakApart().merge().setValue('Choose the student and meeting, then use SST Case Manager → Create or open selected meeting tracker. Existing records open safely without replacing progress.').setWrap(true).setBackground(SST_COLORS.pale);
 s.getRange('A6:H6').breakApart().merge().setValue('Include Writing evidence in the NEW meeting record? Check only the pieces you want.').setFontWeight('bold');
 s.getRange('B7:G7').setValues([['Spanish narrative','Spanish opinion','Spanish informational','English narrative','English opinion','English informational']]).setWrap(true).setBackground(SST_COLORS.blue);
 if(fresh)s.getRange('B8:G8').insertCheckboxes();
 s.getRange('A10:H10').breakApart().merge().setValue('Click a link below. For new goals, use the matching SST tab. For dated results, use the student’s Progress tab.').setWrap(true);
 s.getRange('A12:H12').setValues([['Student','Previous-year SST count','Suggested next SST — confirm','Student tracker','SST 1 record','SST 2 record','SST 3 record','Notes']]).setBackground(SST_COLORS.green).setFontColor('white').setFontWeight('bold').setWrap(true);
 const registry=sstRegistry_();let rows=people.map(p=>{
  const previous=Number(p.values[3])||0,key=sstKey_(p.name),entry=registry[key],tracker=entry&&entry.sheetId&&ss.getSheetById(entry.sheetId);
  const links=[1,2,3].map(n=>{const record=entry&&entry.records&&entry.records[n],r=record&&ss.getRangeByName(record.rangeName);return r?sstLink_(r.getSheet(),r.getRow(),'Open SST '+n):'';});
  return[p.name,previous,Math.min(3,previous+1),tracker?sstLink_(tracker,1,'Open progress'):'',...links,previous?'Confirm the suggested meeting number before entering progress.':'First SST this year, unless other records indicate otherwise.'];
 });
 if(rows.length)s.getRange(13,1,rows.length,8).setValues(rows).setWrap(true).setVerticalAlignment('top');else s.getRange('A13:H13').setValues([['No active students','','','','','','','Use SST Progress > Tracker controls to restore a student.']]);
 const help=[['1. General','Edit background information, assessments, and teacher comments here.'],['2. SST 1 / SST2 / SST3','Enter goals and plans only in AS:BJ for the correct student and meeting. Background fields link to General.'],['3. Create the tracker','Choose student and meeting above. Use the SST Progress menu. Only populated goals appear.'],['4. Record progress','Use the dropdowns and result cells in the student’s Progress tab. Type your own words in the custom cells.'],['5. Next meeting','Use Copy selected goals to the next SST draft. Review the new goals, then create the new meeting record.'],['Saved records','Goals, plans, and selected Writing comments are copied into a stable meeting record. Later source edits do not rewrite it.'],['Adding a student','Add the student to General, then use Refresh General links and student index. Student names must be unique.']];
 let h=people.length+16;help.forEach((r,i)=>{s.getRange(h+i,1).setValue(r[0]).setFontWeight('bold');s.getRange(h+i,2,1,7).breakApart().merge().setValue(r[1]).setWrap(true);s.setRowHeight(h+i,44)});
 s.getRange('A3:H'+(h+help.length)).setFontFamily('Arial').setFontSize(11).setVerticalAlignment('top');
 [145,120,150,135,110,110,110,260].forEach((w,i)=>s.setColumnWidth(i+1,w));s.setRowHeight(7,48);s.setRowHeight(12,58);for(let r=13;r<13+people.length;r++)s.setRowHeight(r,65);s.setFrozenRows(2);s.setHiddenGridlines(true);s.setTabColor(SST_COLORS.green);ss.setActiveSheet(s);ss.moveActiveSheet(1);
}
function sstLegacyRecord_(ss,p,n,goals,activate){
 let registry=sstRegistry_(),key=sstKey_(p.name),entry=registry[key]||{sheetId:null,records:{}},existing=entry.records[n];
 if(existing){const r=ss.getRangeByName(existing.rangeName);if(!r)throw Error('Saved record marker is missing for '+p.name+'. No record was overwritten.');if(activate){ss.setActiveSheet(r.getSheet());r.activate();}return r;}
 let s=entry.sheetId&&ss.getSheetById(entry.sheetId);if(!s){const title=sstProgressTitle_(p.name);const conflict=ss.getSheetByName(title);if(conflict)throw Error('A sheet called '+title+' already exists but is not registered. It was not overwritten.');s=ss.insertSheet(title);entry.sheetId=s.getSheetId();s.getRange('A1:H2').merge().setValue(p.name+' — SST PROGRESS').setBackground(SST_COLORS.ink).setFontColor('white').setFontWeight('bold').setFontSize(20);s.getRange('A3:H3').merge().setValue('Use dropdowns where helpful. Enter your own wording in the custom fields. Each meeting keeps its own goals and progress.').setWrap(true).setBackground(SST_COLORS.pale);s.setRowHeight(3,45);s.setFrozenRows(3);s.setHiddenGridlines(true);[100,110,145,140,135,160,170,250].forEach((w,i)=>s.setColumnWidth(i+1,w));s.setTabColor(SST_COLORS.green);}
 let start=Math.max(5,s.getLastRow()+3),r=start;
 function bar(text){s.getRange(r,1,1,8).merge().setValue(text).setBackground(SST_COLORS.green).setFontColor('white').setFontWeight('bold').setFontSize(14);s.setRowHeight(r,35);r++;}
 function field(label,value,color){s.getRange(r,1).setValue(label).setFontWeight('bold');s.getRange(r,2,1,7).merge().setValue(value).setBackground(color||'white').setWrap(true);s.setRowHeight(r,Math.max(42,Math.ceil(String(value).length/135)*18+14));r++;}
 bar('SST '+n+' — meeting record');
 s.getRange(r,1,1,8).setValues([['Meeting date','','Review date','','Meeting notes','','','']]);s.getRange(r,6,1,3).merge().setBackground(SST_COLORS.input);s.getRange(r,2).setNumberFormat('mm/dd/yy').setBackground(SST_COLORS.input);s.getRange(r,4).setNumberFormat('mm/dd/yy').setBackground(SST_COLORS.input);s.setRowHeight(r,48);r++;
 s.getRange(r,1,1,8).merge().setValue('Saved goals and plans below. Update the SST source for a future meeting; this record will stay unchanged.').setWrap(true).setFontColor('#61717d');s.setRowHeight(r,34);r+=2;
 const status=['Not yet met','Making progress','Met this check','Goal met — criteria checked','Not assessed','Absent','Other / my words'];
 const method=['Word reading check','Spelling dictation','Writing sample / rubric','Math quiz','Work sample','Reading passage','Observation','Other / my words'];
 const support=['Independent','Verbal prompt','Visual support','Model / example','Guided practice','One-to-one support','Other / my words'];
 const observation=['More accurate','More fluent','More independent','Needs reteaching','Inconsistent performance','No clear change yet','Other / my words'];
 goals.forEach((g,i)=>{bar('Goal '+(i+1)+' — '+g.area);field('Student goal',g.goal);field('Intervention plan',g.plan||'Plan not entered yet.');
  s.getRange(r,1).setValue('Assessment').setFontWeight('bold');s.getRange(r,2,1,2).merge().setDataValidation(sstRule_(method)).setBackground(SST_COLORS.input);s.getRange(r,4).setValue('Other method');s.getRange(r,5,1,4).merge().setBackground(SST_COLORS.input);s.setRowHeight(r,42);r++;
  s.getRange(r,1).setValue('Frequency').setFontWeight('bold');s.getRange(r,2,1,2).merge().setDataValidation(sstRule_(['Daily','Twice weekly','Weekly','Every two weeks','As scheduled','Other / my words'])).setBackground(SST_COLORS.input);s.getRange(r,4).setValue('Other frequency');s.getRange(r,5,1,4).merge().setBackground(SST_COLORS.input);s.setRowHeight(r,42);r++;
  s.getRange(r,1).setValue('Baseline date').setFontWeight('bold');s.getRange(r,2).setNumberFormat('mm/dd/yy').setBackground(SST_COLORS.input);s.getRange(r,3).setValue('Baseline result');s.getRange(r,4,1,5).merge().setBackground(SST_COLORS.input);s.setRowHeight(r,42);r++;
  s.getRange(r,1,1,8).setValues([['Date','Result / score','Progress status','Assessment type','Support provided','Observation','Custom status / other','My notes / evidence']]).setBackground(SST_COLORS.blue).setFontWeight('bold').setWrap(true);s.setRowHeight(r,45);r++;
  const count=15;s.getRange(r,1,count,8).setBackground('#fffdf5').setBorder(true,true,true,true,true,true,'#d5dfe4',SpreadsheetApp.BorderStyle.SOLID).setWrap(true).setVerticalAlignment('top');s.getRange(r,1,count,1).setNumberFormat('mm/dd/yy');s.getRange(r,2,count,1).setNumberFormat('@');[[3,status],[4,method],[5,support],[6,observation]].forEach(pair=>s.getRange(r,pair[0],count,1).setDataValidation(sstRule_(pair[1])));s.setRowHeights(r,count,60);r+=count;
  field('Goal review / next step','',SST_COLORS.input);r++;
 });
 const chosen=[];
 const writing=p.values.slice(23,29);writing.forEach((v,i)=>{if(chosen[i]===true&&String(v).trim()){bar('Writing evidence — '+['Spanish narrative','Spanish opinion','Spanish informational','English narrative','English opinion','English informational'][i]);field('Assessment date','',SST_COLORS.input);field('Saved comments',v);}});
 s.getRange(start,1,r-start,8).setFontFamily('Arial').setFontSize(11).setVerticalAlignment('top').setWrap(true);
 let rangeName='SSTREC_'+Utilities.getUuid().replace(/-/g,'');ss.setNamedRange(rangeName,s.getRange(start,1,r-start,8));entry.records[n]={rangeName:rangeName,source:SST_ROUNDS[n-1],goals:goals.length};registry[key]=entry;sstSaveRegistry_(registry);
 if(activate){ss.setActiveSheet(s);s.getRange(start,1).activate();}return s.getRange(start,1,r-start,8);
}
function sstSelection_(ss){const s=ss.getSheetByName(SST_INDEX);if(!s)throw Error('Build the SST workspace first.');const name=s.getRange('B3').getDisplayValue().trim(),n=Number(s.getRange('B4').getDisplayValue().replace(/\D/g,''));if(!name||![1,2,3].includes(n))throw Error('Choose a student and SST round in Student Index.');const p=sstPeople_(ss).find(p=>sstKey_(p.name)===sstKey_(name));if(!p)throw Error('Student is missing from General.');return{p,n};}
function sstDiagnoseWarnings(){
 const ss=sstBook_();
 const named=ss.getNamedRanges().map(n=>{try{let r=n.getRange();return{name:n.getName(),sheet:r.getSheet().getName(),a1:r.getA1Notation()};}catch(e){return{name:n.getName(),error:e.message};}});
 console.log('NAMED_RANGES '+JSON.stringify(named));
 const registry=sstRegistry_();console.log('REGISTRY '+JSON.stringify(Object.entries(registry).map(([key,e])=>({key,sheetId:e.sheetId,studentId:e.studentId,records:Object.entries(e.records||{}).map(([n,r])=>({n,rangeName:r.rangeName,controlName:r.controlName,autoVersion:r.autoVersion,state:r.state,blocks:(r.blocks||[]).map(b=>({id:b.id,rangeName:b.rangeName,sourceColumn:b.sourceColumn,area:b.area,active:b.active}))}))}))));
 ['General',...SST_ROUNDS].forEach(t=>{const s=ss.getSheetByName(t);console.log('SOURCE '+JSON.stringify({sheet:t,id:s.getSheetId(),cols:s.getMaxColumns(),headers:s.getRange(2,1,1,s.getLastColumn()).getDisplayValues()[0],names:s.getRange(3,1,Math.max(1,s.getLastRow()-2),1).getDisplayValues().flat().filter(Boolean)}));});
 ss.getSheets().filter(s=>s.getName().startsWith('Progress - ')).forEach(s=>{let a=s.getRange(1,1,s.getLastRow(),8).getDisplayValues();console.log('PROGRESS '+JSON.stringify({sheet:s.getName(),id:s.getSheetId(),lastRow:s.getLastRow(),labels:a.map((r,i)=>({row:i+1,label:r[0],value:r[1]})).filter(r=>/SST \d|goal|Start this|Revise|Date|Goal review/i.test(r.label)),merges:s.getRange(1,1,s.getLastRow(),8).getMergedRanges().map(r=>r.getA1Notation())}));});
}

function sstOpenIndex(){const ss=sstBook_(),s=ss.getSheetByName(SST_INDEX);if(!s){SpreadsheetApp.getUi().alert('Use Build / finish SST workspace first.');return;}ss.setActiveSheet(s);s.getRange('B3').activate();}

function sstAddWriting(){
 const ss=sstBook_();if(ss.getActiveSheet().getName()!==SST_INDEX){ss.toast('Go to Go to tab → Student Index. Choose the student, SST number and writing pieces there first.','Writing evidence — Student Index only',-1);return;}const lock=LockService.getDocumentLock();lock.waitLock(10000);
 try{const x=sstSelection_(ss),registry=sstRegistry_(),entry=registry[sstKey_(x.p.name)],record=entry&&entry.records[x.n];if(!record){SpreadsheetApp.getUi().alert('Create the selected meeting tracker first.');return;}
 const range=ss.getRangeByName(record.rangeName),sheet=range.getSheet(),existing=range.getDisplayValues().flat().join('\n'),selected=ss.getSheetByName(SST_INDEX).getRange('B8:G8').getValues()[0];
 const titles=['Spanish narrative','Spanish opinion','Spanish informational','English narrative','English opinion','English informational'];
 const writing=x.p.values.slice(23,29).map((value,i)=>({value:String(value),i})).filter(v=>selected[v.i]===true&&v.value.trim()&&!existing.includes(v.value));
 if(!writing.length){ss.toast('Check the Writing pieces in Student Index. Empty or already saved comments are skipped.','SST Progress',8);return;}
 const count=writing.length*4,after=range.getLastRow();sheet.insertRowsAfter(after,count);let r=after+1;
 writing.forEach(v=>{sheet.getRange(r,1,1,8).merge().setValue('Writing evidence — '+titles[v.i]).setBackground(SST_COLORS.green).setFontColor('white').setFontWeight('bold');sheet.setRowHeight(r,35);r++;
 sheet.getRange(r,1).setValue('Assessment date');sheet.getRange(r,2,1,7).merge().setBackground(SST_COLORS.input).setNumberFormat('mm/dd/yy');sheet.setRowHeight(r,40);r++;
 sheet.getRange(r,1).setValue('Saved comments');sheet.getRange(r,2,1,7).merge().setValue(v.value).setWrap(true);sheet.setRowHeight(r,Math.max(60,Math.ceil(v.value.length/130)*18+15));r+=2;});
 sheet.getRange(after+1,1,count,8).setFontFamily('Arial').setFontSize(11).setWrap(true).setVerticalAlignment('top');ss.setNamedRange(record.rangeName,sheet.getRange(range.getRow(),1,range.getNumRows()+count,8));
 sstWarningProgress_(ss,sheet);sstIndex_(ss,sstPeople_(ss));ss.setActiveSheet(sheet);sheet.getRange(after+1,1).activate();ss.toast('Selected Writing comments saved with this meeting. Goals and progress are unchanged.','SST Progress',8);
 }finally{lock.releaseLock();}
}

// Automatic drafts and explicit meeting approval. This project remains limited to SST_COPY_ID.
var SST_ROW_CACHE_={};
var SST_SOURCE_READ_CACHE_=null,SST_INDEX_PEOPLE_CACHE_=null;
function onOpen(){
 try{sstBuildMenus_();}catch(e){console.log('Case Manager menus load when the spreadsheet is opened.');}
 if(PropertiesService.getDocumentProperties().getProperty('SST_AUTO_VERSION')==='2'){
  try{sstSafeStartup_();}catch(e){sstFailure_(e);}
 }
}
function sstFindRow_(sheet,name){
 const cacheKey=sheet.getSheetId()+':'+sstKey_(name);if(SST_ROW_CACHE_[cacheKey])return SST_ROW_CACHE_[cacheKey];
 const names=sheet.getRange(3,1,sheet.getMaxRows()-2,1).getDisplayValues().flat();
 const matches=[];names.forEach((v,i)=>{if(sstKey_(v)===sstKey_(name))matches.push(i+3);});
 if(matches.length>1)throw Error('Duplicate student name in '+sheet.getName()+': '+name);
 if(matches.length){SST_ROW_CACHE_[cacheKey]=matches[0];return matches[0];}
 let row=names.findIndex(v=>!String(v).trim())+3;if(row<3)row=names.length+3;
 if(row>sheet.getMaxRows())sheet.insertRowsAfter(sheet.getMaxRows(),20);
 if(sheet.getRange(row,45,1,18).getDisplayValues()[0].some(v=>v.trim()))throw Error('A blank name row contains goals in '+sheet.getName()+'. Review it before adding a student.');
 sheet.getRange(row,1).setValue(name);SST_ROW_CACHE_[cacheKey]=row;return row;
}
function sstEnsurePeople_(ss,people){
 SST_ROUNDS.forEach(title=>{const s=ss.getSheetByName(title);if(!s)throw Error('Missing source tab: '+title);
  const previous=s.getRange(3,1,s.getMaxRows()-2,1).getDisplayValues().flat().map(sstKey_);
  people.forEach(p=>{const isNew=!previous.includes(sstKey_(p.name));if(!isNew){SST_ROW_CACHE_[s.getSheetId()+':'+sstKey_(p.name)]=previous.indexOf(sstKey_(p.name))+3;return;}const row=sstFindRow_(s,p.name);
   s.getRange(row,2).setFormula(sstFormula_(row,'B','AR'));
   s.getRange(row,63).setFormula(sstFormula_(row,'BK','BO'));
   if(isNew&&row>3)s.getRange(3,45,1,18).copyTo(s.getRange(row,45,1,18),SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION,false);
  });
 });
}
function sstNamed_(ss,name){const r=ss.getRangeByName(name);if(!r)throw Error('A saved record marker is missing: '+name+'. No progress was overwritten.');return r;}
function sstMakeName_(ss,prefix,range){const name=prefix+'_'+Utilities.getUuid().replace(/-/g,'');ss.setNamedRange(name,range);return name;}
function sstSourceValues_(ss,p,n){const cached=SST_SOURCE_READ_CACHE_&&SST_SOURCE_READ_CACHE_[SST_ROUNDS[n-1]+':'+sstKey_(p.name)];if(cached)return cached.values.slice();const s=ss.getSheetByName(SST_ROUNDS[n-1]);return s.getRange(sstFindRow_(s,p.name),45,1,18).getDisplayValues()[0];}
function sstColumn_(c){let x='';while(c){c--;x=String.fromCharCode(65+c%26)+x;c=Math.floor(c/26);}return x;}
function sstFormulaText_(s){return String(s).replace(/"/g,'""');}
function sstSourceFormula_(name,n,offset){
 const tab="'"+SST_ROUNDS[n-1]+"'",col=sstColumn_(45+offset),q=sstFormulaText_(name);
 const value='INDEX('+tab+'!$'+col+'$3:$'+col+'$1000,MATCH("'+q+'",INDEX(TRIM('+tab+'!$A$3:$A$1000),0),0))';
 return '=IFNA(IF('+value+'="","",'+value+'),"SOURCE MISSING")';
}
function sstBlockInputs_(range){
 const s=range.getSheet(),r=range.getRow();
 return s.getRange(r+3,2,3,7).getDisplayValues().some((v,i)=>v.some((x,j)=>x.trim()&&!((i===0||i===1)&&j===2)&&!(i===2&&j===1)))||
 s.getRange(sstProgressBounds_(range).first,1,sstProgressBounds_(range).count,8).getDisplayValues().flat().some(v=>v.trim())||String(s.getRange(sstProgressBounds_(range).review,2).getDisplayValue()).trim();
}
function sstImportRecord_(ss,p,n,rec){
 if(rec.autoVersion===2)return rec;
 const range=sstNamed_(ss,rec.rangeName),s=range.getSheet(),start=range.getRow(),data=range.getDisplayValues();
 const recovery=JSON.stringify({name:p.name,round:n,values:data});for(let i=0;i<recovery.length;i+=7000)PropertiesService.getDocumentProperties().setProperty('SST_RECOVERY_'+rec.rangeName+'_'+(i/7000),recovery.slice(i,i+7000));
 const heads=ss.getSheetByName(SST_ROUNDS[n-1]).getRange(2,45,1,18).getDisplayValues()[0];let blocks=[];
 data.forEach((v,i)=>{const m=String(v[0]).match(/^Goal \d+ — (.*)$/);if(!m)return;
  const index=heads.findIndex((h,j)=>j%2===0&&h.replace(/\s*goal\s*$/i,'').trim()===m[1].trim());if(index<0)throw Error('Cannot match saved goal area: '+m[1]);
  const row=start+i,goal=String(s.getRange(row+1,2).getDisplayValue()),plan=String(s.getRange(row+2,2).getDisplayValue());
  blocks.push({id:Utilities.getUuid(),rangeName:sstMakeName_(ss,'SSTGOAL',s.getRange(row,1,24,8)),sourceColumn:index+45,area:m[1],goal:goal,plan:plan==='Plan not entered yet.'?'':plan,version:1,active:true});
 });
 const hasInput=blocks.some(b=>sstBlockInputs_(sstNamed_(ss,b.rangeName)))||s.getRange(start+1,2).getDisplayValue().trim()||s.getRange(start+1,4).getDisplayValue().trim()||s.getRange(start+1,6).getDisplayValue().trim();
 s.insertRowsAfter(start+1,2);
 const expanded=sstNamed_(ss,rec.rangeName); // All existing content and named goal blocks move with inserted rows.
 rec.autoVersion=2;rec.state=hasInput?'started':'draft';rec.blocks=blocks;
 rec.controlName=sstMakeName_(ss,'SSTCONTROL',s.getRange(start+2,1,2,8));
 let snapshot=Array(18).fill('');blocks.forEach(b=>{snapshot[b.sourceColumn-45]=b.goal;snapshot[b.sourceColumn-44]=b.plan;});rec.snapshot=snapshot;
 if(hasInput)rec.startedAt='Existing meeting preserved during upgrade';
 return rec;
}
function sstControls_(ss,p,n,rec){
 const range=sstNamed_(ss,rec.rangeName),s=range.getSheet(),r=range.getRow(),control=sstNamed_(ss,rec.controlName),c=control.getRow();
 s.getRange(c,1).setValue('Start this meeting').setFontWeight('bold');s.getRange(c,2).insertCheckboxes().setValue(rec.state==='started').setBackground(SST_COLORS.input);
 s.getRange(c+1,1).setValue('Revise goal / plan').setFontWeight('bold');s.getRange(c+1,2).insertCheckboxes().setValue(false).setBackground(SST_COLORS.input);sstReviseHelp_(s,c+1,n);
 s.getRange(c,3,1,6).breakApart().merge();s.getRange(c+1,3,1,6).breakApart().merge();
 s.setRowHeights(c,2,48);s.getRange(c,1,2,8).setFontFamily('Arial').setFontSize(11).setWrap(true).setVerticalAlignment('middle');
 s.getRange(r,9,1,18).setFormulas([Array.from({length:18},(_,i)=>sstSourceFormula_(p.name,n,i))]);
 s.getRange(r+1,9,1,18).setValues([rec.snapshot||Array(18).fill('')]);s.hideColumns(9,s.getMaxColumns()-8);
 const count='SUMPRODUCT(--(TRIM(\''+SST_ROUNDS[n-1]+'\'!$A$3:$A$1000)="'+sstFormulaText_(p.name)+'"))';
 const missing=Array.from({length:9},(_,i)=>'AND('+sstColumn_(9+i*2)+r+'<>"",'+sstColumn_(10+i*2)+r+'="")').join(',');
 const orphan=Array.from({length:9},(_,i)=>'AND('+sstColumn_(9+i*2)+r+'="",'+sstColumn_(10+i*2)+r+'<>"")').join(',');
 const mismatch='SUMPRODUCT(--(I'+r+':Z'+r+'<>I'+(r+1)+':Z'+(r+1)+'))>0';
 const draft='IF(OR('+missing+'),"NEEDS ATTENTION: a goal has no intervention plan.",IF(OR('+orphan+'),"NEEDS ATTENTION: an intervention has no goal.",IF(COUNTA(I'+r+':Z'+r+')=0,"DRAFT: enter goals in '+SST_ROUNDS[n-1]+'.","DRAFT: updates automatically. Check Start this meeting when ready.")))';
 s.getRange(c,3).setFormula('=IF('+count+'<>1,"CONNECTION PROBLEM: student name missing or duplicated.",'+(rec.state==='started'?'"STARTED: saved plan and progress stay together. Edit the meeting date if needed."':draft)+')');
 s.getRange(c+1,3).setFormula('=IF('+count+'<>1,"CONNECTION PROBLEM: check the source tab.",'+(rec.state==='started'?'IF('+mismatch+',"PLAN CHANGED: review '+SST_ROUNDS[n-1]+', then check Revise goal / plan. Old results stay with the old wording.",IF(OR('+missing+'),"NEEDS ATTENTION: a saved goal has no intervention plan.","Saved plan. No source changes."))':'"Before starting, keep editing goals and plans in '+SST_ROUNDS[n-1]+'. Progress rows appear after you start."')+')');
 s.getRange(r,1).setValue('SST '+n+' — '+(rec.state==='started'?'Started meeting':'Draft meeting')).setBackground(rec.state==='started'?SST_COLORS.green:'#476d8c');
 s.getRange(c+2,1,1,8).setValue(rec.state==='started'?'Record dated checks below. For a changed plan, use Revise goal / plan so earlier results keep their original meaning.':'Review this draft. Nothing here means the meeting has happened yet. Start it when the plan is ready.').setFontColor('#61717d').setWrap(true);
 s.getRange('A3:H3').setValue('Goals create a draft automatically. Start the meeting here when ready. Each SST and goal version keeps its own progress.').setWrap(true);
 sstRecordRules_(ss,rec);
}
function sstRecordRules_(ss,rec){
 const range=sstNamed_(ss,rec.rangeName),s=range.getSheet(),c=sstNamed_(ss,rec.controlName).getRow();
 let rules=s.getConditionalFormatRules().filter(rule=>!rule.getRanges().some(r=>r.getColumn()===3&&r.getRow()===c&&r.getNumRows()===2));
 const area=s.getRange(c,3,2,6);
 rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains('CONNECTION PROBLEM').setBackground('#fce4e4').setFontColor('#9c2020').setRanges([area]).build());
 rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains('PLAN CHANGED').setBackground('#ffe6c7').setFontColor('#854300').setRanges([area]).build());
 rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains('NEEDS ATTENTION').setBackground('#fff1bd').setFontColor('#6e5200').setRanges([area]).build());s.setConditionalFormatRules(rules);
}
function sstNewBlock_(ss,rec,g,version){
 const range=sstNamed_(ss,rec.rangeName),s=range.getSheet(),after=range.getLastRow();s.insertRowsAfter(after,24);const r=after+1;
 ss.setNamedRange(rec.rangeName,s.getRange(range.getRow(),1,range.getNumRows()+24,8));
 function field(row,label,value){s.getRange(row,1).setValue(label).setFontWeight('bold');s.getRange(row,2,1,7).merge().setValue(value).setWrap(true);s.setRowHeight(row,Math.max(48,Math.ceil(String(value).length/125)*18+16));}
 s.getRange(r,1,1,8).merge().setValue('Current goal — '+g.area+' (version '+version+')').setBackground(SST_COLORS.green).setFontColor('white').setFontWeight('bold');s.setRowHeight(r,35);
 field(r+1,'Student goal',g.goal);field(r+2,'Intervention plan',g.plan||'Plan not entered yet.');
 const method=['Word reading check','Spelling dictation','Writing sample / rubric','Math quiz','Work sample','Reading passage','Observation','Other / my words'];
 ['Assessment','Frequency'].forEach((label,i)=>{const row=r+3+i;s.getRange(row,1).setValue(label).setFontWeight('bold');s.getRange(row,2,1,2).merge().setDataValidation(sstRule_(i?['Daily','Twice weekly','Weekly','Every two weeks','As scheduled','Other / my words']:method)).setBackground(SST_COLORS.input);s.getRange(row,4).setValue(i?'Other frequency':'Other method');s.getRange(row,5,1,4).merge().setBackground(SST_COLORS.input);s.setRowHeight(row,42);});
 s.getRange(r+5,1,1,3).setValues([['Baseline date','','Baseline result']]);s.getRange(r+5,2).setNumberFormat('mm/dd/yy').setBackground(SST_COLORS.input);s.getRange(r+5,4,1,5).merge().setBackground(SST_COLORS.input);s.setRowHeight(r+5,42);
 s.getRange(r+6,1,1,8).setValues([['Date','Result / score','Progress status','Assessment type','Support provided','Observation','Custom status / other','My notes / evidence']]).setBackground(SST_COLORS.blue).setFontWeight('bold');s.setRowHeight(r+6,45);
 const table=s.getRange(r+7,1,15,8);table.setBackground('#fffdf5').setBorder(true,true,true,true,true,true,'#d5dfe4',SpreadsheetApp.BorderStyle.SOLID).setVerticalAlignment('top');
 s.getRange(r+7,1,15,1).setNumberFormat('mm/dd/yy');s.getRange(r+7,2,15,1).setNumberFormat('@');
 [[3,['Not yet met','Making progress','Met this check','Goal met — criteria checked','Not assessed','Absent','Other / my words']],[4,method],[5,['Independent','Verbal prompt','Visual support','Model / example','Guided practice','One-to-one support','Other / my words']],[6,['More accurate','More fluent','More independent','Needs reteaching','Inconsistent performance','No clear change yet','Other / my words']]].forEach(x=>s.getRange(r+7,x[0],15,1).setDataValidation(sstRule_(x[1])));
 s.setRowHeights(r+7,15,60);field(r+22,'Goal review / next step','');s.getRange(r+22,2,1,7).setBackground(SST_COLORS.input);
 s.getRange(r,1,24,8).setFontFamily('Arial').setFontSize(11).setWrap(true).setVerticalAlignment('top');
 const block={id:Utilities.getUuid(),rangeName:sstMakeName_(ss,'SSTGOAL',s.getRange(r,1,24,8)),sourceColumn:g.sourceColumn,area:g.area,goal:g.goal,plan:g.plan,version:version,active:true};rec.blocks.push(block);sstBlockRules_(ss,block);return block;
}
function sstBlockRules_(ss,b){
 if(b.rulesInstalled)return;const range=sstNamed_(ss,b.rangeName),s=range.getSheet(),r=range.getRow(),first=r+7,last=first+14;let rules=s.getConditionalFormatRules();
 rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND(COUNTA($B'+first+':$H'+first+')>0,$A'+first+'="")').setBackground('#fce4e4').setRanges([s.getRange(first,1,15,1)]).build());
 rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND(COUNTIF($C'+first+':$F'+first+',"Other / my words")>0,COUNTA($G'+first+':$H'+first+')=0)').setBackground('#fff1bd').setRanges([s.getRange(first,7,15,2)]).build());
 [r+3,r+4].forEach(row=>rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND($B'+row+'="Other / my words",$E'+row+'="")').setBackground('#fff1bd').setRanges([s.getRange(row,5,1,4)]).build()));
 s.setConditionalFormatRules(rules);b.rulesInstalled=true;
}
function sstDraft_(ss,p,n,rec,goals){
 const s=sstNamed_(ss,rec.rangeName).getSheet();
 rec.blocks.forEach(b=>{const range=sstNamed_(ss,b.rangeName),g=goals.find(g=>g.sourceColumn===b.sourceColumn);b.active=!!g;
  if(!g){s.hideRows(range.getRow(),24);return;}
  b.goal=g.goal;b.plan=g.plan;s.showRows(range.getRow(),3);s.hideRows(range.getRow()+3,21);
  const helperRow=sstNamed_(ss,rec.rangeName).getRow(),col=sstColumn_(9+b.sourceColumn-45);
  s.getRange(range.getRow(),1).setValue('Draft goal — '+g.area).setBackground('#476d8c');
  s.getRange(range.getRow()+1,2).setFormula('='+col+helperRow);
  const planCol=sstColumn_(10+b.sourceColumn-45);s.getRange(range.getRow()+2,2).setFormula('=IF('+planCol+helperRow+'="","Plan not entered yet.",'+planCol+helperRow+')');
  s.setRowHeight(range.getRow()+1,Math.max(48,Math.ceil(g.goal.length/125)*18+16));s.setRowHeight(range.getRow()+2,Math.max(48,Math.ceil(g.plan.length/125)*18+16));
 });
 goals.filter(g=>!rec.blocks.some(b=>b.sourceColumn===g.sourceColumn)).forEach(g=>{sstNewBlock_(ss,rec,g,1);});
 // Newly appended blocks also get their live links and compact draft view.
 rec.blocks.filter(b=>b.active).forEach(b=>{const range=sstNamed_(ss,b.rangeName),r=range.getRow(),hr=sstNamed_(ss,rec.rangeName).getRow(),gc=sstColumn_(9+b.sourceColumn-45),pc=sstColumn_(10+b.sourceColumn-45);s.getRange(r+1,2).setFormula('='+gc+hr);s.getRange(r+2,2).setFormula('=IF('+pc+hr+'="","Plan not entered yet.",'+pc+hr+')');s.showRows(r,3);s.hideRows(r+3,21);sstBlockRules_(ss,b);});
 rec.goals=goals.length;
}
function sstRecord_(ss,p,n,goals,activate){
 let registry=sstRegistry_(),entry=registry[sstKey_(p.name)],rec=entry&&entry.records[n];if(entry&&entry.archived)return rec?sstNamed_(ss,rec.rangeName):null;
 if(rec&&rec.autoVersion===2&&(rec.state==='started'||rec.draftSignature===JSON.stringify(goals))){const r=sstNamed_(ss,rec.rangeName);if(activate){ss.setActiveSheet(r.getSheet());r.getSheet().getRange(r.getRow(),1).activate();}return r;}
 const isNew=!rec;
 if(!rec){if(!goals.length)return null;sstLegacyRecord_(ss,p,n,goals,false);registry=sstRegistry_();entry=registry[sstKey_(p.name)];rec=entry.records[n];}
 rec=sstImportRecord_(ss,p,n,rec);sstControls_(ss,p,n,rec);
 if(rec.state==='draft')sstDraft_(ss,p,n,rec,goals);
 else rec.blocks.forEach(b=>sstBlockRules_(ss,b));
 rec.draftSignature=JSON.stringify(goals);
 registry=sstRegistry_();registry[sstKey_(p.name)].records[n]=rec;sstSaveRegistry_(registry);
 const range=sstNamed_(ss,rec.rangeName);
 if(isNew){const later=[1,2,3].filter(x=>x>n&&registry[sstKey_(p.name)].records[x]).map(x=>sstNamed_(ss,registry[sstKey_(p.name)].records[x].rangeName).getRow());if(later.length&&Math.min(...later)<range.getRow())range.getSheet().moveRows(range,Math.min(...later));}
 sstStyleProgress_(ss,range.getSheet());
 if(activate){ss.setActiveSheet(range.getSheet());range.getSheet().getRange(range.getRow(),1).activate();}
 return range;
}
function sstContext_(ss){
 const active=ss.getActiveSheet(),row=active.getActiveRange().getRow(),registry=sstRegistry_(),people=sstPeople_(ss);
 for(const p of people){const e=registry[sstKey_(p.name)];if(!e||e.sheetId!==active.getSheetId())continue;
  for(const n of [1,2,3]){const rec=e.records[n];if(!rec)continue;const r=sstNamed_(ss,rec.rangeName);if(row>=r.getRow()&&row<=r.getLastRow())return{p,n,rec};}
  const available=[1,2,3].filter(n=>e.records[n]);if(available.length===1)return{p,n:available[0],rec:e.records[available[0]]};
  throw Error('Click inside the SST meeting section you want, then use this command.');
 }
 const selected=sstSelection_(ss),entry=registry[sstKey_(selected.p.name)];if(!entry||!entry.records[selected.n])throw Error('Enter goals for this student and round first.');return{...selected,rec:entry.records[selected.n]};
}
function sstStart_(ss,p,n){
 const goals=sstGoals_(ss,p.name,n);if(!goals.length)throw Error('Enter at least one goal before starting this meeting.');sstRecord_(ss,p,n,goals,false);
 const registry=sstRegistry_(),rec=registry[sstKey_(p.name)].records[n];if(rec.state==='started')return;
 rec.snapshot=sstSourceValues_(ss,p,n);rec.state='started';rec.startedAt=new Date().toISOString();
 rec.blocks.forEach(b=>{const range=sstNamed_(ss,b.rangeName),s=range.getSheet(),r=range.getRow();if(!b.active)return;s.showRows(r,24);s.getRange(r,1).setValue('Current goal — '+b.area+' (version '+b.version+')').setBackground(SST_COLORS.green);s.getRange(r+1,2).setValue(b.goal);s.getRange(r+2,2).setValue(b.plan||'Plan not entered yet.');});
 const range=sstNamed_(ss,rec.rangeName),date=range.getSheet().getRange(range.getRow()+1,2);if(!date.getValue())date.setValue(new Date()).setNumberFormat('mm/dd/yy');
 registry[sstKey_(p.name)].records[n]=rec;sstSaveRegistry_(registry);sstControls_(ss,p,n,rec);sstIndexRow_(ss,p);ss.toast('SST '+n+' started. The plan is saved. Record progress below.','SST Progress',7);
}
function sstStartMeeting(){try{const ss=sstBook_();if(!sstRequireProgressTab_(ss,'Start a draft meeting'))return;const x=sstContext_(ss);sstStart_(ss,x.p,x.n);}catch(e){SpreadsheetApp.getUi().alert(e.message);}}
function sstRevise_(ss,p,n){
 let registry=sstRegistry_(),rec=registry[sstKey_(p.name)].records[n];if(rec.state!=='started')throw Error('This is still a draft. Edit goals in the SST tab, then Start this meeting.');
 const values=sstSourceValues_(ss,p,n);if(JSON.stringify(values)===JSON.stringify(rec.snapshot)){ss.toast('The saved plan already matches the source. No revision was added.','SST Progress',6);return;}
 const goals=sstGoals_(ss,p.name,n),s=sstNamed_(ss,rec.rangeName).getSheet();
 rec.blocks.filter(b=>b.active).forEach(b=>{const g=goals.find(g=>g.sourceColumn===b.sourceColumn);if(g&&g.goal===b.goal&&g.plan===b.plan)return;b.active=false;const row=sstNamed_(ss,b.rangeName).getRow();s.getRange(row,1).setValue('Earlier goal — '+b.area+' (version '+b.version+')').setBackground('#657783');});
 goals.forEach(g=>{if(rec.blocks.some(b=>b.active&&b.sourceColumn===g.sourceColumn))return;const previous=rec.blocks.filter(b=>b.sourceColumn===g.sourceColumn);const b=sstNewBlock_(ss,rec,g,Math.max(0,...previous.map(b=>b.version))+1);b.revisedAt=new Date().toISOString();});
 rec.snapshot=values;rec.goals=goals.length;rec.revisedAt=new Date().toISOString();registry[sstKey_(p.name)].records[n]=rec;sstSaveRegistry_(registry);sstControls_(ss,p,n,rec);sstIndexRow_(ss,p);sstStyleProgress_(ss,s);ss.toast('Revised plan saved. Earlier goal wording and progress were kept. Use the new Current goal block for new results.','SST Progress',10);
}
function sstReviseMeeting(){try{const ss=sstBook_();if(!sstRequireProgressTab_(ss,'Save a revised goal or plan'))return;const x=sstContext_(ss);sstRevise_(ss,x.p,x.n);}catch(e){SpreadsheetApp.getUi().alert(e.message);}}
function sstIndexRow_(ss,p){
 const s=ss.getSheetByName(SST_INDEX),people=SST_INDEX_PEOPLE_CACHE_||sstActivePeople_(ss),row=13+people.findIndex(x=>sstKey_(x.name)===sstKey_(p.name)),entry=sstRegistry_()[sstKey_(p.name)];if(!s||row<13)return;
 const links=[1,2,3].map(n=>{const rec=entry&&entry.records[n];if(!rec)return '';const range=sstNamed_(ss,rec.rangeName);return sstLink_(range.getSheet(),range.getRow(),(rec.state==='started'?'SST ':'Draft SST ')+n);});
 s.getRange(row,4,1,4).setValues([[entry&&entry.sheetId&&ss.getSheetById(entry.sheetId)?sstLink_(ss.getSheetById(entry.sheetId),1,'Open progress'):'',...links]]);
 const meetingStatus=sstMeetingCells_(ss,p,row,entry);let alerts=sstMeetingAlerts_(p,entry);
 for(const n of [1,2,3]){const rec=entry&&entry.records[n];if(rec&&rec.state==='started'&&JSON.stringify(sstSourceValues_(ss,p,n))!==JSON.stringify(rec.snapshot))alerts.push('SST '+n+' plan changed. Review and use Revise goal / plan.');
  const goals=sstGoals_(ss,p.name,n);if(goals.some(g=>!g.plan.trim()))alerts.push('SST '+n+': a goal has no plan.');
  const text=goals.map(g=>g.plan).join(' ').toLowerCase();const others=people.filter(q=>sstKey_(q.name)!==sstKey_(p.name)&&new RegExp('\\b'+sstKey_(q.name).replace(/[.*+?^${}()|[\]\\]/g,'\\$&')+'\\b','i').test(text));if(others.length)alerts.push('SST '+n+' plan mentions '+others.map(q=>q.name).join(', ')+'. Review the wording.');
 }
 s.getRange(row,8).setValue(alerts.join(' ')||(meetingStatus.checked?'Meeting number checked: SST '+meetingStatus.round+'.':'Enter goals in the correct SST tab. Start the meeting from the student’s Progress tab.')).setWrap(true).setBackground(alerts.length?'#fff1bd':'white');
}
function sstIndex_(ss,people){const priorSources=SST_SOURCE_READ_CACHE_,priorPeople=SST_INDEX_PEOPLE_CACHE_;try{
 people=people.filter(p=>!(sstRegistry_()[sstKey_(p.name)]||{}).archived);SST_INDEX_PEOPLE_CACHE_=people;if(!SST_SOURCE_READ_CACHE_)SST_SOURCE_READ_CACHE_=sstCaseSourceReads_(ss);
 const active=ss.getActiveSheet(),selection=active.getActiveRange().getA1Notation(),props=PropertiesService.getDocumentProperties(),signature=JSON.stringify(people.map(p=>[p.name,p.values[3]]));const previous=props.getProperty('SST_INDEX_ROSTER_V2'),rebuild=signature!==previous;
 if(rebuild){const old=ss.getSheetByName(SST_INDEX);if(old){let count=previous?JSON.parse(previous).length:0;if(!count){const names=old.getRange(13,1,Math.max(1,old.getLastRow()-12),1).getDisplayValues().flat();while(count<names.length&&String(names[count]).trim())count++;}const ownedEnd=Math.max(28,count+23);old.getRange(13,1,ownedEnd-12,8).breakApart().clearContent();old.getRange(13,9,ownedEnd-12,2).clearContent().clearDataValidations();}sstLegacyIndex_(ss,people);}const s=ss.getSheetByName(SST_INDEX);
 s.getRange('C3:H4').setValue('Optional shortcuts. Enter goals in an SST tab to create a draft automatically. Open the student’s Progress tab and check Start this meeting when ready.');
 s.getRange('A6:H6').setValue('Optional Writing evidence: choose student + meeting, check the pieces, then use Add checked Writing evidence.');
 s.getRange('A10:H10').setValue('Drafts update automatically. Started meetings keep their saved plan. Orange alerts mean source changes need review.').setBackground(SST_COLORS.pale).setFontColor(SST_COLORS.ink);
 let h=people.length+16;const help=[['1. General','Add each student and shared background here. New names are added to the SST tabs automatically. Keep names unique and consistent.'],['2. SST tabs','Enter goals and intervention plans in AS:BJ. Repeated background columns are hidden, not deleted.'],['3. Automatic draft','A matching draft appears in that student’s Progress tab. Keep editing the source while preparing.'],['4. Start this meeting','Open the student’s Progress tab. Check Start this meeting for the correct round. Progress rows open and the plan is saved.'],['5. Changed plan','After starting, source edits show a PLAN CHANGED alert. Review the source and check Revise goal / plan. Old results stay under the earlier wording.'],['6. Next meeting','Enter goals in the next SST tab. Its draft appears in the same Progress tab. You may also copy the previous plan with the menu command.'],['7. Print the folder copy','On the student’s Progress tab, use Prepare this student for printing, then File → Print → Current sheet → Landscape → Fit to width. Return to progress entry view afterward.']];
 if(rebuild){help.forEach((r,i)=>{s.getRange(h+i,1).setValue(r[0]);s.getRange(h+i,2,1,7).setValue(r[1]);s.setRowHeight(h+i,55);});props.setProperty('SST_INDEX_ROSTER_V2',signature);sstWarningIndex_(ss);sstMaybeRefreshNativeTabs_(ss);}sstMeetingHeader_(ss);people.forEach(p=>sstIndexRow_(ss,p));ss.setActiveSheet(active);active.getRange(selection).activate();
}finally{SST_SOURCE_READ_CACHE_=priorSources;SST_INDEX_PEOPLE_CACHE_=priorPeople;}}
function sstReconcile_(ss){const priorSources=SST_SOURCE_READ_CACHE_;try{
 sstRepairNameConnections_(ss);const people=sstPeople_(ss);sstEnsurePeople_(ss,people);if(PropertiesService.getDocumentProperties().getProperty('SST_NAME_BINDINGS_VERSION')==='1')sstBindStudentIds_(ss,people);
 SST_SOURCE_READ_CACHE_=sstCaseSourceReads_(ss);people.filter(p=>!(sstRegistry_()[sstKey_(p.name)]||{}).archived).forEach(p=>[1,2,3].forEach(n=>{const goals=sstGoals_(ss,p.name,n),e=sstRegistry_()[sstKey_(p.name)];if(goals.length||(e&&e.records[n]))sstRecord_(ss,p,n,goals,false);}));sstIndex_(ss,people);return people;
}finally{SST_SOURCE_READ_CACHE_=priorSources;}}
function sstFailure_(e){try{const ss=sstBook_(),s=ss.getSheetByName(SST_INDEX);s.getRange('A10:H10').setValue('UPDATE PROBLEM: '+e.message+' Nothing should be deleted to fix this.').setBackground('#fce4e4').setFontColor('#9c2020');ss.toast(e.message,'SST update needs attention',10);}catch(ignore){}console.error(e.stack||e.message);}
function onEdit(e){
 if(!e||!e.range||!e.source||e.source.getId()!==SST_COPY_ID)return;const lock=LockService.getDocumentLock();if(!lock.tryLock(5000))return;
 try{
  const ss=e.source,s=e.range.getSheet(),title=s.getName();
  if(typeof sstRubricOnEdit_==='function')sstRubricOnEdit_(e);
  if(title==='General'&&e.range.getColumn()===1&&e.range.getLastRow()>=3&&PropertiesService.getDocumentProperties().getProperty('SST_NAME_BINDINGS_VERSION')==='1'){sstGeneralNamesEdited_(e);return;}
  const people=sstPeople_(ss);
  if(title===SST_INDEX&&e.range.getLastRow()>=13&&e.range.getColumn()<=10&&e.range.getLastColumn()>=9){sstMeetingEdit_(ss,e.range);return;}
  if(title==='General'&&e.range.getColumn()<=4&&e.range.getLastRow()>=3){if(e.range.getColumn()===1)sstEnsurePeople_(ss,people);sstIndex_(ss,people);return;}
  if(SST_ROUNDS.includes(title)&&e.range.getLastRow()>=3&&e.range.getColumn()<=62&&e.range.getLastColumn()>=45){
   const n=SST_ROUNDS.indexOf(title)+1;for(let row=Math.max(3,e.range.getRow());row<=e.range.getLastRow();row++){const name=sstKey_(s.getRange(row,1).getDisplayValue());if(!name)continue;const p=people.find(p=>sstKey_(p.name)===name);if(!p)throw Error('This name is missing from General: '+name);sstRecord_(ss,p,n,sstGoals_(ss,p.name,n),false);sstIndexRow_(ss,p);}return;
  }
  if(title.startsWith('Progress - ')){
   const registry=sstRegistry_();for(const p of people){const entry=registry[sstKey_(p.name)];if(!entry||entry.sheetId!==s.getSheetId())continue;
    for(const n of [1,2,3]){const rec=entry.records[n];if(!rec||!rec.controlName)continue;const c=sstNamed_(ss,rec.controlName).getRow();if(e.range.getColumn()===2&&e.range.getRow()===c){if(e.range.getValue()===true)sstStart_(ss,p,n);else{e.range.setValue(rec.state==='started');ss.toast('Unchecking does not erase or restart a saved meeting.','SST Progress',6);}return;}
     if(e.range.getColumn()===2&&e.range.getRow()===c+1&&e.range.getValue()===true){e.range.setValue(false);sstRevise_(ss,p,n);return;}
    }
   }
  }
 }catch(error){if(e.range.getSheet().getName().startsWith('Progress - ')&&e.range.getColumn()===2){const registry=sstRegistry_();Object.values(registry).forEach(entry=>{if(entry.sheetId!==e.range.getSheet().getSheetId())return;Object.values(entry.records).forEach(rec=>{if(!rec.controlName)return;const c=sstNamed_(e.source,rec.controlName).getRow();if(e.range.getRow()===c)e.range.setValue(rec.state==='started');if(e.range.getRow()===c+1)e.range.setValue(false);});});}sstFailure_(error);}finally{lock.releaseLock();}
}
function sstCreateSelected(){const ss=sstBook_(),x=sstSelection_(ss),goals=sstGoals_(ss,x.p.name,x.n);const range=sstRecord_(ss,x.p,x.n,goals,true);if(!range)SpreadsheetApp.getUi().alert('Enter goals in the correct SST tab first. A draft will appear automatically.');}
function sstNextDraft(){const ss=sstBook_();if(!sstRequireProgressTab_(ss,'Copy a plan to the next SST draft'))return;const x=sstContext_(ss);if(x.n===3){SpreadsheetApp.getUi().alert('This workbook currently supports SST1–SST3 only.');return;}const to=ss.getSheetByName(SST_ROUNDS[x.n]),row=sstFindRow_(to,x.p.name),dest=to.getRange(row,45,1,18);if(dest.getDisplayValues()[0].some(v=>v.trim())){SpreadsheetApp.getUi().alert('The next SST already contains goals or plans. Nothing was overwritten.');return;}dest.setValues([sstSourceValues_(ss,x.p,x.n)]);sstRecord_(ss,x.p,x.n+1,sstGoals_(ss,x.p.name,x.n+1),false);sstIndexRow_(ss,x.p);ss.setActiveSheet(to);dest.activate();ss.toast('Next draft created automatically. Edit this round’s plan, then Start this meeting from the Progress tab.','SST Progress',9);}
function sstRefresh(){const ss=sstBook_();try{if(!sstSafeStartup_(30000))throw Error('An update is still running. Try again in a moment.');ss.toast('Drafts and connections checked. Started plans and progress were preserved.','SST Progress',7);}catch(e){sstFailure_(e);}}
function sstSetup(){sstRefresh();}
function sstPreparePrint(){
 const ss=sstBook_(),sheet=ss.getActiveSheet(),entry=Object.values(sstRegistry_()).find(e=>e.sheetId===sheet.getSheetId());if(!entry){SpreadsheetApp.getUi().alert('Open a student’s Progress tab first.');return;}
 const records=Object.values(entry.records);if(!records.some(r=>r.state==='started')){ss.toast('No meeting has been started yet. Review the draft and Start this meeting first.','SST Progress',8);return;}
 records.forEach(rec=>{const range=sstNamed_(ss,rec.rangeName);if(rec.state!=='started'){sheet.hideRows(range.getRow(),range.getNumRows());return;}sheet.showRows(range.getRow(),range.getNumRows());rec.blocks.forEach(b=>{const blockRange=sstNamed_(ss,b.rangeName),bounds=sstProgressBounds_(blockRange),r=blockRange.getRow(),rows=sheet.getRange(bounds.first,1,bounds.count,8).getDisplayValues();let blankShown=false;rows.forEach((v,i)=>{if(v.some(x=>x.trim()))return;if(!blankShown){blankShown=true;return;}sheet.hideRows(bounds.first+i);});});});
 PropertiesService.getDocumentProperties().setProperty('SST_PRINT_'+sheet.getSheetId(),'yes');ss.toast('Print view ready: started meetings only, extra empty progress rows hidden. File → Print → Current sheet → Landscape → Fit to width. Afterward use Return to progress entry view.','SST Progress',15);
}
function sstEntryView(){const ss=sstBook_(),sheet=ss.getActiveSheet(),entry=Object.values(sstRegistry_()).find(e=>e.sheetId===sheet.getSheetId());if(!entry)return;Object.values(entry.records).forEach(rec=>{const range=sstNamed_(ss,rec.rangeName);sheet.showRows(range.getRow(),range.getNumRows());rec.blocks.forEach(b=>{const r=sstNamed_(ss,b.rangeName).getRow();if(rec.state==='draft'){if(b.active)sheet.hideRows(r+3,21);else sheet.hideRows(r,24);}});});PropertiesService.getDocumentProperties().deleteProperty('SST_PRINT_'+sheet.getSheetId());ss.toast('Progress entry rows restored.','SST Progress',6);}
function sstUpgradeAutomatic(){
 const ss=sstBook_(),lock=LockService.getDocumentLock();lock.waitLock(30000);
 try{sstReconcile_(ss);SST_ROUNDS.forEach(title=>{const s=ss.getSheetByName(title);s.hideColumns(2,43);s.hideColumns(63,5);s.setFrozenColumns(1);
  const rules=s.getConditionalFormatRules();for(let c=45;c<=61;c+=2){const goal=sstColumn_(c),plan=sstColumn_(c+1);rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND('+goal+'3<>"",'+plan+'3="")').setBackground('#fff1bd').setRanges([s.getRange(3,c+1,998,1)]).build());rules.push(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND('+goal+'3="",'+plan+'3<>"")').setBackground('#fff1bd').setRanges([s.getRange(3,c,998,1)]).build());}s.setConditionalFormatRules(rules);
 });sstUpdateGuide_();PropertiesService.getDocumentProperties().setProperty('SST_AUTO_VERSION','2');onOpen();const first=sstPeople_(ss)[0],entry=sstRegistry_()[sstKey_(first.name)],rec=Object.values(entry.records)[0],range=sstNamed_(ss,rec.rangeName);ss.setActiveSheet(range.getSheet());range.getSheet().getRange(range.getRow(),1).activate();SpreadsheetApp.flush();console.log('Automatic drafts installed. Existing sources, Writing dropdowns, saved evidence, and entered results were preserved.');
 }finally{lock.releaseLock();}
}
function sstUpdateGuide_(){
 const ss=sstBook_(),s=ss.getSheetByName('Q&A - Start Here');if(!s)return;const rows=s.getRange(4,1,s.getLastRow()-3,4).getValues();
 const answers={
 'What belongs in an SST tab?':['Enter that meeting’s goals and intervention plans in AS:BJ on the correct student’s row. Repeated background columns are hidden; General still supplies that information. A draft appears automatically in the student’s Progress tab.','Works now'],
 'Where do I record progress?':['Open the student’s Progress tab, find the correct SST round, and check Start this meeting when ready. Dated progress rows then appear beneath each saved goal. Use dropdowns and your own notes.','Works now'],
 'How do I add a student today?':['1. Add a unique name in a new General row. 2. Add the student’s background. Their name is added to all SST tabs automatically. 3. Enter goals/plans in the correct SST tab. A draft appears in their Progress tab. 4. Review it and check Start this meeting.','Works now'],
 'Will a new student appear in all three SST tabs?':['Yes, after entering the name in General, their name and shared background are added automatically. Allow a few seconds. Goals stay separate for each round. A listing alone does not mean a meeting happened.','Works now'],
 'Will adding a student become completely automatic?':['This is now installed. Typing a new, unique name in General adds them to the SST tabs. Entering goals creates their draft Progress section. No routine Index or refresh step is needed.','Works now'],
 'Which changes update automatically today?':['Shared background updates from General. New names are added to the SST tabs. Goals create matching draft sections automatically. Draft goals/plans stay linked. Started meetings keep their saved plan and show an alert when the source changes.','Works now'],
 'Why did a changed SST3 goal not appear in Progress?':['That was a limitation of the earlier saved-copy setup. Drafts now update automatically. After Start this meeting, the saved plan stays unchanged deliberately; a PLAN CHANGED alert prompts you to review and use Revise goal / plan.','Works now'],
 'Will goals and intervention plans update automatically in Progress?':['Yes while the meeting is a draft. After you Start this meeting, source changes produce an orange PLAN CHANGED alert. Use Revise goal / plan to approve the change. Earlier goal wording and progress remain together.','Works now'],
 'Can I fix the goal update by clicking Create or open again?':['Opening a meeting does not approve a changed started plan. Drafts update automatically. For a started meeting, review the source edit and check Revise goal / plan in that meeting’s Progress section.','Works now'],
 'What happens if a goal changes after I recorded progress?':['Edit the correct SST source. The saved goal and results stay intact, and an orange PLAN CHANGED alert appears. Check Revise goal / plan when ready. The earlier block stays visible; a new Current goal block receives future results.','Works now'],
 'Have the repeated background columns been hidden yet?':['Yes. In SST 1, SST2, and SST3, the repeated background columns are hidden, not deleted. Names and AS:BJ goal/plan columns stay visible. Edit shared background in General.','Works now'],
 'How do I prepare the next SST without erasing the previous one?':['Enter goals directly in the next SST tab. Its draft appears automatically in the same student’s Progress tab. Or click in the current meeting and use SST Case Manager → Copy goals to the next SST draft. Review the draft and Start this meeting when ready.','Works now'],
 'Why do I still see empty progress rows?':['Progress rows appear only after you start a meeting. They give you room for future checks. Prepare this student for printing hides extra empty progress rows and unstarted drafts. Return to progress entry view brings them back.','Works now'],
 'What if I add or remove a goal after the tracker was created?':['In a draft, goal blocks are added or hidden automatically. After starting, changes show PLAN CHANGED. Revise goal / plan saves new goal versions; removed goals remain as earlier history with their original results.','Works now'],
 'Are the new color alerts working now?':['Yes. Yellow flags missing goal/plan pairs or Other without custom wording. Red flags missing dates on entered results or a connection problem. Orange means a started plan differs from its SST source. Each meeting has a short message explaining what needs attention.','Works now'],
 'What if the SST Progress menu is missing?':['Make sure you are signed in with editing access and reload after edits finish saving. The menu should show Start this meeting and Revise goal / plan. If it does not, stop and check the bound automation rather than recreating records.','Troubleshooting'],
 'How would you simplify this if building it again?':['Your dropdowns are preserved. General is the shared student folder; SST tabs show compact meeting plans; each student has one Progress tab with separate rounds, automatic drafts, one Start decision, saved goal versions, and specific review alerts.','Works now'],
 'What would permanent student and goal IDs help with?':['Goal versions now have internal markers that keep old wording and results together. Student matching still uses unique names. Correcting a student’s name needs a coordinated update; permanent student IDs remain a possible future improvement.','Ideas to consider'],
 'What does keeping progress attached to its original goal mean?':['Example: You recorded 15 out of 20 for a goal about 20 words. Later you change to a goal about 50 words. After Revise goal / plan, the old 15 out of 20 stays with the earlier goal. A new block holds the 50-word goal and future checks.','Works now']
 };
 rows.forEach(r=>{if(answers[r[1]]){r[2]=answers[r[1]][0];r[3]=answers[r[1]][1];}});s.getRange(4,1,rows.length,4).setValues(rows);
 s.getRange('A2:D2').setValue('Automatic drafts are active. Start a meeting when ready. Revise a started plan without changing earlier results.');
 const additional=[
 ['Meeting workflow','Do I need to use Student Index every time?','No. Enter goals in the SST tab, then open the student’s Progress tab. The Index remains an optional shortcut and the place to choose optional writing evidence.','Works now'],
 ['Meeting workflow','How do I start a meeting?','Open the student’s Progress tab and find the right SST round. Review the draft. Check the box beside Start this meeting, or click inside that section and use SST Case Manager → Start this meeting. Progress rows appear and the plan is saved.','Works now'],
 ['Meeting workflow','Does the default meeting date mean the meeting already happened?','No. Starting fills today’s date only when the date is blank. Edit it to the actual meeting date. Preparing a draft does not complete a meeting.','Works now'],
 ['Goal revisions','When should I use Revise goal / plan?','After starting, use it when you approve a change of direction, a new goal, or a changed intervention. Edit the SST source first. Review the orange alert, then check Revise goal / plan. Before starting, just edit the draft freely.','Works now'],
 ['Goal revisions','What if I change only the intervention plan?','That also counts as a plan revision after starting. The earlier block keeps its original goal, intervention, and results. A new current block starts with the revised intervention for future checks.','Works now'],
 ['Goal revisions','What if I correct only a typo after starting?','The tracker cannot judge whether a wording change is a typo or a new direction. It shows the same review alert. Approving the correction creates a new version and preserves the old results.','Works now'],
 ['Printing','How do I print one student’s meeting history?','1. Open the student’s Progress tab. 2. Use SST Case Manager → Prepare this student for printing. 3. Choose File → Print. 4. Select Current sheet, Landscape, and Fit to width. 5. After printing, use Return to progress entry view. Draft meetings and extra blank progress rows are hidden in print view.','Works now'],
 ['Troubleshooting','What if an automatic draft did not appear?','Allow a few seconds after a normal edit. Check for an UPDATE PROBLEM message in Student Index, missing/duplicate names, or a missing source tab. Reopening the spreadsheet checks connections. Check automatic drafts and connections is an optional recovery command, not a routine step.','Troubleshooting'],
 ['Meeting workflow','Can I undo starting by unchecking Start this meeting?','No. Unchecking does not erase or unlock a saved meeting. This avoids accidentally changing the meaning of progress. Use Revise goal / plan for a changed started plan. Use version history carefully for a larger rollback.','Works now']
 ];
 let next=s.getLastRow()+1;additional.forEach(r=>{if(rows.some(x=>x[1]===r[1]))return;s.getRange(next,1,1,4).setValues([r]).setFontFamily('Arial').setFontSize(11).setFontColor(SST_COLORS.ink).setBackground(next%2?'#eff5f4':'white').setVerticalAlignment('top').setWrap(true);s.getRange(next,1,1,2).setFontWeight('bold');s.setRowHeight(next,Math.max(64,Math.ceil(r[2].length/85)*18+16));next++;});
 for(let i=0;i<rows.length;i++){const row=i+4;s.getRange(row,1,1,4).setBackground(rows[i][3]==='Pending'?'#fff8e6':row%2?'#eff5f4':'white');s.setRowHeight(row,Math.max(60,Math.ceil(String(rows[i][2]).length/85)*18+16));}
}

function sstFinishAutomatic(){
 const ss=sstBook_(),people=sstReconcile_(ss);sstUpdateGuide_();
 const registry=sstRegistry_();let report=[];
 people.forEach(p=>{const entry=registry[sstKey_(p.name)];if(!entry)return;Object.entries(entry.records).forEach(([n,rec])=>{const range=sstNamed_(ss,rec.rangeName),control=sstNamed_(ss,rec.controlName);report.push({student:p.name,round:n,state:rec.state,controlRow:control.getRow(),goals:rec.blocks.filter(b=>b.active).length,sourceError:range.getSheet().getRange(range.getRow(),9,1,18).getDisplayValues().flat().some(v=>v==='SOURCE MISSING'||/^#(REF!|VALUE!|N\/A|ERROR!)/.test(v))});});});
 const first=people[0],entry=registry[sstKey_(first.name)];if(entry){const rec=Object.values(entry.records)[0],range=sstNamed_(ss,rec.rangeName);ss.setActiveSheet(range.getSheet());range.getSheet().getRange(range.getRow(),1).activate();}SpreadsheetApp.flush();console.log(JSON.stringify({records:report,students:people.length,testRecords:Object.keys(registry).some(k=>k.includes('automation test')),sourceTabs:SST_ROUNDS.map(n=>({tab:n,backgroundHidden:ss.getSheetByName(n).isColumnHiddenByUser(2)}))}));
}


const SST_BUTTON_IMAGES={"controls": "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", "print": "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"};
// Read-only report view. Does not modify trackers, source goals, or registry.
function sstPrintReport(){
 const ss=sstBook_(),registry=sstRegistry_(),active=ss.getActiveSheet().getSheetId();
 const people=sstPeople_(ss).filter(p=>registry[sstKey_(p.name)]&&!registry[sstKey_(p.name)].archived);
 const list=people.map(p=>{const entry=registry[sstKey_(p.name)];return {name:p.name,selected:entry.sheetId===active,meetings:[1,2,3].filter(n=>entry.records[n]).map(n=>({n,state:entry.records[n].state}))};});
 if(!list.length){SpreadsheetApp.getUi().alert('No student progress sections exist yet. Enter goals in an SST tab first.');return;}
 const html=sstReportShell_(list);
 SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(1150).setHeight(800),'Print SST report');
}

function sstReportData(name,selected,options){
 const ss=sstBook_(),person=sstPeople_(ss).find(p=>sstKey_(p.name)===sstKey_(name));
 if(!person)throw Error('The student was not found in General.');
 const entry=sstRegistry_()[sstKey_(person.name)];if(!entry)throw Error('No Progress tab exists for this student.');
 const numbers=[...new Set((selected||[]).map(Number))].filter(n=>[1,2,3].includes(n)).sort();
 if(!numbers.length)throw Error('Check at least one SST meeting.');
 const opts={notes:options&&options.notes!==false,writing:!!(options&&options.writing),earlier:options&&options.earlier!==false,blankRows:[0,3,5].includes(Number(options&&options.blankRows))?Number(options.blankRows):0};
 const meetings=numbers.map(n=>{
  const rec=entry.records[n];if(!rec)throw Error('SST '+n+' does not have a Progress section yet.');
  const range=sstNamed_(ss,rec.rangeName),values=range.getDisplayValues();
  const blocks=(rec.blocks||[]).filter(b=>b.active||opts.earlier).map(b=>{
   const data=sstNamed_(ss,b.rangeName).getDisplayValues();
   const table=data.findIndex(v=>v[0]==='Date'&&v[1]==='Result / score');
   const end=data.findIndex((v,i)=>i>table&&v[0]==='Goal review / next step');
   if(table<0||end<0)throw Error('The progress table for '+b.area+' could not be located. Nothing was changed.');
   return {area:b.area,version:b.version||1,active:b.active,goal:data[1][1],plan:data[2][1],method:data[3][1],otherMethod:data[3][4],frequency:data[4][1],otherFrequency:data[4][4],baselineDate:data[5][1],baselineResult:data[5][3],review:data[end][1],checks:data.slice(table+1,end).filter(v=>v.some(x=>String(x).trim()))};
  }).filter(b=>String(b.goal).trim());
  const writing=[];if(opts.writing)values.forEach((v,i)=>{if(/^Writing evidence — /.test(v[0])&&values[i+2])writing.push({title:v[0].replace(/^Writing evidence — /,''),date:values[i+1]&&values[i+1][1]||'',comments:values[i+2][1]||''});});
  const changed=rec.state==='started'&&JSON.stringify(sstSourceValues_(ss,person,n))!==JSON.stringify(rec.snapshot);
  return {n,state:rec.state,date:values[1][1],reviewDate:values[1][3],notes:values[1][5],changed,blocks,writing};
 });
 return {name:person.name,generated:Utilities.formatDate(new Date(),ss.getSpreadsheetTimeZone(),'MM/dd/yyyy'),meetings,options:opts};
}

function sstReportRender_(report){
 const esc=x=>String(x==null?'':x).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
 const text=x=>esc(x).replace(/\n/g,'<br>');
 const cell=x=>'<td>'+text(x)+'</td>';
 const custom=(value,other)=>[value,other].filter(x=>String(x||'').trim()).join(' — ');
 let html='';
 report.meetings.forEach(meeting=>{
  const groups=[];for(let i=0;i<meeting.blocks.length;i+=2)groups.push(meeting.blocks.slice(i,i+2));
  if(!groups.length)groups.push([]);
  groups.forEach((goals,part)=>{
   html+='<section class="report-section"><header><h1>'+text(report.name)+'</h1><div class="meta">SST '+meeting.n+' on '+(meeting.date?text(meeting.date):'________________')+'</div></header>';
   if(goals.length){
    html+='<table class="goal-table"><colgroup><col style="width:17%">'+goals.map(()=>'<col style="width:'+(83/goals.length)+'%">').join('')+'</colgroup><thead><tr><th>Goal information</th>'+goals.map(g=>'<th>'+text(g.area)+(g.active&&g.version===1?'':'<span class="version">'+(g.active?'Revised':'Earlier')+' goal (version '+g.version+')</span>')+'</th>').join('')+'</tr></thead><tbody>';
    const fields=[['Student goal',g=>g.goal],['Intervention plan',g=>g.plan],['Assessment method',g=>custom(g.method,g.otherMethod)],['Frequency',g=>custom(g.frequency,g.otherFrequency)],['Baseline date',g=>g.baselineDate],['Baseline result',g=>g.baselineResult]];
    fields.forEach(([label,get])=>{if(label==='Student goal'||label==='Intervention plan'||goals.some(g=>String(get(g)||'').trim()))html+='<tr><th>'+label+'</th>'+goals.map(g=>cell(get(g))).join('')+'</tr>';});
    html+='<tr class="progress-band"><th colspan="'+(goals.length+1)+'">PROGRESS MONITORING RESULTS</th></tr>';
    const count=Math.max(...goals.map(g=>g.checks.length),0);
    for(let i=0;i<count;i++){
     html+='<tr class="check"><th>Recorded check '+(i+1)+'</th>'+goals.map(g=>{
      const v=g.checks[i];if(!v)return '<td class="unused"></td>';
      const labels=['Date','Result / score','Progress status','Assessment','Support','Observation','Custom status / other','Notes / evidence'];
      return '<td>'+v.map((value,j)=>String(value||'').trim()&&(j!==7||report.options.notes)?'<div class="check-line"><b>'+labels[j]+':</b> '+text(value)+'</div>':'').join('')+'</td>';
     }).join('')+'</tr>';
    }
    if(!count&&!report.options.blankRows)html+='<tr><th>Recorded progress</th><td colspan="'+goals.length+'" class="empty">No progress checks recorded yet.</td></tr>';
    for(let i=0;i<report.options.blankRows;i++)html+='<tr class="handwriting"><th>Additional check</th>'+goals.map(()=>'<td><span>Date: ____________________</span><br><br>Result / notes:</td>').join('')+'</tr>';
    if(report.options.notes&&goals.some(g=>String(g.review||'').trim()))html+='<tr><th>Goal review / next step</th>'+goals.map(g=>cell(g.review)).join('')+'</tr>';
    html+='</tbody></table>';
   }else html+='<p>No goals are available for this selection.</p>';
   if(part===groups.length-1){
    if(report.options.notes&&meeting.notes)html+='<div class="meeting-notes"><h2>Meeting notes</h2><p>'+text(meeting.notes)+'</p></div>';
    if(report.options.writing)meeting.writing.filter(w=>w.comments).forEach(w=>{html+='<div class="writing"><h2>Writing evidence — '+text(w.title)+'</h2>'+(w.date?'<p><b>Assessment date:</b> '+text(w.date)+'</p>':'')+'<p>'+text(w.comments)+'</p></div>';});
   }
   html+='</section>';
  });
 });
 return html;
}

function sstReportCss_(){return `
*{box-sizing:border-box}body{font-family:Arial,Helvetica,sans-serif;color:#233847;margin:0;background:#edf0f2;font-size:13px}h1{font-size:27px;margin:7px 0 9px}h2{font-size:14px;margin:0 0 8px}p{line-height:1.5}button,select{font:inherit}button{cursor:pointer;border:1px solid #254f68;border-radius:3px;padding:9px 17px;background:white;color:#254f68;font-weight:bold}button.primary{background:#254f68;color:white}button:disabled{opacity:.45;cursor:default}select{padding:7px;border:1px solid #a7b6c0;border-radius:3px}label{cursor:pointer}input{accent-color:#254f68}.controls{background:white;padding:18px 24px;border-bottom:1px solid #bcc9d0;position:sticky;top:0;z-index:2;box-shadow:0 2px 5px #0001}.controls h2{font-size:18px}.control-row{display:flex;align-items:center;gap:18px;flex-wrap:wrap;margin:12px 0}.control-row label{display:flex;align-items:center;gap:6px}.hint{color:#536674;font-size:12px;margin:6px 0}.error{color:#9a2828;white-space:pre-wrap}.report-section{background:white;width:100%;max-width:1080px;margin:20px auto;padding:30px 32px 20px;break-before:page}.report-section:first-child{break-before:auto}header{border-bottom:3px solid #254f68;padding-bottom:15px;margin-bottom:18px}.eyebrow{font-size:11px;font-weight:bold;letter-spacing:1.3px;color:#486577}.meta{font-size:13px;color:#465d6c}.goal-table{width:100%;border-collapse:collapse;table-layout:fixed;font-size:12px;line-height:1.45}.goal-table th,.goal-table td{border:1px solid #a8b6bf;padding:10px 12px;text-align:left;vertical-align:top;overflow-wrap:anywhere;white-space:normal}.goal-table thead th{background:#dbe8f4;color:#1f3d53;font-size:14px;border-bottom:2px solid #6f8ca3}.version{display:block;font-size:11px;font-weight:normal;margin-top:4px}.goal-table tbody th{background:#f0f5f8;font-size:12px}.goal-table .progress-band th{background:#dbece2;color:#245343;letter-spacing:.6px;font-size:12px;padding:10px 12px}.check-line{margin-bottom:5px}.check-line:last-child{margin:0}.goal-table .unused{background:#f6f7f8}.goal-table .empty{color:#5b6e7b;font-style:italic}.handwriting td{height:105px;color:#61727b}.footnote{font-size:10px;color:#61727b;margin:9px 0 0}.warning{background:#fff1d3;border-left:3px solid #bf8d35;padding:10px;font-size:12px}.meeting-notes,.writing{margin-top:22px;border-top:1px solid #b7c8d2;padding-top:13px}.writing p,.meeting-notes p{font-size:12px;line-height:1.5}footer{border-top:1px solid #ccd6dc;margin-top:24px;padding-top:9px;color:#5c6e7a;font-size:10px;display:flex;justify-content:space-between}tr{break-inside:avoid}thead{display:table-header-group}@page{size:letter landscape;margin:.45in}@media print{body{background:white}.controls{display:none!important}.report-section{margin:0;padding:0;max-width:none;width:auto;box-shadow:none}.goal-table{font-size:10.5pt}.goal-table th,.goal-table td{padding:8px 10px}.goal-table thead th{font-size:12pt}*{-webkit-print-color-adjust:exact;print-color-adjust:exact}}
header{border:0;padding-bottom:0;margin-bottom:18px}h1{font-size:24px;color:#000;letter-spacing:normal}.meta{font-size:14px;color:#000}.report-section,.goal-table thead th,.goal-table .progress-band th,.goal-table .empty{color:#000;letter-spacing:normal}.goal-table .empty{font-style:normal}
`;}

function sstReportShell_(students){
 const json=JSON.stringify(students).replace(/</g,'\\u003c');
 return '<!doctype html><html><head><meta charset="utf-8"><title>SST progress report</title><style>'+sstReportCss_()+'</style></head><body><div class="controls"><h2>Print SST report</h2><div class="control-row"><label><b>Student</b> <select id="student"></select></label><b>Include meetings:</b><span id="meetings"></span></div><div class="control-row"><label><input id="notes" type="checkbox" checked> Notes and next steps</label><label><input id="writing" type="checkbox"> Saved writing evidence</label><label><input id="earlier" type="checkbox" checked> Earlier goal versions and their progress</label></div><div class="control-row"><label>Extra rows for handwriting <select id="blank"><option value="0">None</option><option value="3">3 per goal</option><option value="5">5 per goal</option></select></label><button class="primary" id="preview">Preview report</button><button id="print" disabled>Print / Save PDF</button></div><p class="hint">Preview reads the saved tracker. It does not change any goals or results. Draft meetings are clearly labeled. For color printing, enable Background graphics in the print window.</p><p id="message" role="status" class="hint">Choose what to include, then click Preview report.</p></div><main id="report"></main><script>const students='+json+'; const renderReport='+sstReportRender_.toString()+';('+sstReportClient_.toString()+')();</script></body></html>';
}

function sstReportClient_(){
 const $=id=>document.getElementById(id),student=$('student'),message=$('message'),print=$('print');let generation=0;
 students.forEach(p=>{const option=document.createElement('option');option.textContent=p.name;option.value=p.name;option.selected=p.selected;student.appendChild(option);});
 function invalidate(){generation++;print.disabled=true;$('report').innerHTML='';message.textContent='Click Preview report to apply your choices.';}
 function meetings(){const p=students.find(p=>p.name===student.value),box=$('meetings');box.replaceChildren();p.meetings.forEach(m=>{const label=document.createElement('label'),check=document.createElement('input');check.type='checkbox';check.value=m.n;check.checked=m.state==='started';check.addEventListener('change',invalidate);label.style.display='inline-flex';label.style.marginRight='18px';label.append(check,document.createTextNode('SST '+m.n+(m.state==='started'?'':' (draft)')));box.appendChild(label);});invalidate();}
 student.addEventListener('change',meetings);['notes','writing','earlier','blank'].forEach(id=>$(id).addEventListener('change',invalidate));meetings();
 $('preview').addEventListener('click',()=>{
  const numbers=[...$('meetings').querySelectorAll('input:checked')].map(x=>Number(x.value));if(!numbers.length){message.textContent='Check at least one SST meeting above.';return;}
  const request=++generation;$('preview').disabled=true;print.disabled=true;message.className='hint';message.textContent='Preparing report…';
  const options={notes:$('notes').checked,writing:$('writing').checked,earlier:$('earlier').checked,blankRows:Number($('blank').value)};
  google.script.run.withSuccessHandler(report=>{$('preview').disabled=false;if(request!==generation)return;$('report').innerHTML=renderReport(report);document.title=report.name+' — SST progress report';print.disabled=false;const alerts=[];report.meetings.forEach(m=>{if(m.state!=='started')alerts.push('SST '+m.n+' is still a draft.');if(m.changed)alerts.push('SST '+m.n+': the source plan has changed. This report uses the saved goals and their recorded results. Review changes in the Progress tab before printing.');});message.textContent=alerts.length?alerts.join(' '):'Preview ready. Click Print / Save PDF when you are happy with it.';}).withFailureHandler(error=>{$('preview').disabled=false;if(request!==generation)return;message.className='error';message.textContent=error.message||String(error);}).sstReportData(student.value,numbers,options);
 });
 print.addEventListener('click',()=>{
  const output=window.open('','_blank');if(!output){message.textContent='Allow the report window to open, then click Print / Save PDF again.';return;}
  const css=document.querySelector('style').textContent;
  output.document.open();output.document.write('<!doctype html><html><head><meta charset="utf-8"><style>'+css+'</style></head><body><div class="controls"><button onclick="window.print()">Print / Save PDF</button></div><main>'+$('report').innerHTML+'</main></body></html>');output.document.close();output.document.title=document.title;output.focus();output.print();
 });
}

function sstActivePeople_(ss){const registry=sstRegistry_();return sstPeople_(ss).filter(p=>!(registry[sstKey_(p.name)]||{}).archived);}
function sstProgressBounds_(range){
 const data=range.getDisplayValues(),header=data.findIndex(v=>v[0]==='Date'&&v[1]==='Result / score');
 const review=data.findIndex((v,i)=>i>header&&v[0]==='Goal review / next step');
 if(header<0||review<=header+1)throw Error('The goal table could not be identified. No rows were changed.');
 return {first:range.getRow()+header+1,count:review-header-1,review:range.getRow()+review};
}
function sstManageData(){
 const ss=sstBook_(),registry=sstRegistry_(),active=ss.getActiveSheet().getSheetId();
 const students=sstPeople_(ss).map(p=>{
  const entry=registry[sstKey_(p.name)]||{records:{}};
  return {name:p.name,selected:entry.sheetId===active||(ss.getActiveSheet().getName()==='General'&&ss.getActiveSheet().getActiveRange().getRow()===p.row),archived:!!entry.archived,goals:entry.archived?[]:[1,2,3].flatMap(n=>{
   const rec=entry.records[n];if(!rec||rec.state!=='started')return [];
   return rec.blocks.map(b=>({id:b.rangeName,label:'SST '+n+' — '+b.area+(b.active?'':' — Earlier goal')+' (version '+b.version+')',rows:sstProgressBounds_(sstNamed_(ss,b.rangeName)).count}));
  })};
 });return {students};
}
function sstTrackerControls(){
 const data=sstManageData(),json=JSON.stringify(data).replace(/</g,'\\u003c');
 const html='<!doctype html><html><head><meta charset="utf-8"><style>body{font:14px Arial,sans-serif;color:#111;margin:24px}h2{font-size:18px;margin:0 0 16px}h3{font-size:15px;margin:22px 0 12px;border-top:1px solid #ccc;padding-top:18px}label{display:block;margin:12px 0 6px;font-weight:bold}select{font:14px Arial;padding:8px;width:100%;border:1px solid #9caab1;border-radius:3px}button{font:bold 14px Arial;padding:10px 16px;border:1px solid #385c6e;border-radius:3px;background:#e5eef5;color:#111;cursor:pointer;margin:12px 8px 0 0}button:disabled{opacity:.5;cursor:default}p{line-height:1.5}.help{font-size:13px;color:#444}.status{padding:10px;background:#fff2cc;min-height:20px;white-space:pre-wrap}</style></head><body><h2>Student and progress controls</h2><label for="student">Student</label><select id="student"></select><button id="studentActions">Edit name / archive / delete student</button><h3>Add progress rows</h3><label for="goal">Goal to add rows under</label><select id="goal"></select><label for="count">How many rows?</label><select id="count"><option value="5">5 more rows</option><option value="10">10 more rows</option><option value="20">20 more rows</option></select><button id="add">Add progress rows</button><p class="help">New rows keep the dropdowns, colors, borders, and date format. Existing results stay with their saved goal. Start the meeting first if no goals appear here.</p><h3>Archive or restore a student</h3><button id="archive">Archive selected student</button><button id="restore">Restore selected student</button><button id="deleteStudent">Delete student and data…</button><p class="help">Archive hides the Progress tab and removes the student from the active index and report list. General and SST source information remain saved. Select an archived name here to restore it. Nothing is deleted.</p><p id="status" class="status" role="status">Choose a student above.</p><script>let initial='+json+';('+sstManageClient_.toString()+')();</script></body></html>';
 SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(650).setHeight(750),'Tracker controls');
}
function sstManageClient_(){
 const $=id=>document.getElementById(id);let data=initial;
 function selected(){return data.students.find(p=>p.name===$('student').value);}
 function busy(value){document.querySelectorAll('button,select').forEach(el=>el.disabled=value);if(!value)choose();}
 function choose(){const p=selected();$('goal').replaceChildren();if(!p){$('add').disabled=$('archive').disabled=$('restore').disabled=true;return;}
  p.goals.forEach(g=>{const o=document.createElement('option');o.value=g.id;o.textContent=g.label+' — '+g.rows+' rows';$('goal').appendChild(o);});
  $('add').disabled=p.archived||!p.goals.length;$('archive').disabled=p.archived;$('restore').disabled=!p.archived;
 }
 function populate(keep){$('student').replaceChildren();data.students.forEach(p=>{const o=document.createElement('option');o.value=p.name;o.textContent=p.name+(p.archived?' (archived)':'');o.selected=keep?p.name===keep:p.selected;$('student').appendChild(o);});choose();}
 function run(action,argument){const p=selected();if(!p)return;const name=p.name;busy(true);$('status').textContent='Working…';
  google.script.run.withSuccessHandler(result=>{data=result.data;populate(name);busy(false);$('status').textContent=result.message;}).withFailureHandler(e=>{busy(false);$('status').textContent=e.message||String(e);}).sstManageAction(action,name,argument);
 }
 function openActions(mode){const p=selected();if(!p)return;google.script.run.withFailureHandler(e=>{$('status').textContent=e.message;}).sstShowStudentActions(p.name,mode||'');}
 $('studentActions').addEventListener('click',()=>openActions(''));$('deleteStudent').addEventListener('click',()=>openActions('delete'));
 $('student').addEventListener('change',()=>{choose();$('status').textContent=selected().archived?'This student is archived. Restore them to continue.':'Choose the goal and number of rows, or archive this student.';});
 $('add').addEventListener('click',()=>run('rows',{id:$('goal').value,count:Number($('count').value)}));
 $('archive').addEventListener('click',()=>{const p=selected();if(p&&window.confirm('Archive '+p.name+'? Their Progress tab will be hidden and all records will remain saved. You can restore them here.'))run('archive',{});});
 $('restore').addEventListener('click',()=>run('restore',{}));populate();
}
function sstManageAction(action,name,arg){
 const ss=sstBook_(),lock=LockService.getDocumentLock();lock.waitLock(30000);
 try{
  const p=sstPeople_(ss).find(p=>sstKey_(p.name)===sstKey_(name));if(!p)throw Error('This student is missing from General. Nothing was changed.');
  let registry=sstRegistry_(),key=sstKey_(p.name),entry=registry[key];
  if(action==='rows'){
   if(!entry||entry.archived)throw Error('Restore this student before adding rows.');
   const count=Number(arg&&arg.count);if(![5,10,20].includes(count))throw Error('Choose 5, 10, or 20 rows.');
   let record,block,n;for(const number of [1,2,3]){const rec=entry.records[number];if(!rec||rec.state!=='started')continue;const b=rec.blocks.find(b=>b.rangeName===(arg&&arg.id));if(b){record=rec;block=b;n=number;break;}}
   if(!block)throw Error('Choose a goal from a started meeting.');
   const range=sstNamed_(ss,block.rangeName),meeting=sstNamed_(ss,record.rangeName),sheet=range.getSheet(),bounds=sstProgressBounds_(range);
   const start=range.getRow(),length=range.getNumRows(),meetingStart=meeting.getRow(),meetingLength=meeting.getNumRows();
   if(bounds.count+count>1000)throw Error('This goal already has '+bounds.count+' rows. Please review the table before extending it past 1,000 rows.');
   const validation=sheet.getRange(bounds.first,1,1,8).getDataValidations()[0];
   sheet.insertRowsBefore(bounds.review,count);
   ss.setNamedRange(block.rangeName,sheet.getRange(start,1,length+count,8));
   ss.setNamedRange(record.rangeName,sheet.getRange(meetingStart,1,meetingLength+count,8));
   const dest=sheet.getRange(bounds.review,1,count,8);
   sheet.getRange(bounds.first,1,1,8).copyTo(dest,SpreadsheetApp.CopyPasteType.PASTE_FORMAT,false);
   dest.clearContent().setDataValidations(Array.from({length:count},()=>validation.slice())).setFontFamily('Arial').setFontSize(11).setWrap(true).setVerticalAlignment('top').setBorder(true,true,true,true,true,true,'#d5dfe4',SpreadsheetApp.BorderStyle.SOLID);
   sheet.getRange(bounds.review,1,count,2).setBackground('#ffe599');sheet.getRange(bounds.review,3,count,4).setBackground('#dbe8f4');sheet.getRange(bounds.review,7,count,2).setBackground('#ffe599');
   sheet.getRange(bounds.review,1,count,1).setNumberFormat('mm/dd/yy');sheet.getRange(bounds.review,2,count,1).setNumberFormat('@');sheet.setRowHeights(bounds.review,count,60);sheet.showRows(bounds.review,count);sstWarningProgress_(ss,sheet);
   const rules=sheet.getConditionalFormatRules();
   rules.unshift(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND(COUNTA($B'+bounds.review+':$H'+bounds.review+')>0,$A'+bounds.review+'="")').setBackground('#fce4e4').setRanges([sheet.getRange(bounds.review,1,count,1)]).build());
   rules.unshift(SpreadsheetApp.newConditionalFormatRule().whenFormulaSatisfied('=AND(COUNTIF($C'+bounds.review+':$F'+bounds.review+',"Other / my words")>0,COUNTA($G'+bounds.review+':$H'+bounds.review+')=0)').setBackground('#fff1bd').setRanges([sheet.getRange(bounds.review,7,count,2)]).build());sheet.setConditionalFormatRules(rules);
   ss.setActiveSheet(sheet);dest.getCell(1,1).activate();SpreadsheetApp.flush();
   return {message:'Added '+count+' rows for '+p.name+' — SST '+n+', '+block.area+'. This goal now has '+(bounds.count+count)+' progress rows. Existing results were kept.',data:sstManageData()};
  }
  if(!['archive','restore'].includes(action))throw Error('Unknown action. Nothing was changed.');
  if(!entry){entry={records:{},sheetId:null};registry[key]=entry;}
  const sheet=entry.sheetId?ss.getSheetById(entry.sheetId):null;
  if(action==='archive'){
   entry.archived=true;entry.archivedAt=new Date().toISOString();sstSaveRegistry_(registry);
   const index=ss.getSheetByName(SST_INDEX);ss.setActiveSheet(index);if(sheet&&!sheet.isSheetHidden())sheet.hideSheet();
  }else{
   entry.archived=false;delete entry.archivedAt;sstSaveRegistry_(registry);if(sheet)sheet.showSheet();
  }
  PropertiesService.getDocumentProperties().deleteProperty('SST_INDEX_ROSTER_V2');sstIndex_(ss,sstPeople_(ss));
  if(action==='restore'){sstReconcile_(ss);if(sheet){ss.setActiveSheet(sheet);sheet.getRange('A1').activate();}}
  return {message:p.name+(action==='archive'?' is archived. Their Progress tab is hidden. All source information, goals, and results remain saved. Use Restore selected student to bring them back.':' is active again. Their saved goals and results are available.'),data:sstManageData()};
 }finally{lock.releaseLock();}
}
function sstStyleProgress_(ss,sheet){
 const note='SHARED DROPDOWNS: Open Dropdown settings to add a choice for all student Progress trackers. Type it in the next blank gold box below the matching heading.\n\nQUICK HELP\n\nGold boxes: type dates, results, or your words. Blue boxes: choose dropdowns.\n\nClear a choice: Esc, then Backspace/Delete. A space is not blank. Undo a manual edit: Ctrl+Z (Mac: Command+Z).\n\nUse the SST1, SST2, or SST3 section you are working on. Draft goals update automatically. After starting, use Revise goal / plan for changes.\n\nNeed more rows? Click TRACKER CONTROLS at the top, choose the goal, then Add progress rows. You can add 5, 10, or 20 at a time.\n\nPRINT REPORT lets you choose which SST meetings to print. Archive and Restore are in TRACKER CONTROLS. Archive keeps all data and hides the Progress tab. Script actions are not reliably reversed by Ctrl+Z. Restore reverses archiving.';
 sheet.getRange('A1').setNote(note);sheet.getRange('A4:H5').breakApart();
 sheet.getRange('A4:D4').breakApart().merge().clearContent().setFontWeight('bold').setFontColor('#123c55').setBackground('#dbe8f4').setHorizontalAlignment('center').setNote('Click the TRACKER CONTROLS button to add progress rows or archive/restore a student.');
 sheet.getRange('E4:H4').breakApart().merge().clearContent().setFontWeight('bold').setFontColor('#123c55').setBackground('#dbece2').setHorizontalAlignment('center').setNote('Click the PRINT REPORT button to choose SST meetings, preview, then print or save PDF.');
 sheet.showRows(4);sheet.setRowHeight(4,40);sheet.hideRows(5);sheet.setFrozenRows(5);sstControlButtons_(sheet);
 const registry=sstRegistry_(),entry=Object.values(registry).find(e=>e.sheetId===sheet.getSheetId());if(!entry)return;
 const ids=entry.styledBlocks||[];
 Object.values(entry.records).forEach(rec=>{
  const meeting=sstNamed_(ss,rec.rangeName),r=meeting.getRow();
  sheet.getRange(r+1,1,1,8).setBorder(true,true,true,true,true,true,'#b8c4cd',SpreadsheetApp.BorderStyle.SOLID);[1,3,5].forEach(c=>sheet.getRange(r+1,c).setFontWeight('bold'));
  [2,4,6].forEach(c=>sheet.getRange(r+1,c,1,c===6?3:1).setBackground('#ffe599'));
  rec.blocks.forEach(b=>{if(ids.includes(b.id))return;const range=sstNamed_(ss,b.rangeName),start=range.getRow(),bounds=sstProgressBounds_(range);
   sheet.getRange(start+3,1,3,8).setBorder(true,true,true,true,true,true,'#b8c4cd',SpreadsheetApp.BorderStyle.SOLID);
   [start+3,start+4].forEach(row=>{sheet.getRange(row,1).setFontWeight('bold');sheet.getRange(row,4).setFontWeight('bold');sheet.getRange(row,2,1,2).setBackground('#dbe8f4');sheet.getRange(row,5,1,4).setBackground('#ffe599');});
   sheet.getRange(start+5,1).setFontWeight('bold');sheet.getRange(start+5,3).setFontWeight('bold');sheet.getRange(start+5,2).setBackground('#ffe599');sheet.getRange(start+5,4,1,5).setBackground('#ffe599');
   sheet.getRange(bounds.first,1,bounds.count,2).setBackground('#ffe599');sheet.getRange(bounds.first,3,bounds.count,4).setBackground('#dbe8f4');sheet.getRange(bounds.first,7,bounds.count,2).setBackground('#ffe599');
   sheet.getRange(bounds.review,1).setFontWeight('bold');sheet.getRange(bounds.review,2,1,7).setBackground('#ffe599').setBorder(true,true,true,true,true,true,'#b8c4cd',SpreadsheetApp.BorderStyle.SOLID);ids.push(b.id);
  });
 });entry.styledBlocks=ids;sstSaveRegistry_(registry);sstWarningProgress_(ss,sheet);sstMaybeRefreshNativeTabs_(ss);
}
function sstControlButtons_(sheet){
 const images=sheet.getImages();
 [['SST_TRACKER_CONTROLS','controls',1,'sstTrackerControls'],['SST_PRINT_REPORT','print',5,'sstPrintReport']].forEach(([title,key,column,fn])=>{
  if(images.some(image=>image.getAltTextTitle()===title))return;
  sheet.insertImage(Utilities.newBlob(Utilities.base64Decode(SST_BUTTON_IMAGES[key]),'image/png',title+'.png'),column,4,6,4).setWidth(270).setHeight(32).setAltTextTitle(title).setAltTextDescription(title==='SST_PRINT_REPORT'?'Print report: choose SST meetings':'Add progress rows, archive or restore student').assignScript(fn);
 });
}
function sstInstallControls(){
 const ss=sstBook_(),registry=sstRegistry_();Object.values(registry).forEach(entry=>{if(entry.sheetId){const sheet=ss.getSheetById(entry.sheetId);if(sheet)sstStyleProgress_(ss,sheet);}});
 const q=ss.getSheetByName('Q&A - Start Here');if(q){const updates={
  'What if I run out of progress rows?':'Click TRACKER CONTROLS at the top of the student’s Progress tab (or SST Progress > Tracker controls). Choose the student, choose the goal in the correct SST meeting, choose 5, 10, or 20 rows, and click Add progress rows. It adds blank rows with dropdowns, colors, borders, and date formats. Existing results remain with their saved goal. You can repeat this whenever needed. Rows do not add themselves automatically.',
  'Where can I see reminders without opening Q&A?':'Hover over the student’s name in the dark title bar. Quick Help appears as a note. Use TRACKER CONTROLS for more rows, Archive, or Restore. Use PRINT REPORT to choose the meetings to print.'
 };const values=q.getDataRange().getDisplayValues();values.forEach((row,i)=>{if(updates[row[1]]){q.getRange(i+1,3).setValue(updates[row[1]]);q.getRange(i+1,4).setValue('Works now');}});
 const extra=[['Student management','What should I do when a student moves away?','Use TRACKER CONTROLS > choose student > Archive selected student. Confirm the name. Their Progress tab is hidden and they leave the active index and report list. General and SST source rows stay saved. Nothing is erased.','Works now'],['Student management','How do I bring an archived student back?','Open SST Progress > Tracker controls from any tab. Select the name marked (archived), then click Restore selected student. Their existing tracker and records return.','Works now'],['Student management','Should I delete their name from General?','No. Archive the student instead. The name connects their source information to their saved records. Removing or changing it manually can break those connections.','Works now'],['Printing','How do I print one or more SST meetings?','Click PRINT REPORT on a student’s Progress tab, or SST Progress > Print report. Choose the student, check the SST meetings, choose optional notes, writing evidence, or handwriting rows, then Preview report and Print / Save PDF. Printed headings show the student name and SST number with its meeting date.','Works now'],['Adding students','Will new students have these same controls?','Yes. Add a unique student name in General. Enter goals in the appropriate SST tab. Their Progress tab is created automatically with the same dropdowns, entry boxes, hover help, row controls, and report controls. The meeting must be started before dated progress rows can be extended.','Works now'],['Undo','Can Ctrl+Z undo a button action?','Do not rely on Ctrl+Z for script actions. Archive has its own Restore control. Added blank progress rows can stay unused and do not appear in the report unless you choose handwriting rows. Existing goals and results are not erased by these controls.','Works now']];
 extra.forEach(row=>{const found=values.findIndex(v=>v[1]===row[1]);const target=found>=0?found+1:q.getLastRow()+1;q.getRange(target,1,1,4).setValues([row]).setFontFamily('Arial').setFontSize(11).setWrap(true).setVerticalAlignment('top');q.setRowHeight(target,100);});
 }
 PropertiesService.getDocumentProperties().setProperty('SST_CONTROLS_VERSION','1');console.log('Installed tracker and print controls for '+Object.values(registry).filter(e=>e.sheetId).length+' student tabs.');
}

const SST_SHARED_OPTIONS='Dropdown settings';
const SST_SHARED_GROUPS=[
 {key:'status',title:'Progress status',col:1,seed:['Not yet met','Making progress','Met this check','Goal met — criteria checked','Not assessed','Absent','Other / my words']},
 {key:'method',title:'Assessment type',col:2,seed:['Word reading check','Spelling dictation','Writing sample / rubric','Math quiz','Work sample','Reading passage','Observation','Other / my words']},
 {key:'frequency',title:'Frequency',col:3,seed:['Daily','Twice weekly','Weekly','Every two weeks','As scheduled','Other / my words']},
 {key:'support',title:'Support provided',col:4,seed:['Independent','Verbal prompt','Visual support','Model / example','Guided practice','One-to-one support','Other / my words']},
 {key:'observation',title:'Observation',col:5,seed:['More accurate','More fluent','More independent','Needs reteaching','Inconsistent performance','No clear change yet','Other / my words']}
];
var SST_SHARED_RULE_CACHE_={};
function sstSharedRule_(options){
 const group=SST_SHARED_GROUPS.find(g=>g.seed.length===options.length&&g.seed.every((v,i)=>options[i]===v));
 if(group){
  if(SST_SHARED_RULE_CACHE_[group.key])return SST_SHARED_RULE_CACHE_[group.key];
  const sheet=sstBook_().getSheetByName(SST_SHARED_OPTIONS);
  if(sheet){const rule=SpreadsheetApp.newDataValidation().requireValueInRange(sheet.getRange(5,group.col,1000,1),true).setAllowInvalid(false).build();SST_SHARED_RULE_CACHE_[group.key]=rule;return rule;}
 }
 return SpreadsheetApp.newDataValidation().requireValueInList(options,true).setAllowInvalid(false).build();
}
function sstSharedAddUnique_(list,values){values.forEach(v=>{const text=String(v==null?'':v);if(text.trim()&&!list.includes(text))list.push(text);});}
function sstSharedHarvest_(ss){
 const lists={};SST_SHARED_GROUPS.forEach(g=>lists[g.key]=g.seed.slice());
 const settings=ss.getSheetByName(SST_SHARED_OPTIONS);
 if(settings)SST_SHARED_GROUPS.forEach(g=>sstSharedAddUnique_(lists[g.key],settings.getRange(5,g.col,Math.max(1,settings.getLastRow()-4),1).getDisplayValues().flat()));
 const seenRanges={};function collect(key,range){
  sstSharedAddUnique_(lists[key],range.getDisplayValues().flat());
  range.getDataValidations().flat().forEach(rule=>{if(!rule)return;const type=rule.getCriteriaType(),args=rule.getCriteriaValues();if(type===SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST)sstSharedAddUnique_(lists[key],args[0]);else if(type===SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE){const id=key+':'+args[0].getSheet().getSheetId()+':'+args[0].getA1Notation();if(!seenRanges[id]){seenRanges[id]=true;sstSharedAddUnique_(lists[key],args[0].getDisplayValues().flat());}}});
 }
 Object.values(sstRegistry_()).forEach(entry=>Object.values(entry.records||{}).forEach(rec=>(rec.blocks||[]).forEach(b=>{
  const range=sstNamed_(ss,b.rangeName),sheet=range.getSheet(),row=range.getRow(),bounds=sstProgressBounds_(range);
  collect('method',sheet.getRange(row+3,2));collect('frequency',sheet.getRange(row+4,2));
  [['status',3],['method',4],['support',5],['observation',6]].forEach(([key,col])=>collect(key,sheet.getRange(bounds.first,col,bounds.count,1)));
 })));
 sstSharedAddUnique_(lists.support,['Small group instruction']);return lists;
}
function sstApplySharedRules_(ss){
 let blocks=0;
 Object.values(sstRegistry_()).forEach(entry=>Object.values(entry.records||{}).forEach(rec=>(rec.blocks||[]).forEach(b=>{
  const range=sstNamed_(ss,b.rangeName),sheet=range.getSheet(),row=range.getRow(),bounds=sstProgressBounds_(range);
  function rule(key){return sstSharedRule_(SST_SHARED_GROUPS.find(g=>g.key===key).seed);}
  sheet.getRange(row+3,2,1,2).setDataValidation(rule('method'));sheet.getRange(row+4,2,1,2).setDataValidation(rule('frequency'));
  [['status',3],['method',4],['support',5],['observation',6]].forEach(([key,col])=>sheet.getRange(bounds.first,col,bounds.count,1).setDataValidation(rule(key)));
  blocks++;
 })));return blocks;
}
function sstOpenDropdownSettings(){const ss=sstBook_(),sheet=ss.getSheetByName(SST_SHARED_OPTIONS);if(!sheet)throw Error('Shared dropdown settings have not been installed.');ss.setActiveSheet(sheet);sheet.getRange('A5').activate();}
function sstInstallSharedDropdowns(){
 const ss=sstBook_(),lock=LockService.getDocumentLock();lock.waitLock(30000);
 try{
  const before={};Object.values(sstRegistry_()).forEach(entry=>Object.values(entry.records||{}).forEach(rec=>(rec.blocks||[]).forEach(b=>before[b.rangeName]=JSON.stringify(sstNamed_(ss,b.rangeName).getDisplayValues()))));
  const lists=sstSharedHarvest_(ss);let sheet=ss.getSheetByName(SST_SHARED_OPTIONS);
  if(!sheet)sheet=ss.insertSheet(SST_SHARED_OPTIONS);
  if(sheet.getMaxRows()<1004)sheet.insertRowsAfter(sheet.getMaxRows(),1004-sheet.getMaxRows());
  sheet.getRange('A1:E1').breakApart().merge().setValue('Shared dropdown choices').setFontFamily('Arial').setFontSize(18).setFontWeight('bold').setFontColor('white').setBackground('#223f53');sheet.setRowHeight(1,40);
  sheet.getRange('A2:E2').breakApart().merge().setValue('To add a choice everywhere: find the matching heading below, click its next empty gold box, type your choice, and press Enter. It is available in all student Progress trackers.').setWrap(true).setFontSize(11).setBackground('#e9f3ef');sheet.setRowHeight(2,45);
  sheet.getRange('A3:E3').breakApart().merge().setValue('Add new choices in blank cells. Keep existing choices so old results stay valid. For wording you need only once, choose Other / my words in the progress row and use the custom box or notes.').setWrap(true).setFontSize(11);sheet.setRowHeight(3,45);
  sheet.getRange('A4:E4').setValues([SST_SHARED_GROUPS.map(g=>g.title)]).setFontWeight('bold').setFontSize(11).setBackground('#dbe8f4').setWrap(true);sheet.setRowHeight(4,32);
  sheet.getRange(5,1,1000,5).setFontFamily('Arial').setFontSize(11).setBackground('#ffe599').setNumberFormat('@').setWrap(true).setVerticalAlignment('top');sheet.setRowHeights(5,50,30);
  SST_SHARED_GROUPS.forEach(g=>{
   // On first install merge existing local options into this shared column.
   // Later runs leave the teacher's ordering and blank rows untouched.
   const existing=sheet.getRange(5,g.col,1000,1).getDisplayValues().flat();
   const missing=lists[g.key].filter(v=>!existing.includes(v));let last=existing.reduce((n,v,i)=>v.trim()?i+1:n,0);
   if(last+missing.length>1000)throw Error('The '+g.title+' shared list is full. Nothing was removed.');
   if(missing.length)sheet.getRange(5+last,g.col,missing.length,1).setValues(missing.map(v=>[v]));
   sheet.getRange(4,g.col).setNote('Add a new '+g.title.toLowerCase()+' choice in the next empty gold box in this column. All current and future Progress trackers share this list. Keep existing choices that have already been used.');sheet.setColumnWidth(g.col,220);
  });
  sheet.setFrozenRows(4);sheet.setHiddenGridlines(true);sheet.setTabColor('#bfa04b');sheet.getRange('A1:E1004').setFontFamily('Arial');
  SST_SHARED_RULE_CACHE_={};const count=sstApplySharedRules_(ss);
  Object.entries(before).forEach(([name,values])=>{if(JSON.stringify(sstNamed_(ss,name).getDisplayValues())!==values)throw Error('A tracker value changed during the dropdown update. Please review '+name);});
  const registry=sstRegistry_();Object.values(registry).forEach(entry=>{if(!entry.sheetId)return;const progress=ss.getSheetById(entry.sheetId);if(progress){const cell=progress.getRange('A1'),note=cell.getNote();if(!note.includes('SHARED DROPDOWNS'))cell.setNote(note+'\n\nSHARED DROPDOWNS: Open Dropdown settings to add a choice for every student. Type it in the next blank gold box under the matching heading.');}});
  const q=ss.getSheetByName('Q&A - Start Here');if(q){const old=q.getDataRange().getDisplayValues();const questions=[
   ['Dropdown choices','How do I add a dropdown choice for every student?','Open Dropdown settings. Find the correct heading: Progress status, Assessment type, Frequency, Support provided, or Observation. Click the next empty gold box below that list, type the new choice, and press Enter. It becomes available throughout current Progress trackers, added progress rows, and future student trackers. No refresh or code edit is needed.','Works now'],
   ['Dropdown choices','Can I add Small group instruction everywhere?','Yes. Small group instruction is already included under Support provided in Dropdown settings. Open the Support provided dropdown in any progress row to select it.','Works now'],
   ['Dropdown choices','What if I need a choice for only one progress check?','Choose Other / my words in that row and type the detail in Custom status / other or My notes / evidence. Do not add it to Dropdown settings unless you want it available to every student.','Works now'],
   ['Dropdown choices','Should I use the pencil to change all the dropdowns?','Use Dropdown settings for a shared change. The pencil edits a cell’s validation rule and may affect only that cell. The shared lists now provide the same choices to all Progress trackers.','Works now'],
   ['Dropdown choices','Can I rename or remove an old shared choice?','Adding a new choice in a blank gold box is the safest change. Renaming or removing a choice does not rewrite past results and can make an old selection show a validation warning. Keep choices already used in records.','Works now'],
   ['Dropdown choices','Which dropdowns use these shared lists?','These lists control the Progress trackers: Progress status; Assessment type (including the Assessment field above each goal); Frequency; Support provided; and Observation. General, SST source dropdowns, and Writing/Escritura are not changed by these settings.','Works now']
  ];questions.forEach(row=>{const i=old.findIndex(v=>v[1]===row[1]),r=i>=0?i+1:q.getLastRow()+1;q.getRange(r,1,1,4).setValues([row]).setFontFamily('Arial').setFontSize(11).setWrap(true).setVerticalAlignment('top');q.setRowHeight(r,100);});}
  PropertiesService.getDocumentProperties().setProperty('SST_SHARED_DROPDOWNS_VERSION','1');
  ss.setActiveSheet(sheet);ss.moveActiveSheet(3);sheet.getRange('D5').activate();SpreadsheetApp.flush();console.log('Shared dropdowns connected to '+count+' goal blocks. Existing values and local choices were preserved.');
 }finally{lock.releaseLock();}
}
function sstCheckSharedDropdowns(){
 const ss=sstBook_(),checks=[],settings=ss.getSheetByName(SST_SHARED_OPTIONS);if(!settings)throw Error('Dropdown settings is missing.');
 function verify(range,col,label){const rule=range.getDataValidation();if(!rule||rule.getCriteriaType()!==SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE)throw Error(label+' is not linked to a shared range.');const source=rule.getCriteriaValues()[0];if(source.getSheet().getSheetId()!==settings.getSheetId()||source.getColumn()!==col||source.getRow()!==5||source.getNumRows()!==1000)throw Error(label+' has the wrong shared range.');checks.push(label);}
 Object.entries(sstRegistry_()).forEach(([name,entry])=>Object.entries(entry.records||{}).forEach(([n,rec])=>rec.blocks.forEach(b=>{const range=sstNamed_(ss,b.rangeName),sheet=range.getSheet(),r=range.getRow(),p=sstProgressBounds_(range);verify(sheet.getRange(r+3,2),2,name+' SST '+n+' '+b.area+' assessment');verify(sheet.getRange(r+4,2),3,name+' SST '+n+' '+b.area+' frequency');[[3,1],[4,2],[5,4],[6,5]].forEach(([c,shared])=>{verify(sheet.getRange(p.first,c),shared,name+' '+b.area+' first '+c);verify(sheet.getRange(p.first+p.count-1,c),shared,name+' '+b.area+' last '+c);});})));
 SST_SHARED_GROUPS.forEach(g=>{const rule=sstRule_(g.seed);if(rule.getCriteriaType()!==SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE||rule.getCriteriaValues()[0].getColumn()!==g.col)throw Error('Future '+g.title+' rule is not shared.');});
 console.log(JSON.stringify({linkedCellsChecked:checks.length,futureCategoriesChecked:5,smallGroupAvailable:settings.getRange('D5:D1004').getDisplayValues().flat().includes('Small group instruction'),sample:checks.slice(0,4)}));
}

function sstMeetingStatus_(p,entry){
 entry=entry||{};const previous=Number(p.values[3])||0,suggested=Math.max(1,Math.min(3,previous+1));
 const round=[1,2,3].includes(Number(entry.meetingChoice))?Number(entry.meetingChoice):suggested,check=entry.meetingCheck;
 const checked=!!(check&&check.round===round&&check.previousCount===previous);
 return {previous,suggested,round,checked,needsReview:previous>0||entry.meetingChoice!=null};
}
function sstMeetingAlerts_(p,entry){
 const status=sstMeetingStatus_(p,entry),alerts=[];
 if(status.needsReview&&!status.checked)alerts.push('Check school records for SST '+status.round+', then tick Meeting number checked.');
 if(status.checked&&entry&&Object.keys(entry.records||{}).length&&!entry.records[status.round]){
  alerts.push('Confirmed SST '+status.round+' has no saved goals yet. Enter goals in the SST '+status.round+' tab. Existing SST '+Object.keys(entry.records).join(', SST ')+' records were kept.');
 }return alerts;
}
function sstMeetingHeader_(ss){
 const s=ss.getSheetByName(SST_INDEX);if(!s)return;
 s.getRange('I12:J12').setValues([['Meeting number','Meeting number checked']]).setFontFamily('Arial').setFontSize(11).setFontWeight('bold').setFontColor('white').setBackground(SST_COLORS.green).setWrap(true).setVerticalAlignment('top');
 s.showColumns(9,2);[135,100,125,115,90,90,90,240,120,130].forEach((w,i)=>s.setColumnWidth(i+1,w));
 s.getRange('I12').setNote('Choose the actual SST number shown in school records. This can differ from the suggestion in column C. Changing this choice does not move or rename goals or progress.');
 s.getRange('J12').setNote('Tick only after you check the school records. Changing the meeting number or the previous-year SST count clears the confirmation. Other warnings stay until resolved.');
}
function sstMeetingCells_(ss,p,row,entry){
 const s=ss.getSheetByName(SST_INDEX),status=sstMeetingStatus_(p,entry);
 s.getRange(row,9).setDataValidation(sstRule_(['SST 1','SST 2','SST 3'])).setValue('SST '+status.round).setFontFamily('Arial').setFontSize(11).setBackground('#dbe8f4').setVerticalAlignment('top');
 s.getRange(row,10).setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build()).setValue(status.checked).setBackground(status.checked?'#dbece2':'#ffe599').setHorizontalAlignment('center').setVerticalAlignment('top');
 s.getRange(row,10).setNote(status.checked?'You checked SST '+status.round+'. To correct a mistake, change Meeting number and check again. This does not move saved goals or progress.':'After checking school records, choose SST1, SST2, or SST3 in Meeting number, then tick this box.');
 return status;
}
function sstMeetingEdit_(ss,range){
 const s=range.getSheet(),people=sstActivePeople_(ss),choiceEdited=range.getColumn()<=9&&range.getLastColumn()>=9,checkEdited=range.getColumn()<=10&&range.getLastColumn()>=10;
 const registry=sstRegistry_();const updated=[];
 for(let row=Math.max(13,range.getRow());row<=range.getLastRow();row++){
  const key=sstKey_(s.getRange(row,1).getDisplayValue()),p=people.find(p=>sstKey_(p.name)===key);if(!p)continue;
  const entry=registry[key]||(registry[key]={sheetId:null,records:{}}),old=sstMeetingStatus_(p,entry);
  const number=Number(s.getRange(row,9).getDisplayValue().replace(/\D/g,''));if(![1,2,3].includes(number)){sstMeetingCells_(ss,p,row,entry);ss.toast('Choose SST 1, SST 2, or SST 3. The last meeting choice was restored.','Student Index',7);continue;}
  if(choiceEdited){entry.meetingChoice=number;if(number!==old.round)delete entry.meetingCheck;}
  if(checkEdited){entry.meetingChoice=number;if(s.getRange(row,10).getValue()===true)entry.meetingCheck={round:number,previousCount:Number(p.values[3])||0,checkedAt:new Date().toISOString()};else delete entry.meetingCheck;}
  updated.push(p);
 }
 sstSaveRegistry_(registry);sstMeetingHeader_(ss);updated.forEach(p=>sstIndexRow_(ss,p));
 if(choiceEdited&&!checkEdited)ss.toast('Meeting choice updated. Check the school records and tick Meeting number checked. Saved goals and progress were not moved.','Student Index',8);
}
function sstInstallMeetingChecks(){
 const ss=sstBook_(),index=ss.getSheetByName(SST_INDEX),people=sstActivePeople_(ss),registry=sstRegistry_();
 const before={};Object.values(registry).forEach(entry=>Object.values(entry.records||{}).forEach(rec=>before[rec.rangeName]=JSON.stringify(sstNamed_(ss,rec.rangeName).getDisplayValues())));
 sstMeetingHeader_(ss);people.forEach(p=>sstIndexRow_(ss,p));
 const q=ss.getSheetByName('Q&A - Start Here');if(q){const old=q.getDataRange().getDisplayValues(),rows=[
  ['Meeting number','How do I mark the school-records reminder as done?','On Student Index, find the student. Beside Notes, choose the actual SST1, SST2, or SST3 in Meeting number (column I). After checking school records, tick Meeting number checked (column J). The school-records reminder clears. Other warnings remain until their cause is resolved.','Works now'],
  ['Meeting number','What if I picked SST2 but it should be SST1 or SST3?','Change Meeting number in column I to the correct SST. The checkmark clears automatically. Check the school records, then tick column J again. You can also uncheck column J if you marked it done by mistake.','Works now'],
  ['Meeting number','Does changing the confirmed meeting move my goals or progress?','No. These controls record which meeting number you verified. They do not rename, move, or delete saved SST sections, goals, intervention plans, or results. If the chosen SST has no goals while another SST already has records, Notes explains that those records were kept. Enter new goals in the appropriate SST tab.','Works now'],
  ['Meeting number','Can the confirmed number differ from the suggested number?','Yes. The suggested number in column C comes from the previous-year SST count. The number you choose in column I comes from your school-records check and can be SST1, SST2, or SST3. Keep the historical count accurate. Changing that count requires confirmation again.','Works now'],
  ['Meeting number','Will checking the box hide every warning?','No. It clears only the reminder to check the meeting number. Missing plans, references to another student, and changed saved plans still show their own warnings.','Works now']
 ];rows.forEach(row=>{const i=old.findIndex(v=>v[1]===row[1]),r=i<0?q.getLastRow()+1:i+1;q.getRange(r,1,1,4).setValues([row]).setFontFamily('Arial').setFontSize(11).setWrap(true).setVerticalAlignment('top');q.setRowHeight(r,100);});}
 Object.entries(before).forEach(([name,data])=>{if(JSON.stringify(sstNamed_(ss,name).getDisplayValues())!==data)throw Error('Saved meeting values changed unexpectedly. Review '+name);});
 ss.setActiveSheet(index);index.getRange('I13').activate();SpreadsheetApp.flush();console.log('Meeting-number controls installed for '+people.length+' students. Saved goals and progress are unchanged. No school-records checks were marked complete.');
}
function sstTestMeetingChecks(){
 const tests=[];function ok(v,label){if(!v)throw Error(label);tests.push(label);}
 const p={values:['','','',1]},entry={records:{2:{}}};
 ok(sstMeetingStatus_(p,entry).round===2&&!sstMeetingStatus_(p,entry).checked,'Prior SST1 suggests SST2 without marking checked');
 entry.meetingChoice=2;entry.meetingCheck={round:2,previousCount:1};ok(sstMeetingStatus_(p,entry).checked&&sstMeetingAlerts_(p,entry).length===0,'Valid check clears only the school-records reminder');
 entry.meetingChoice=1;ok(!sstMeetingStatus_(p,entry).checked&&sstMeetingAlerts_(p,entry)[0].includes('Check school records'),'Changing SST2 to SST1 invalidates the old check');
 entry.meetingCheck={round:1,previousCount:1};ok(sstMeetingStatus_(p,entry).checked&&sstMeetingAlerts_(p,entry).some(v=>v.includes('records were kept')),'Confirmed different number preserves existing SST2 and warns about missing SST1 goals');
 entry.meetingChoice=3;ok(!sstMeetingStatus_(p,entry).checked,'Changing SST1 to SST3 requires another check');
 entry.meetingCheck={round:3,previousCount:1};ok(sstMeetingStatus_(p,entry).checked,'Corrected SST3 can be confirmed');
 p.values[3]=2;ok(!sstMeetingStatus_(p,entry).checked,'Changing the prior-year count requires confirmation again');
 delete entry.meetingCheck;ok(!sstMeetingStatus_(p,entry).checked,'Unchecking restores the reminder');
 const ss=sstBook_(),people=sstActivePeople_(ss),index=ss.getSheetByName(SST_INDEX);people.forEach((person,i)=>{const status=sstMeetingStatus_(person,sstRegistry_()[sstKey_(person.name)]),r=13+i;ok(index.getRange(r,9).getDisplayValue()==='SST '+status.round&&index.getRange(r,10).getValue()===status.checked,'Index controls match saved state for '+person.name);});
 console.log(JSON.stringify({checks:tests,students:people.length,schoolRecordsNotAssumedVerified:true}));
}

function sstProgressTitle_(name){return ('Progress - '+String(name).replace(/[\\/:?*\[\]]/g,' ')).slice(0,100);}
function sstNameAtRow_(ss,row,registry){
 return Object.entries(registry).find(([key,entry])=>{if(!entry.generalNameRange)return false;const r=ss.getRangeByName(entry.generalNameRange);return r&&r.getSheet().getName()==='General'&&r.getRow()===row;});
}
function sstBindStudentIds_(ss,people){
 const registry=sstRegistry_();people.forEach(p=>{
  const key=sstKey_(p.name),entry=registry[key]||(registry[key]={sheetId:null,records:{}});
  if(!entry.studentId)entry.studentId=Utilities.getUuid();entry.name=p.name;
  if(!entry.generalNameRange)entry.generalNameRange='SSTSTUDENT_'+entry.studentId.replace(/-/g,'');
  const range=ss.getRangeByName(entry.generalNameRange);if(!range||range.getRow()!==p.row)ss.setNamedRange(entry.generalNameRange,ss.getSheetByName('General').getRange(p.row,1));
 });sstSaveRegistry_(registry);
}
function sstRenameIndex_(ss,oldName,newName){
 const index=ss.getSheetByName(SST_INDEX);if(!index)return;
 const people=sstActivePeople_(ss),selected=index.getRange('B3').getDisplayValue();
 if(people.length)index.getRange(13,1,people.length,1).setValues(people.map(p=>[p.name]));
 index.getRange('B3').clearDataValidations();if(people.length)index.getRange('B3').setDataValidation(sstRule_(people.map(p=>p.name)));
 if(sstKey_(selected)===sstKey_(oldName))index.getRange('B3').setValue(newName);
 PropertiesService.getDocumentProperties().setProperty('SST_INDEX_ROSTER_V2',JSON.stringify(people.map(p=>[p.name,p.values[3]])));
 const person=sstPeople_(ss).find(p=>sstKey_(p.name)===sstKey_(newName));if(person)sstIndexRow_(ss,person);
}
function sstPrepareNameCorrection_(ss,oldName,newName,row,entry){
 const registry=sstRegistry_(),oldKey=sstKey_(oldName),newKey=sstKey_(newName),general=ss.getSheetByName('General');
 if(!newName)throw Error('The name cannot be blank. Archive the student instead of erasing their name.');
 if(general.getRange(row,1).getFormula())throw Error('Type the student name as text, not as a formula.');
 const matches=general.getRange(3,1,general.getMaxRows()-2,1).getDisplayValues().flat().filter(v=>sstKey_(v)===newKey);
 if(matches.length!==1)throw Error('That name is already used by another student. Use a unique name; no records were merged.');
 if(registry[newKey]&&registry[newKey].studentId!==entry.studentId&&newKey!==oldKey)throw Error('Another saved student already uses that name. No records were merged.');
 const title=sstProgressTitle_(newName),conflict=ss.getSheets().find(s=>s.getName().toLowerCase()===title.toLowerCase()&&s.getSheetId()!==entry.sheetId);
 if(conflict)throw Error('Another Progress tab already uses that name. Nothing was overwritten.');
 const sources=SST_ROUNDS.map((title,n)=>{
  const s=ss.getSheetByName(title),values=s.getRange(3,1,s.getMaxRows()-2,1).getDisplayValues().flat(),oldRows=[],newRows=[];
  values.forEach((v,i)=>{if(sstKey_(v)===oldKey)oldRows.push(i+3);if(sstKey_(v)===newKey)newRows.push(i+3);});
  if(oldRows.length>1||(oldKey!==newKey&&newRows.length))throw Error('Check duplicate names in '+title+'. The correction was stopped.');
  if(!oldRows.length)throw Error('The old name is missing from '+title+'. No student records were moved.');
  return {title,row:oldRows[0]};
 });
 return {studentId:entry.studentId,oldName,newName,row,sources,names:[oldName,newName],createdAt:new Date().toISOString()};
}
function sstFinishNameCorrection_(ss,job){
 const props=PropertiesService.getDocumentProperties(),general=ss.getSheetByName('General'),registry=sstRegistry_();
 const pair=Object.entries(registry).find(([key,e])=>e.studentId===job.studentId);if(!pair)throw Error('The saved student identity is missing. No progress was deleted.');
 const [currentKey,entry]=pair,current=String(general.getRange(job.row,1).getDisplayValue()).trim().replace(/\s+/g,' ');
 if(!current)throw Error('The corrected name is blank. Enter the student name to finish the correction.');
 if(!job.names.includes(current))job.names.push(current);job.newName=current;
 const newKey=sstKey_(current),count=general.getRange(3,1,general.getMaxRows()-2,1).getDisplayValues().flat().filter(v=>sstKey_(v)===newKey).length;
 if(count!==1||(registry[newKey]&&registry[newKey].studentId!==job.studentId))throw Error('The name correction conflicts with another student. No records were merged.');
 const progress=entry.sheetId?ss.getSheetById(entry.sheetId):null,title=sstProgressTitle_(current);
 if(progress&&ss.getSheets().some(s=>s.getName().toLowerCase()===title.toLowerCase()&&s.getSheetId()!==progress.getSheetId()))throw Error('A different Progress tab already has that name.');
 job.sources.forEach(source=>{const value=ss.getSheetByName(source.title).getRange(source.row,1).getDisplayValue();if(!job.names.some(name=>sstKey_(name)===sstKey_(value)))throw Error('A source name changed unexpectedly in '+source.title+'. Review the correction before continuing.');});
 props.setProperty('SST_PENDING_NAME_CORRECTION',JSON.stringify(job));
 general.getRange(job.row,1).setValue(current);job.sources.forEach(source=>ss.getSheetByName(source.title).getRange(source.row,1).setValue(current));
 delete registry[currentKey];entry.name=current;registry[newKey]=entry;sstSaveRegistry_(registry);SST_ROW_CACHE_={};
 if(progress){progress.setName(title);progress.getRange('A1').setValue(current+' — SST PROGRESS');
  const p={name:current,row:job.row,values:general.getRange(job.row,1,1,71).getValues()[0]};
  Object.entries(entry.records||{}).forEach(([n,rec])=>sstControls_(ss,p,Number(n),rec));
 }
 sstRenameIndex_(ss,job.oldName,current);props.deleteProperty('SST_PENDING_NAME_CORRECTION');sstMaybeRefreshNativeTabs_(ss);
 ss.toast('Name corrected: '+job.oldName+' → '+current+'. The same saved records and progress were kept.','Student name updated',8);
 return {name:current,studentId:entry.studentId,sheetId:entry.sheetId};
}
function sstRepairNameConnections_(ss){
 const props=PropertiesService.getDocumentProperties(),pending=props.getProperty('SST_PENDING_NAME_CORRECTION');
 if(pending)sstFinishNameCorrection_(ss,JSON.parse(pending));
 if(props.getProperty('SST_NAME_BINDINGS_VERSION')!=='1')return;
 const general=ss.getSheetByName('General'),names=general.getRange(3,1,general.getMaxRows()-2,1).getDisplayValues().flat(),registry=sstRegistry_();
 for(const [key,entry] of Object.entries(registry)){
  if(!entry.generalNameRange||!entry.name)continue;const anchor=ss.getRangeByName(entry.generalNameRange);if(!anchor)continue;
  const row=anchor.getRow(),actual=general.getRange(row,1).getDisplayValue().trim();if(actual===entry.name)continue;
  const existing=names.findIndex(v=>sstKey_(v)===key);
  if(existing>=0&&existing+3!==row){ss.setNamedRange(entry.generalNameRange,general.getRange(existing+3,1));continue;}
  if(!actual){general.getRange(row,1).setValue(entry.name);continue;}
  const job=sstPrepareNameCorrection_(ss,entry.name,actual,row,entry);props.setProperty('SST_PENDING_NAME_CORRECTION',JSON.stringify(job));sstFinishNameCorrection_(ss,job);
 }
}
function sstGeneralNamesEdited_(e){
 const ss=e.source,general=e.range.getSheet(),first=Math.max(3,e.range.getRow()),last=e.range.getLastRow(),registry=sstRegistry_();
 if(e.range.getNumRows()>1){
  const corrections=[];for(let row=first;row<=last;row++){const pair=sstNameAtRow_(ss,row,registry);if(pair&&general.getRange(row,1).getDisplayValue().trim()!==pair[1].name)corrections.push({row,name:pair[1].name});}
  if(corrections.length){corrections.forEach(x=>general.getRange(x.row,1).setValue(x.name));ss.toast('For spelling corrections, change one existing name cell at a time. The original names were restored.','General names',9);return;}
 }
 for(let row=first;row<=last;row++){
  const cell=general.getRange(row,1),newName=String(cell.getDisplayValue()).trim().replace(/\s+/g,' ');
  let pair=e.range.getNumRows()===1&&e.oldValue!=null&&registry[sstKey_(e.oldValue)]?[sstKey_(e.oldValue),registry[sstKey_(e.oldValue)]]:sstNameAtRow_(ss,row,registry);
  if(pair&&pair[1].studentId){const entry=pair[1],oldName=entry.name||String(e.oldValue||pair[0]);
   if(newName===oldName)continue;
   if(!newName){cell.setValue(oldName);ss.toast('Go to SST Case Manager → Students — General → Edit name / archive / delete student.','What would you like to do?',-1);return;}
   try{const job=sstPrepareNameCorrection_(ss,oldName,newName,row,entry);PropertiesService.getDocumentProperties().setProperty('SST_PENDING_NAME_CORRECTION',JSON.stringify(job));sstFinishNameCorrection_(ss,job);}
   catch(error){if(!PropertiesService.getDocumentProperties().getProperty('SST_PENDING_NAME_CORRECTION')){cell.setValue(oldName);ss.toast(error.message+' The original name was restored.','Name correction stopped',10);return;}throw error;}
  }else{
   if(!newName)continue;
   const duplicates=general.getRange(3,1,general.getMaxRows()-2,1).getDisplayValues().flat().filter(v=>sstKey_(v)===sstKey_(newName));
   if(duplicates.length>1||registry[sstKey_(newName)]){cell.setValue(e.oldValue||'');ss.toast('That student name already exists. The new entry was not connected to their records.','Duplicate name',9);return;}
  }
 }
 const people=sstPeople_(ss);sstEnsurePeople_(ss,people);sstBindStudentIds_(ss,people);
 // A correction kept the same roster length; avoid rebuilding unrelated records.
 if(e.oldValue==null||!String(e.oldValue).trim()||e.range.getNumRows()>1)sstIndex_(ss,people);
}
function sstInstallAutomaticNames(){
 const ss=sstBook_(),lock=LockService.getDocumentLock();lock.waitLock(30000);
 try{const people=sstPeople_(ss);sstEnsurePeople_(ss,people);sstBindStudentIds_(ss,people);PropertiesService.getDocumentProperties().setProperty('SST_NAME_BINDINGS_VERSION','1');
  const general=ss.getSheetByName('General'),header=general.getRange('A2'),note=header.getNote();if(!note.includes('AUTOMATIC NAME CORRECTIONS'))header.setNote(note+'\nAUTOMATIC NAME CORRECTIONS: To fix spelling, edit the existing Name cell in this same row and press Enter. Wait a moment for the connected tabs and records to update. Add a different student on a new blank row. Keep names unique. Do not erase a student name; use Archive. Correct one name at a time.');
  const q=ss.getSheetByName('Q&A - Start Here');if(q){const data=q.getDataRange().getDisplayValues(),items=[
   ['Names and rows','Can I rename a student?','Yes. In General, correct the name in the existing student row in column A and press Enter. The connected SST name cells, Student Index, Progress tab title, saved record connections, and report selection update together. The student keeps the same internal identity and their saved goals, dates, scores, notes, and meeting history. Correct one name at a time.','Works now'],
   ['Names and rows','Will changing one letter create a new student?','No. Correcting the existing Name cell in the same General row keeps that student and their records together. A genuinely different student belongs on a new blank General row.','Works now'],
   ['Names and rows','What if I use a name that already belongs to someone else?','The correction stops and restores the original name. No student records are merged. Use a unique name for each student, such as adding a last initial when needed.','Works now'],
   ['Names and rows','Does a name correction rewrite old goals or notes?','The student name fields and their connections update. Existing written goals, intervention wording, comments, scores, dates, and progress entries are preserved. A name typed inside an old goal or comment is part of that saved text and is not automatically rewritten.','Works now'],
   ['Names and rows','Can I replace a departing student with a different child in the same row?','No. Editing an existing name means you are correcting the spelling for the same child. Archive the departing student, then add the new child in a new blank General row.','Works now'],
   ['Names and rows','What if I accidentally clear the name?','Clearing an existing student name is stopped and the previous name is restored. Use Tracker controls > Archive to hide a student while keeping their records.','Works now']
  ];items.forEach(item=>{const i=data.findIndex(row=>row[1]===item[1]),r=i>=0?i+1:q.getLastRow()+1;q.getRange(r,1,1,4).setValues([item]).setFontFamily('Arial').setFontSize(11).setWrap(true).setVerticalAlignment('top');q.setRowHeight(r,110);});}
  console.log('Automatic same-row name corrections installed for '+people.length+' students. Existing names and progress were kept.');
 }finally{lock.releaseLock();}
}
