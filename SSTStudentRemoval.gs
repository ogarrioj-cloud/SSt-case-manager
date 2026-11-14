// Remove only empty name-only entries. Saved student work must be archived.
function sstInstallRemovalChoice(){
 const ss=sstBook_(),q=ss.getSheetByName('Q&A - Start Here');
 const header=ss.getSheetByName('General').getRange('A2'),note=header.getNote().split('\nREMOVE OR ARCHIVE')[0];header.setNote(note+'\nREMOVE OR ARCHIVE: Clearing a Name cell keeps the name temporarily. Select the Name cell and use SST Progress → Remove unused entry or archive student to choose. Remove unused entry is available only for a name-only entry with no student information or saved work. Archive keeps the student information and hides them from active lists. Cancel keeps the student.');
 if(q){const items=[
 ['Names and rows','What if I accidentally clear the name?','The name is temporarily restored so its connections stay together. To choose what happens, select its Name cell in General and open SST Progress → Remove unused entry or archive student. Choose Keep student / Cancel if this was a mistake. Remove unused entry is available only for a name-only entry with no student information or saved work. For a student with information or progress, use Archive student to keep everything.','Works now'],
 ['Names and rows','I entered a test name. How do I remove it?','Select the test Name cell in General. Open SST Progress → Remove unused entry or archive student. Choose Remove unused entry. Remove is available only when the row has no student information and no goals, saved meetings, or Progress tab. It clears the unused name and its automatic links without deleting spreadsheet rows. If you simply clear the Name cell and it returns, use this menu instead.','Works now'],
 ['Names and rows','Should I remove or archive a student?','Remove unused entry is for an accidental name with no other information or saved work. Archive student keeps all information and history, hides the Progress tab, and removes the student from active lists. General and SST source rows remain saved. Use Tracker controls → Restore selected student to bring an archived student back. If you are unsure, choose Keep student.','Safe practice']];
 items.forEach(item=>{const data=q.getDataRange().getDisplayValues(),i=data.findIndex(r=>r[1]===item[1]),row=i>=0?i+1:q.getLastRow()+1;q.getRange('A106:D106').copyTo(q.getRange(row,1,1,4),SpreadsheetApp.CopyPasteType.PASTE_FORMAT,false);q.getRange(row,1,1,4).setValues([item]).setWrap(true).setFontFamily('Arial').setFontSize(11);q.autoResizeRows(row,1);});}
 const audit=sstPeople_(ss).map(p=>{const c=sstUnusedEntryCheck_(ss,p.name);return{name:p.name,removalAllowed:c.allowed,reason:c.reason};});
 console.log('REMOVAL CHOICE INSTALLED '+JSON.stringify(audit));
}
function sstUnusedEntryCheck_(ss,name){
 const key=sstKey_(name),people=sstPeople_(ss),p=people.find(p=>sstKey_(p.name)===key),registry=sstRegistry_(),entry=registry[key];
 const no=reason=>({allowed:false,reason,name:p?p.name:name});
 if(!p||!entry||!entry.studentId)return no('The student connection needs checking first.');
 if(PropertiesService.getDocumentProperties().getProperty('SST_PENDING_NAME_CORRECTION'))return no('Finish the pending name correction first.');
 if(Object.keys(entry.records||{}).length||entry.sheetId)return no('This student has a saved meeting or Progress tracker. Archive the student to keep their work.');
 if(entry.meetingCheck)return no('This student has a saved meeting number confirmation. Archive the student to keep it.');
 const g=ss.getSheetByName('General'),r=g.getRange(p.row,2,1,g.getMaxColumns()-1);
 if(r.getValues()[0].some(v=>v!==''&&v!==null)||r.getFormulas()[0].some(Boolean)||g.getRange(p.row,1,1,g.getMaxColumns()).getNotes()[0].some(Boolean))return no('This General row contains information or notes. Archive the student to keep them.');
 const sources=[];
 for(const title of SST_ROUNDS){
  const s=ss.getSheetByName(title),names=s.getRange(3,1,s.getMaxRows()-2,1).getDisplayValues().flat(),matches=[];names.forEach((n,i)=>{if(sstKey_(n)===key)matches.push(i+3);});
  if(matches.length!==1)return no('The name connection in '+title+' needs checking first.');
  const row=matches[0],goals=s.getRange(row,45,1,18);
  if(goals.getValues()[0].some(v=>v!==''&&v!==null)||goals.getFormulas()[0].some(Boolean)||s.getRange(row,1,1,s.getMaxColumns()).getNotes()[0].some(Boolean))return no(title+' contains goals, plans, or notes. Archive the student to keep them.');
  if(s.getRange(row,2).getFormula()!==sstSafeBackgroundFormula_(row,'B','AR')||s.getRange(row,63).getFormula()!==sstSafeBackgroundFormula_(row,'BK','BO'))return no('The linked background in '+title+' needs checking first.');
  sources.push({title,row});
 }
 return{allowed:true,reason:'This is a name-only entry. There is no student information, saved meeting, goal, or progress tracker.',name:p.name,row:p.row,sources,entry};
}
function sstReviewStudentRemoval(){
 const ss=sstBook_(),sheet=ss.getActiveSheet();let name='';
 if(sheet.getName()==='General'&&sheet.getActiveRange().getRow()>=3)name=sheet.getRange(sheet.getActiveRange().getRow(),1).getDisplayValue().trim();
 else{const pair=Object.values(sstRegistry_()).find(e=>e.sheetId===sheet.getSheetId());name=pair?pair.name:ss.getSheetByName(SST_INDEX).getRange('B3').getDisplayValue().trim();}
 if(!name){SpreadsheetApp.getUi().alert('Select the student’s Name cell in General first.');return;}
 sstShowRemovalChoice_(ss,name);
}
function sstShowRemovalChoice_(ss,name){return sstShowStudentActions(name);}
function sstStudentRemovalAction(action,name){
 if(action==='archive')return sstManageAction('archive',name,{});
 if(action!=='remove')throw Error('Choose Remove unused entry or Archive student.');
 const ss=sstBook_(),lock=LockService.getDocumentLock();lock.waitLock(30000);
 try{
  const c=sstUnusedEntryCheck_(ss,name);if(!c.allowed)throw Error('Nothing was removed. '+c.reason);
  const registry=sstRegistry_(),other=JSON.stringify(Object.entries(registry).filter(([k])=>k!==sstKey_(name))),proof=sstRemovalKeptData_(ss,name);
  const backup={at:new Date().toISOString(),name:c.name,row:c.row,entry:c.entry,sources:c.sources.map(x=>({...x,background:ss.getSheetByName(x.title).getRange(x.row,2).getFormula(),tail:ss.getSheetByName(x.title).getRange(x.row,63).getFormula()}))};
  PropertiesService.getDocumentProperties().setProperty('SST_UNUSED_REMOVAL_'+c.entry.studentId,JSON.stringify(backup));
  c.sources.forEach(x=>{const s=ss.getSheetByName(x.title);s.getRange(x.row,2).clearContent();s.getRange(x.row,63).clearContent();s.getRange(x.row,1).clearContent();});
  if(c.entry.generalNameRange)ss.removeNamedRange(c.entry.generalNameRange);
  delete registry[sstKey_(name)];sstSaveRegistry_(registry);ss.getSheetByName('General').getRange(c.row,1).clearContent();SST_ROW_CACHE_={};
  const index=ss.getSheetByName(SST_INDEX);if(sstKey_(index.getRange('B3').getDisplayValue())===sstKey_(name))index.getRange('B3').clearContent();
  PropertiesService.getDocumentProperties().deleteProperty('SST_INDEX_ROSTER_V2');sstIndex_(ss,sstPeople_(ss));SpreadsheetApp.flush();
  if(proof!==sstRemovalKeptData_(ss,name)||other!==JSON.stringify(Object.entries(sstRegistry_()).filter(([k])=>k!==sstKey_(name))))throw Error('Removal finished but the preservation check needs review. Do not make further changes.');
  console.log('UNUSED ENTRY REMOVED '+JSON.stringify({name:c.name,row:c.row,sourceRows:c.sources,otherStudentsAndProgressPreserved:true}));
  return{message:'Removed the unused entry '+c.name+'. Its General name and automatic SST links are clear. No spreadsheet rows were deleted, and other students and progress were kept.'};
 }finally{lock.releaseLock();}
}
function sstRemovalKeptData_(ss,exclude){
 const people=sstPeople_(ss).filter(p=>sstKey_(p.name)!==sstKey_(exclude)),g=ss.getSheetByName('General');
 const data=people.map(p=>({name:p.name,general:g.getRange(p.row,1,1,g.getMaxColumns()).getValues(),goals:SST_ROUNDS.map(t=>{const s=ss.getSheetByName(t),r=sstFindRow_(s,p.name);return s.getRange(r,45,1,18).getValues();})}));
 ss.getSheets().filter(s=>s.getName().startsWith('Progress - ')&&s.getName()!==sstProgressTitle_(exclude)).forEach(s=>{const r=s.getDataRange();data.push({sheet:s.getSheetId(),values:r.getValues(),formulas:r.getFormulas()});});
 return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,JSON.stringify(data)));
}
