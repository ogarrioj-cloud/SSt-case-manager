// Warning-only safeguards. These do not remove anyone's editing access.
function sstInstallEditingWarnings(){
 const ss=sstBook_(),lock=LockService.getDocumentLock();lock.waitLock(30000);
 try{
  const before=sstWarningFingerprint_(ss);PropertiesService.getDocumentProperties().setProperty('SST_EDIT_WARNINGS_VERSION','1');
  ss.getSheets().forEach(s=>{
   if(s.getName().startsWith('Progress - ')){sstWarningProgress_(ss,s);return;}
   if(s.getName()===SST_INDEX){sstWarningIndex_(ss);return;}
   let inputs=[];
   if(s.getName()==='General')inputs=[s.getRange('A3:AS'+s.getMaxRows()),s.getRange('BQ3:BR'+s.getMaxRows())];
   else if(SST_ROUNDS.includes(s.getName()))inputs=[s.getRange('AS3:BJ'+s.getMaxRows())];
   sstSetEditingWarning_(s,inputs);
  });
  const after=sstWarningFingerprint_(ss);if(before!==after)throw Error('Cell-content preservation check did not match; inspect before continuing.');
  const result=ss.getSheets().map(s=>{const p=s.getProtections(SpreadsheetApp.ProtectionType.SHEET).find(p=>p.getDescription()==='SST Safety — review before changing protected areas');if(!p||!p.isWarningOnly())throw Error('Warning setup missing on '+s.getName());return{sheet:s.getName(),warningOnly:p.isWarningOnly(),inputAreas:p.getUnprotectedRanges().length};});
  console.log('EDIT WARNINGS INSTALLED '+JSON.stringify({sheets:result.length,cellContentsAndFormulasPreserved:true,details:result}));ss.toast('Editing warnings are on for headings, linked cells, saved text and shared lists. Normal input boxes remain available.','SST safety',10);
 }finally{lock.releaseLock();}
}
function sstWarningOn_(){return PropertiesService.getDocumentProperties().getProperty('SST_EDIT_WARNINGS_VERSION')==='1';}
function sstSetEditingWarning_(sheet,inputs){
 const title='SST Safety — review before changing protected areas';let p=sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET).find(x=>x.getDescription()===title);
 if(!p)p=sheet.protect().setDescription(title).setWarningOnly(true);
 p.setWarningOnly(true);p.setUnprotectedRanges(inputs);
}
function sstWarningIndex_(ss){if(!sstWarningOn_())return;const s=ss.getSheetByName(SST_INDEX);if(!s)return;const count=sstActivePeople_(ss).length,inputs=[s.getRange('B3:B4'),s.getRange('B8:G8')];if(count)inputs.push(s.getRange(13,9,count,2));sstSetEditingWarning_(s,inputs);}
function sstWarningProgress_(ss,sheet){
 if(!sstWarningOn_())return;const e=Object.values(sstRegistry_()).find(e=>e.sheetId===sheet.getSheetId()),inputs=[];
 if(e)Object.values(e.records||{}).forEach(rec=>{
  if(rec.autoVersion!==2)return;const r=sstNamed_(ss,rec.rangeName).getRow(),c=sstNamed_(ss,rec.controlName).getRow();
  inputs.push(sheet.getRange(r+1,2),sheet.getRange(r+1,4),sheet.getRange(r+1,6,1,3),sheet.getRange(c,2,2,1));
  (rec.blocks||[]).forEach(b=>{const range=sstNamed_(ss,b.rangeName),r=range.getRow(),bounds=sstProgressBounds_(range);inputs.push(sheet.getRange(r+3,2,2,2),sheet.getRange(r+3,5,2,4),sheet.getRange(r+5,2),sheet.getRange(r+5,4,1,5),sheet.getRange(bounds.first,1,bounds.count,8),sheet.getRange(bounds.review,2,1,7));});
 });
 const labels=sheet.getRange(1,1,Math.max(1,sheet.getLastRow()),1).getDisplayValues().flat();labels.forEach((v,i)=>{if(v==='Assessment date')inputs.push(sheet.getRange(i+1,2,1,7));});
 sstSetEditingWarning_(sheet,inputs);
}
function sstWarningFingerprint_(ss){
 const payload=ss.getSheets().map(s=>{const r=s.getDataRange();return{id:s.getSheetId(),name:s.getName(),values:r.getValues(),formulas:r.getFormulas()};});
 return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,JSON.stringify(payload)));
}
