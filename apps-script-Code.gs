const SPREADSHEET_ID = '';
const GOOGLE_CLIENT_ID = '663179170370-bp361u1avupqp2ut4fqt1hlhbcikbal0.apps.googleusercontent.com';
const USERS_SHEET = 'Users';
const HENS_SHEET = 'Hens';
const EGGS_SHEET = 'Eggs';

function db_() { return SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet(); }
function json_(obj) { return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON); }

function ensureSheets_() {
  const ss = db_();
  const specs = [
    [USERS_SHEET, ['email','name','active']],
    [HENS_SHEET, ['userEmail','id','name','breed','emoji','photo','status','deceasedAt','updatedAt','deleted']],
    [EGGS_SHEET, ['userEmail','id','date','weight','henId','note','updatedAt','deleted']]
  ];
  specs.forEach(([name,headers]) => { let sh=ss.getSheetByName(name); if(!sh) sh=ss.insertSheet(name); if(sh.getLastRow()===0) sh.getRange(1,1,1,headers.length).setValues([headers]); });
}

function doGet() { ensureSheets_(); repairEggSheet_(); return json_({ok:true,service:'Poulettos',version:'3.12'}); }

function doPost(e) {
  try {
    ensureSheets_();
    repairEggSheet_();
    const body=JSON.parse(e.postData.contents||'{}');
    if(body.action!=='sync') return json_({ok:false,error:'Unknown action'});
    const auth=authenticate_(body.idToken);
    if(!auth.ok) return json_({ok:false,error:auth.error});
    if(!isAllowed_(auth.email)) return json_({ok:false,error:'User not authorized'});

    // The lock is only held while the two user datasets are merged. The merge itself
    // uses bulk reads/writes so large datasets (thousands of eggs) do not hold the lock
    // for dozens of seconds.
    const lock=LockService.getScriptLock();
    if(!lock.tryLock(30000)) return json_({ok:false,error:'Synchronisation temporairement occupée. Réessayez dans quelques secondes.'});
    try {
      mergeRows_(auth.email,body.clientState||{});
      return json_({ok:true,state:readState_(auth.email),user:{email:auth.email,name:auth.name}});
    } finally { lock.releaseLock(); }
  } catch(err) {
    return json_({ok:false,error:String(err&&err.message||err)});
  }
}

function authenticate_(idToken) {
  if(!idToken) return {ok:false,error:'Missing Google ID token'};
  if(!GOOGLE_CLIENT_ID) return {ok:false,error:'GOOGLE_CLIENT_ID is not configured on the server'};
  const url='https://oauth2.googleapis.com/tokeninfo?id_token='+encodeURIComponent(idToken);
  const res=UrlFetchApp.fetch(url,{muteHttpExceptions:true});
  if(res.getResponseCode()!==200) return {ok:false,error:'Invalid Google token'};
  const p=JSON.parse(res.getContentText());
  if(p.aud!==GOOGLE_CLIENT_ID) return {ok:false,error:'Token audience mismatch'};
  if(String(p.email_verified)!=='true') return {ok:false,error:'Google account is not verified'};
  return {ok:true,email:String(p.email).toLowerCase(),name:p.name||String(p.email).split('@')[0]};
}

function isAllowed_(email) {
  const values=db_().getSheetByName(USERS_SHEET).getDataRange().getValues();
  for(let i=1;i<values.length;i++) if(String(values[i][0]).trim().toLowerCase()===email) return String(values[i][2]).toLowerCase()!=='false'&&String(values[i][2]).toLowerCase()!=='0';
  return false;
}

function mergeRows_(email,state) {
  const ss=db_();
  const known=state.meta&&state.meta.serverKnownIds?state.meta.serverKnownIds:{hens:[],entries:[]};
  const knownHens=new Set(known.hens||[]);
  const knownEntries=new Set(known.entries||[]);

  // Permanent deletion model. Deleted IDs live in Script Properties rather than as
  // tombstone rows in the spreadsheet.
  recordClientDeletions_(email,state.deletedHenIds||[],'hen');
  recordClientDeletions_(email,state.deletedEntryIds||[],'egg');
  recordMissingAsDeleted_(email,ss.getSheetByName(HENS_SHEET),knownHens,'hen');
  recordMissingAsDeleted_(email,ss.getSheetByName(EGGS_SHEET),knownEntries,'egg');

  purgeDeletedRows_(ss.getSheetByName(HENS_SHEET),email,'hen');
  purgeDeletedRows_(ss.getSheetByName(EGGS_SHEET),email,'egg');

  // Merge in bulk. New rows are appended with one setValues() call instead of
  // one appendRow() call per egg. Existing rows are only written when their
  // updatedAt is newer than the server copy.
  mergeEntitySheetBulk_(ss.getSheetByName(HENS_SHEET),email,state.hens||[],'hen');
  mergeEntitySheetBulk_(ss.getSheetByName(EGGS_SHEET),email,state.entries||[],'egg');

  purgeDeletedRows_(ss.getSheetByName(HENS_SHEET),email,'hen');
  purgeDeletedRows_(ss.getSheetByName(EGGS_SHEET),email,'egg');
}

function mergeEntitySheetBulk_(sh,email,entities,type) {
  if(!sh) return;
  const lastRow=sh.getLastRow();
  const numCols=type==='hen'?10:8;
  const values=lastRow>=2?sh.getRange(2,1,lastRow-1,numCols).getValues():[];
  const index=new Map();
  for(let i=0;i<values.length;i++){
    if(String(values[i][0]).trim().toLowerCase()===email){
      const id=String(values[i][1]||'');
      if(id) index.set(id,i+2);
    }
  }

  const updates=[];
  const appends=[];
  const updatedCol=type==='hen'?9:7;
  const now=new Date().toISOString();

  (entities||[]).forEach(entity=>{
    const id=String(entity.id||'');
    if(!id||isDeleted_(email,type,id)) return;
    const row=type==='hen'
      ? [email,id,entity.name||'',entity.breed||'',entity.emoji||'',entity.photo||'',entity.status||'active',entity.deceasedAt||'',entity.updatedAt||now,'FALSE']
      : [email,id,entity.date||'',Number(entity.weight)||0,entity.henId||'unknown',entity.note||'',entity.updatedAt||now,'FALSE'];
    const existingRow=index.get(id);
    if(existingRow){
      const existing=sh.getRange(existingRow,1,1,numCols).getValues()[0];
      const serverTime=existing[updatedCol-1]?new Date(existing[updatedCol-1]).getTime():0;
      const clientTime=entity.updatedAt?new Date(entity.updatedAt).getTime():0;
      if(!serverTime||clientTime>=serverTime) updates.push({row:existingRow,values:row});
    } else {
      appends.push(row);
    }
  });

  // Existing rows may be scattered because multiple users share the sheet, so
  // update only the rows that actually changed. This is normally a very small set.
  updates.forEach(u=>sh.getRange(u.row,1,1,numCols).setValues([u.values]));

  // All new rows are appended in one operation — critical for large imports.
  if(appends.length){
    const start=sh.getLastRow()+1;
    sh.getRange(start,1,appends.length,numCols).setValues(appends);
  }
}

function deletionKey_(email,type,id){
  return 'POULETTOS_DELETED|' + String(type) + '|' + String(email).trim().toLowerCase() + '|' + String(id);
}
function isDeleted_(email,type,id){
  return PropertiesService.getScriptProperties().getProperty(deletionKey_(email,type,id)) === '1';
}
function markDeleted_(email,type,id){
  if(id) PropertiesService.getScriptProperties().setProperty(deletionKey_(email,type,id),'1');
}
function recordClientDeletions_(email,ids,type){
  (ids||[]).forEach(id=>markDeleted_(email,type,String(id)));
}
function recordMissingAsDeleted_(email,sh,knownIds,type){
  if(!sh||!knownIds.size)return;
  const values=sh.getDataRange().getValues(),present=new Set();
  for(let i=1;i<values.length;i++){
    if(String(values[i][0]).trim().toLowerCase()!==email)continue;
    const id=String(values[i][1]||'');
    if(id)present.add(id);
  }
  knownIds.forEach(id=>{if(id&&!present.has(id))markDeleted_(email,type,id);});
}
function purgeDeletedRows_(sh,email,type){
  if(!sh||sh.getLastRow()<2)return;
  const numCols=type==='hen'?10:8;
  const lastRow=sh.getLastRow();
  const values=sh.getRange(1,1,lastRow,numCols).getValues();
  const kept=[values[0]];
  let changed=false;
  for(let i=1;i<values.length;i++){
    const r=values[i];
    if(String(r[0]).trim().toLowerCase()!==email){ kept.push(r); continue; }
    const id=String(r[1]||'');
    const deletedCol=type==='hen'?10:8;
    const flag=String(r[deletedCol-1]||'').trim().toUpperCase();
    if(id&&(flag==='TRUE'||isDeleted_(email,type,id))){ changed=true; continue; }
    kept.push(r);
  }
  if(!changed)return;
  // One bulk rewrite avoids thousands of deleteRow() calls on large sheets.
  sh.getRange(1,1,lastRow,numCols).clearContent();
  sh.getRange(1,1,kept.length,numCols).setValues(kept);
  if(kept.length<lastRow) sh.deleteRows(kept.length+1,lastRow-kept.length);
}

function readState_(email) {
  const ss=db_(),henRows=rowsForUser_(ss.getSheetByName(HENS_SHEET),email),eggRows=rowsForUser_(ss.getSheetByName(EGGS_SHEET),email);
  const hens=henRows.filter(r=>String(r.deleted).toUpperCase()!=='TRUE'&&!isDeleted_(email,'hen',String(r.id))).map(r=>({id:r.id,name:r.name,breed:r.breed,emoji:r.emoji||'🐔',photo:r.photo||'',status:r.status||'active',deceasedAt:r.deceasedAt||null,updatedAt:r.updatedAt||null}));
  const entries=eggRows.filter(r=>String(r.deleted).toUpperCase()!=='TRUE'&&!isDeleted_(email,'egg',String(r.id))).map(r=>({id:r.id,date:r.date,weight:Number(r.weight)||0,henId:r.henId||'unknown',note:r.note||'',updatedAt:r.updatedAt||null}));
  const serverKnownIds={hens:hens.map(r=>String(r.id)).filter(Boolean),entries:entries.map(r=>String(r.id)).filter(Boolean)};
  const user=userByEmail_(email)||{email,name:email.split('@')[0]};
  return {user:{name:user.name||email.split('@')[0],email},hens,entries,meta:{updatedAt:new Date().toISOString(),syncedAt:new Date().toISOString(),serverKnownIds}};
}

function repairEggSheet_(){
  const sh=db_().getSheetByName(EGGS_SHEET);
  if(!sh)return;
  const lastRow=sh.getLastRow(); if(lastRow<2)return;
  const values=sh.getRange(2,1,lastRow-1,10).getValues();
  values.forEach((r,i)=>{
    const h=String(r[7]??'').trim().toUpperCase(),j=String(r[9]??'').trim().toUpperCase();
    if(!h && (j==='TRUE'||j==='FALSE')){
      sh.getRange(i+2,8).setValue(j); sh.getRange(i+2,10).clearContent();
    }
  });
}

function rowsForUser_(sh,email){
  const values=sh.getDataRange().getValues(); if(values.length<2)return [];
  const headers=values[0];
  return values.slice(1).filter(r=>String(r[0]).trim().toLowerCase()===email).map(r=>{const o={};headers.forEach((h,i)=>o[h]=r[i]);return o;});
}
function userByEmail_(email){const values=db_().getSheetByName(USERS_SHEET).getDataRange().getValues();for(let i=1;i<values.length;i++)if(String(values[i][0]).trim().toLowerCase()===email)return {email:values[i][0],name:values[i][1],active:values[i][2]};return null;}

// Manual edits in Google Sheets get a fresh timestamp, so the Sheet edit wins
// over an older local copy during the next synchronization.
function onEdit(e){
  try{
    const sh=e&&e.range&&e.range.getSheet();
    if(!sh||(sh.getName()!==HENS_SHEET&&sh.getName()!==EGGS_SHEET))return;
    const row=e.range.getRow(); if(row<=1)return;
    const updatedCol=sh.getName()===HENS_SHEET?9:7;
    const first=e.range.getColumn(),last=first+e.range.getNumColumns()-1;
    if(first<=updatedCol&&updatedCol<=last)return;
    const email=String(sh.getRange(row,1).getValue()).trim(),id=String(sh.getRange(row,2).getValue()).trim();
    if(!email||!id)return;
    sh.getRange(row,updatedCol).setValue(new Date().toISOString());
  }catch(err){console.warn('Poulettos onEdit',err);}
}
