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
    [HENS_SHEET, ['userEmail','id','name','breed','emoji','photo','status','deceasedAt','updatedAt','deleted','createdBy','createdAt','editedBy']],
    [EGGS_SHEET, ['userEmail','id','date','weight','henId','note','updatedAt','deleted','createdBy','createdAt','editedBy']]
  ];
  specs.forEach(([name,headers]) => { let sh=ss.getSheetByName(name); if(!sh) sh=ss.insertSheet(name); if(sh.getLastRow()===0) sh.getRange(1,1,1,headers.length).setValues([headers]); });
}

function doGet() { ensureSheets_(); migrateHeaders_(); return json_({ok:true,service:'Poulettos',version:'3.18'}); }

function doPost(e) {
  try {
    ensureSheets_();
    migrateHeaders_();
    const body=JSON.parse(e.postData.contents||'{}');
    if(body.action!=='sync') return json_({ok:false,error:'Unknown action'});
    const auth=authenticate_(body.idToken);
    if(!auth.ok) return json_({ok:false,error:auth.error});
    if(!isAllowed_(auth.email)) return json_({ok:false,error:'User not authorized'});

    // The lock is only held while the two user datasets are merged. The merge itself
    // uses bulk reads/writes so large datasets (thousands of eggs) do not hold the lock
    // for dozens of seconds.
    const lock=LockService.getScriptLock();
    if(!lock.tryLock(15000)) return json_({ok:false,error:'Synchronisation temporairement occupée. Réessayez dans quelques secondes.'});
    try {
      mergeRows_(auth.email,body.clientState||{});
    } finally { lock.releaseLock(); }
    // Reading the resulting state can be expensive with large datasets. Do it
    // after releasing the global lock so another user's sync is not blocked.
    return json_({ok:true,state:readState_(auth.email),user:{email:auth.email,name:auth.name}});
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

  // Record explicit client deletions first. They are kept in Script Properties,
  // not in the spreadsheet, so deleted rows can be physically removed.
  recordClientDeletions_(email,state.deletedHenIds||[],'hen');
  recordClientDeletions_(email,state.deletedEntryIds||[],'egg');

  // One fast pass per sheet. No deleteRow(), appendRow(), or per-cell writes.
  mergeEntitySheetFast_(ss.getSheetByName(HENS_SHEET),email,state.hens||[],knownHens,'hen');
  // Eggs are a shared dataset: every authorized user sees the same Eggs sheet.
  mergeEntitySheetFast_(ss.getSheetByName(EGGS_SHEET),email,state.entries||[],knownEntries,'egg');
}

function mergeEntitySheetFast_(sh,email,entities,knownIds,type) {
  if(!sh) return;
  const numCols=type==='hen'?13:11;
  const readCols=numCols; // include legacy J column in Eggs for cleanup
  const lastRow=sh.getLastRow();
  const all=lastRow>=1?sh.getRange(1,1,lastRow,readCols).getValues():[];
  if(!all.length) return;
  const header=all[0].slice(0,numCols);
  const rows=all.slice(1);
  const emailKey=String(email).trim().toLowerCase();
  const shared = true;
  const incomingById=new Map();
  const deletedSet=getDeletedSet_(email,type);
  (entities||[]).forEach(entity=>{
    const id=String(entity.id||'');
    if(id) incomingById.set(id,entity);
  });

  // If an ID was known locally but has disappeared from the Sheet, that is a
  // permanent deletion made in Google Sheets. Remember it so the client cannot
  // recreate it on the next sync.
  const presentIds=new Set();
  rows.forEach(r=>{
    if(shared || String(r[0]).trim().toLowerCase()===emailKey){
      const id=String(r[1]||'');
      if(id) presentIds.add(id);
    }
  });
  const missingKnown=[]; knownIds.forEach(id=>{ if(id && !presentIds.has(String(id))) missingKnown.push(String(id)); });
  if(missingKnown.length){ markDeletedMany_(email,type,missingKnown); missingKnown.forEach(id=>deletedSet.add(String(id))); }

  let changed=false;
  const deletedDuringMerge=[];
  const kept=[];
  kept.push(header);

  // Keep all other users untouched. For this user, remove permanent deletions
  // and apply newer client versions in memory.
  const existingIds=new Set();
  rows.forEach(r=>{
    const row=r.slice(0,numCols);
    const rowEmail=String(row[0]).trim().toLowerCase();
    if(!shared && rowEmail!==emailKey){ kept.push(row); return; }
    const id=String(row[1]||'');
    const legacyJ=type==='egg' && String(r[9]??'').trim().toUpperCase()==='TRUE' && String(row[7]??'').trim()==='';
    if(id && (String(row[type==='hen'?9:7]??'').trim().toUpperCase()==='TRUE' || deletedSet.has(id) || legacyJ)){
      if(id) deletedDuringMerge.push(id);
      changed=true;
      return;
    }
    if(id) existingIds.add(id);
    const entity=incomingById.get(id);
    if(entity && !deletedSet.has(id)){
      const updatedCol=type==='hen'?9:7;
      const serverTime=row[updatedCol-1]?new Date(row[updatedCol-1]).getTime():0;
      const clientTime=entity.updatedAt?new Date(entity.updatedAt).getTime():0;
      if(!serverTime || clientTime>serverTime){
        const now=new Date().toISOString();
        const replacement=type==='hen'
          ? [row[0]||entity.createdBy||email,id,entity.name||'',entity.breed||'',entity.emoji||'',entity.photo||'',entity.status||'active',normalizeDateOnly_(entity.deceasedAt),entity.updatedAt||now,'FALSE',row[10]||entity.createdBy||email,row[11]||entity.createdAt||entity.updatedAt||now,entity.editedBy||email]
          : [row[0]||email,id,normalizeDateOnly_(entity.date),Number(entity.weight)||0,entity.henId||'unknown',entity.note||'',entity.updatedAt||now,'FALSE',row[8]||entity.createdBy||email,row[9]||entity.createdAt||entity.updatedAt||now,entity.editedBy||''];
        kept.push(replacement);
        changed=true;
        return;
      }
    }
    kept.push(row);
  });

  if(deletedDuringMerge.length){ markDeletedMany_(email,type,deletedDuringMerge); deletedDuringMerge.forEach(id=>deletedSet.add(String(id))); }

  // Append new client rows in memory. One setValues() writes the entire sheet
  // only when something actually changed. For ~1,500 eggs this is still one
  // spreadsheet operation instead of thousands of appendRow/setValue calls.
  (entities||[]).forEach(entity=>{
    const id=String(entity.id||'');
    if(!id || existingIds.has(id) || deletedSet.has(id)) return;
    const now=new Date().toISOString();
    const row=type==='hen'
      ? [email,id,entity.name||'',entity.breed||'',entity.emoji||'',entity.photo||'',entity.status||'active',normalizeDateOnly_(entity.deceasedAt),entity.updatedAt||now,'FALSE',entity.createdBy||email,entity.createdAt||entity.updatedAt||now,entity.editedBy||'']
      : [email,id,normalizeDateOnly_(entity.date),Number(entity.weight)||0,entity.henId||'unknown',entity.note||'',entity.updatedAt||now,'FALSE',entity.createdBy||email,entity.createdAt||entity.updatedAt||now,entity.editedBy||''];
    kept.push(row);
    existingIds.add(id);
    changed=true;
  });

  // Remove accidental legacy columns beyond the official schema when possible.
  

  if(!changed) return;
  // Single bulk rewrite. This is intentionally outside any per-row operation.
  // The resulting dataset can be larger than the current sheet. Clear only the
  // currently used range, then size the destination range to the actual data.
  // (The previous version tried to write kept.length rows into a lastRow-sized
  // range, which made large first imports fail silently from the client side.)
  if(lastRow>0) sh.getRange(1,1,lastRow,numCols).clearContent();
  if(kept.length>0) sh.getRange(1,1,kept.length,numCols).setValues(kept);
  if(kept.length<lastRow) sh.deleteRows(kept.length+1,lastRow-kept.length);
}

function deletionStoreKey_(email,type){
  return type==='egg' ? 'POULETTOS_DELETED|egg|shared' : 'POULETTOS_DELETED|hen|shared';
}
function getDeletedSet_(email,type){
  const props=PropertiesService.getScriptProperties();
  const out=new Set();
  const raw=props.getProperty(deletionStoreKey_(email,type));
  if(raw){try{JSON.parse(raw).map(String).forEach(id=>out.add(id));}catch(e){}}
  // Migrate legacy per-user egg deletion stores into the shared egg store.
  if(type==='egg'){
    const all=props.getProperties();
    Object.keys(all).filter(k=>k.indexOf('POULETTOS_DELETED|egg|')===0 && k!=='POULETTOS_DELETED|egg|shared').forEach(k=>{try{JSON.parse(all[k]).map(String).forEach(id=>out.add(id));}catch(e){}});
  }
  return out;
}
function saveDeletedSet_(email,type,set){
  PropertiesService.getScriptProperties().setProperty(deletionStoreKey_(email,type),JSON.stringify(Array.from(set)));
}
function isDeleted_(email,type,id){
  return getDeletedSet_(email,type).has(String(id));
}
function markDeleted_(email,type,id){
  if(!id)return;
  const key=deletionStoreKey_(email,type);
  const props=PropertiesService.getScriptProperties();
  let set=new Set();
  const raw=props.getProperty(key);
  if(raw){try{set=new Set(JSON.parse(raw).map(String));}catch(e){}}
  set.add(String(id));
  props.setProperty(key,JSON.stringify(Array.from(set)));
}
function markDeletedMany_(email,type,ids){
  const clean=Array.from(new Set((ids||[]).map(String).filter(Boolean)));
  if(!clean.length)return;
  const props=PropertiesService.getScriptProperties(),key=deletionStoreKey_(email,type);
  let set=new Set(); const raw=props.getProperty(key);
  if(raw){try{set=new Set(JSON.parse(raw).map(String));}catch(e){}}
  clean.forEach(id=>set.add(id));
  props.setProperty(key,JSON.stringify(Array.from(set)));
}
function recordClientDeletions_(email,ids,type){ markDeletedMany_(email,type,ids); }
function recordMissingAsDeleted_(email,sh,knownIds,type){
  if(!sh||!knownIds.size)return;
  const values=sh.getDataRange().getValues(),present=new Set();
  for(let i=1;i<values.length;i++){
    if(String(values[i][0]).trim().toLowerCase()!==email)continue;
    const id=String(values[i][1]||''); if(id)present.add(id);
  }
  const missing=[]; knownIds.forEach(id=>{if(id&&!present.has(String(id)))missing.push(String(id));});
  markDeletedMany_(email,type,missing);
}

function purgeDeletedRows_(sh,email,type){
  if(!sh||sh.getLastRow()<2)return;
  const numCols=type==='hen'?13:11;
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

function normalizeDateOnly_(value){
  if(value===null||value===undefined||value==='') return '';
  if(Object.prototype.toString.call(value)==='[object Date]' && !isNaN(value.getTime())){
    return Utilities.formatDate(value, Session.getScriptTimeZone()||'Europe/Paris','yyyy-MM-dd');
  }
  const s=String(value);
  const m=s.match(/^(\d{4}-\d{2}-\d{2})/);
  if(m) return m[1];
  const d=new Date(value);
  return isNaN(d.getTime())?'':Utilities.formatDate(d, Session.getScriptTimeZone()||'Europe/Paris','yyyy-MM-dd');
}
function normalizeTimestamp_(value){
  if(value===null||value===undefined||value==='') return null;
  if(Object.prototype.toString.call(value)==='[object Date]' && !isNaN(value.getTime())) return value.toISOString();
  const d=new Date(value);
  return isNaN(d.getTime())?String(value):d.toISOString();
}
function rowsAll_(sh){
  if(!sh)return [];
  const values=sh.getDataRange().getValues(); if(values.length<2)return [];
  const headers=values[0];
  return values.slice(1).map(r=>{const o={};headers.forEach((h,i)=>o[h]=r[i]);return o;});
}
function readState_(email) {
  const ss=db_(),henRows=rowsAll_(ss.getSheetByName(HENS_SHEET)),eggRows=rowsAll_(ss.getSheetByName(EGGS_SHEET));
  const deletedHens=getDeletedSet_(email,'hen'), deletedEggs=getDeletedSet_(email,'egg');
  const hens=henRows.filter(r=>String(r.deleted).toUpperCase()!=='TRUE'&&!deletedHens.has(String(r.id))).map(r=>({
    id:r.id,name:r.name,breed:r.breed,emoji:r.emoji||'🐔',photo:r.photo||'',status:r.status||'active',
    deceasedAt:normalizeDateOnly_(r.deceasedAt)||null,updatedAt:normalizeTimestamp_(r.updatedAt),createdBy:r.createdBy||r.userEmail||'',createdAt:normalizeTimestamp_(r.createdAt)||normalizeTimestamp_(r.updatedAt),editedBy:r.editedBy||''
  }));
  const entries=eggRows.filter(r=>String(r.deleted).toUpperCase()!=='TRUE'&&!deletedEggs.has(String(r.id))).map(r=>({
    id:r.id,date:normalizeDateOnly_(r.date),weight:Number(r.weight)||0,henId:r.henId||'unknown',note:r.note||'',updatedAt:normalizeTimestamp_(r.updatedAt),createdBy:r.createdBy||r.userEmail||'',createdAt:normalizeTimestamp_(r.createdAt)||normalizeTimestamp_(r.updatedAt),editedBy:r.editedBy||''
  }));
  const serverKnownIds={hens:hens.map(r=>String(r.id)).filter(Boolean),entries:entries.map(r=>String(r.id)).filter(Boolean)};
  const user=userByEmail_(email)||{email,name:email.split('@')[0]};
  return {user:{name:user.name||email.split('@')[0],email},hens,entries,meta:{updatedAt:new Date().toISOString(),syncedAt:new Date().toISOString(),serverKnownIds}};
}

function migrateHeaders_(){
  repairEggSheet_();
  const ss=db_();
  const specs=[
    [HENS_SHEET,['userEmail','id','name','breed','emoji','photo','status','deceasedAt','updatedAt','deleted','createdBy','createdAt','editedBy']],
    [EGGS_SHEET,['userEmail','id','date','weight','henId','note','updatedAt','deleted','createdBy','createdAt','editedBy']]
  ];
  specs.forEach(([name,headers])=>{
    const sh=ss.getSheetByName(name); if(!sh)return;
    const existing=sh.getRange(1,1,1,Math.max(sh.getLastColumn(),headers.length)).getValues()[0];
    headers.forEach((h,i)=>{if(String(existing[i]||'').trim()!==h) sh.getRange(1,i+1).setValue(h);});
  });
}

function repairEggSheet_(){
  const sh=db_().getSheetByName(EGGS_SHEET);
  if(!sh)return;
  const lastRow=sh.getLastRow(); if(lastRow<2)return;
  const values=sh.getRange(2,1,lastRow-1,10).getValues();
  let changed=false;
  values.forEach(r=>{
    const h=String(r[7]??'').trim().toUpperCase(),j=String(r[9]??'').trim().toUpperCase();
    if(!h && (j==='TRUE'||j==='FALSE')){ r[7]=j; r[9]=''; changed=true; }
  });
  if(changed) sh.getRange(2,1,values.length,10).setValues(values);
}

// Run this manually once only if an old version left TRUE/FALSE in column J of Eggs.
function repairEggSheetManually(){ repairEggSheet_(); }

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
