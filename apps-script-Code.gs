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

function doGet() { ensureSheets_(); return json_({ok:true,service:'Poulettos',version:'3.1'}); }

function doPost(e) {
  try {
    ensureSheets_();
    const body=JSON.parse(e.postData.contents||'{}');
    if(body.action!=='sync') return json_({ok:false,error:'Unknown action'});
    const auth=authenticate_(body.idToken);
    if(!auth.ok) return json_({ok:false,error:auth.error});
    if(!isAllowed_(auth.email)) return json_({ok:false,error:'User not authorized'});
    const lock=LockService.getScriptLock(); lock.waitLock(15000);
    try { mergeRows_(auth.email,body.clientState||{}); return json_({ok:true,state:readState_(auth.email),user:{email:auth.email,name:auth.name}}); }
    finally { lock.releaseLock(); }
  } catch(err) { return json_({ok:false,error:String(err&&err.message||err)}); }
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
  mergeEntitySheet_(ss.getSheetByName(HENS_SHEET),email,state.hens||[],'hen',new Set(known.hens||[]));
  mergeEntitySheet_(ss.getSheetByName(EGGS_SHEET),email,state.entries||[],'egg',new Set(known.entries||[]));
}

function mergeEntitySheet_(sh,email,entities,type,knownIds) {
  const values=sh.getDataRange().getValues(),index=new Map();
  for(let i=1;i<values.length;i++) if(String(values[i][0]).trim().toLowerCase()===email) index.set(String(values[i][1]),i+1);
  entities.forEach(entity=>{
    const id=String(entity.id||''); if(!id)return;
    const existingRow=index.get(id),isDeleted=!!entity.deleted;
    // If the client knew this record at the last successful sync but the row is now absent,
    // the deletion happened in Sheets. Do not recreate it from the stale local copy.
    if(!existingRow&&knownIds.has(id)) return;
    if(!existingRow&&isDeleted) return;
    const now=new Date().toISOString();
    const row=type==='hen'
      ? [email,id,entity.name||'',entity.breed||'',entity.emoji||'',entity.photo||'',entity.status||'active',entity.deceasedAt||'',entity.updatedAt||now,isDeleted?'TRUE':'FALSE']
      : [email,id,entity.date||'',Number(entity.weight)||0,entity.henId||'unknown',entity.note||'',entity.updatedAt||now,isDeleted?'TRUE':'FALSE'];
    if(existingRow){
      const existing=sh.getRange(existingRow,1,1,row.length).getValues()[0];
      const updatedCol=type==='hen'?9:7;
      const serverTime=existing[updatedCol-1]?new Date(existing[updatedCol-1]).getTime():0;
      const clientTime=entity.updatedAt?new Date(entity.updatedAt).getTime():0;
      if(isDeleted){
        if(!serverTime||clientTime>=serverTime){
          sh.getRange(existingRow,updatedCol,1,1).setValue(entity.updatedAt||now);
          sh.getRange(existingRow,10,1,1).setValue('TRUE');
        }
      } else if(!serverTime||clientTime>=serverTime) {
        sh.getRange(existingRow,1,1,row.length).setValues([row]);
      }
    } else {
      sh.appendRow(row);
    }
  });
}

function readState_(email) {
  const ss=db_(),henRows=rowsForUser_(ss.getSheetByName(HENS_SHEET),email),eggRows=rowsForUser_(ss.getSheetByName(EGGS_SHEET),email);
  const hens=henRows.filter(r=>String(r.deleted).toUpperCase()!=='TRUE').map(r=>({id:r.id,name:r.name,breed:r.breed,emoji:r.emoji||'🐔',photo:r.photo||'',status:r.status||'active',deceasedAt:r.deceasedAt||null,updatedAt:r.updatedAt||null}));
  const entries=eggRows.filter(r=>String(r.deleted).toUpperCase()!=='TRUE').map(r=>({id:r.id,date:r.date,weight:Number(r.weight)||0,henId:r.henId||'unknown',note:r.note||'',updatedAt:r.updatedAt||null}));
  const serverKnownIds={hens:henRows.map(r=>String(r.id)).filter(Boolean),entries:eggRows.map(r=>String(r.id)).filter(Boolean)};
  const deletedIds={hens:[],entries:[]};
  const user=userByEmail_(email)||{email,name:email.split('@')[0]};
  return {user:{name:user.name||email.split('@')[0],email},hens,entries,meta:{updatedAt:new Date().toISOString(),syncedAt:new Date().toISOString(),serverKnownIds,deletedIds}};
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
