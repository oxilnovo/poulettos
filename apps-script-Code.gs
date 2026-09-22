/**
 * Poulettos V3 — Google Sheets backend
 *
 * Sheets created automatically:
 *   Users: email | name | active
 *   Hens: userEmail | id | name | breed | emoji | photo | status | deceasedAt | updatedAt | deleted
 *   Eggs: userEmail | id | date | weight | henId | note | updatedAt | deleted
 *
 * Security:
 *   - The browser sends the Google Identity Services ID token.
 *   - The token is validated server-side with Google's tokeninfo endpoint.
 *   - The email must also exist in Users with active=TRUE.
 *   - Never trust the email supplied by the browser.
 *
 * Deployment:
 *   Deploy > New deployment > Web app
 *   Execute as: Me
 *   Who has access: Anyone
 *   Put the /exec URL in config.js.
 */

const SPREADSHEET_ID = ''; // Optional: leave empty when this script is bound to the target Sheet.
const GOOGLE_CLIENT_ID = ''; // Same client ID as Poulettos config.js.
const USERS_SHEET = 'Users';
const HENS_SHEET = 'Hens';
const EGGS_SHEET = 'Eggs';

function db_() {
  return SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function ensureSheets_() {
  const ss = db_();
  const specs = [
    [USERS_SHEET, ['email','name','active']],
    [HENS_SHEET, ['userEmail','id','name','breed','emoji','photo','status','deceasedAt','updatedAt','deleted']],
    [EGGS_SHEET, ['userEmail','id','date','weight','henId','note','updatedAt','deleted']]
  ];
  specs.forEach(([name, headers]) => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) sh.getRange(1,1,1,headers.length).setValues([headers]);
  });
}

function doGet(e) {
  ensureSheets_();
  return json_({ok:true, service:'Poulettos', version:'3.0'});
}

function doPost(e) {
  try {
    ensureSheets_();
    const body = JSON.parse(e.postData.contents || '{}');
    if (body.action !== 'sync') return json_({ok:false,error:'Unknown action'});
    const auth = authenticate_(body.idToken);
    if (!auth.ok) return json_({ok:false,error:auth.error});
    if (!isAllowed_(auth.email)) return json_({ok:false,error:'User not authorized'});

    const lock = LockService.getScriptLock();
    lock.waitLock(15000);
    try {
      mergeRows_(auth.email, body.clientState || {});
      const state = readState_(auth.email);
      return json_({ok:true,state:state,user:{email:auth.email,name:auth.name}});
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    return json_({ok:false,error:String(err && err.message || err)});
  }
}

function authenticate_(idToken) {
  if (!idToken) return {ok:false,error:'Missing Google ID token'};
  if (!GOOGLE_CLIENT_ID) return {ok:false,error:'GOOGLE_CLIENT_ID is not configured on the server'};
  const url = 'https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(idToken);
  const res = UrlFetchApp.fetch(url, {muteHttpExceptions:true});
  if (res.getResponseCode() !== 200) return {ok:false,error:'Invalid Google token'};
  const p = JSON.parse(res.getContentText());
  if (p.aud !== GOOGLE_CLIENT_ID) return {ok:false,error:'Token audience mismatch'};
  if (String(p.email_verified) !== 'true') return {ok:false,error:'Google account is not verified'};
  return {ok:true,email:String(p.email).toLowerCase(),name:p.name || String(p.email).split('@')[0]};
}

function isAllowed_(email) {
  const sh = db_().getSheetByName(USERS_SHEET);
  const values = sh.getDataRange().getValues();
  for (let i=1;i<values.length;i++) {
    if (String(values[i][0]).trim().toLowerCase() === email) {
      return String(values[i][2]).toLowerCase() !== 'false' && String(values[i][2]).toLowerCase() !== '0';
    }
  }
  return false;
}

function mergeRows_(email, state) {
  const ss = db_();
  mergeEntitySheet_(ss.getSheetByName(HENS_SHEET), email, state.hens || [], 'hen');
  mergeEntitySheet_(ss.getSheetByName(EGGS_SHEET), email, state.entries || [], 'egg');
}

function mergeEntitySheet_(sh, email, entities, type) {
  const values = sh.getDataRange().getValues();
  const index = new Map();
  for (let i=1;i<values.length;i++) {
    if (String(values[i][0]).toLowerCase() === email) index.set(String(values[i][1]), i+1);
  }
  entities.forEach(entity => {
    const id = String(entity.id || '');
    if (!id) return;
    const row = type === 'hen'
      ? [email,id,entity.name||'',entity.breed||'',entity.emoji||'',entity.photo||'',entity.status||'active',entity.deceasedAt||'',entity.updatedAt||new Date().toISOString(),entity.deleted?'TRUE':'FALSE']
      : [email,id,entity.date||'',Number(entity.weight)||0,entity.henId||'unknown',entity.note||'',entity.updatedAt||new Date().toISOString(),entity.deleted?'TRUE':'FALSE'];
    const existingRow = index.get(id);
    if (existingRow) {
      const existing = sh.getRange(existingRow,1,1,row.length).getValues()[0];
      const oldTime = new Date(existing[type === 'hen' ? 8 : 6] || 0).getTime();
      const newTime = new Date(row[type === 'hen' ? 8 : 6] || 0).getTime();
      if (!oldTime || newTime >= oldTime) sh.getRange(existingRow,1,1,row.length).setValues([row]);
    } else {
      sh.appendRow(row);
    }
  });
}

function readState_(email) {
  const ss = db_();
  const hens = rowsForUser_(ss.getSheetByName(HENS_SHEET), email).filter(r=>r.deleted !== 'TRUE').map(r=>({
    id:r.id,name:r.name,breed:r.breed,emoji:r.emoji||'🐔',photo:r.photo||'',status:r.status||'active',deceasedAt:r.deceasedAt||null
  }));
  const entries = rowsForUser_(ss.getSheetByName(EGGS_SHEET), email).filter(r=>r.deleted !== 'TRUE').map(r=>({
    id:r.id,date:r.date,weight:Number(r.weight)||0,henId:r.henId||'unknown',note:r.note||'',updatedAt:r.updatedAt||null
  }));
  const user = userByEmail_(email) || {email,name:email.split('@')[0]};
  return {user:{name:user.name||email.split('@')[0],email},hens,entries,meta:{updatedAt:new Date().toISOString()}};
}

function rowsForUser_(sh,email) {
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0];
  return values.slice(1).filter(r=>String(r[0]).toLowerCase()===email).map(r=>{
    const o={}; headers.forEach((h,i)=>o[h]=r[i]); return o;
  });
}

function userByEmail_(email) {
  const values=db_().getSheetByName(USERS_SHEET).getDataRange().getValues();
  for(let i=1;i<values.length;i++) if(String(values[i][0]).toLowerCase()===email) return {email:values[i][0],name:values[i][1],active:values[i][2]};
  return null;
}
