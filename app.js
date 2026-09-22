(() => {
  // Register the service worker as early as possible so the app can boot offline on subsequent launches.
  if ('serviceWorker' in navigator) { navigator.serviceWorker.register('./sw.js', {scope:'./'}).then(r=>r.update()).catch(()=>{}); }
  const KEY='poulettos-state-v3';
  const CACHE_KEY='poulettos-local-cache-v3';
  const APP_VERSION='3.13';
  const TOKEN_KEY='poulettos-google-id-token-v1';
  const UNKNOWN='unknown';
  const defaultState={
    user:{name:'',email:''},
    hens:[],
    entries:[],
    meta:{updatedAt:null,deviceId:null,serverKnownIds:{hens:[],entries:[]}}
  };
  let state=load();
  let localReady=false;
  const DB_NAME='poulettos-local-db-v1', DB_STORE='state', DB_KEY='current';
  let period='week', periodAnchor=isoDate(new Date());
  let historySort='date', historyOrder='desc';
  let totalRangeEnabled=false,totalRangeStart='',totalRangeEnd='',editingId=null,googleIdToken='',syncInProgress=false;
  try{googleIdToken=localStorage.getItem(TOKEN_KEY)||'';}catch{}
  const $=s=>document.querySelector(s), $$=s=>[...document.querySelectorAll(s)];
  const uid=()=>crypto.randomUUID?crypto.randomUUID():'id-'+Date.now()+'-'+Math.random().toString(36).slice(2);
  const deviceId=(()=>{let id=localStorage.getItem('poulettos-device');if(!id){id=uid();localStorage.setItem('poulettos-device',id)}return id})();

  function normalizeState(s){
    const base=structuredClone(defaultState);
    base.user=s?.user||base.user;
    if(String(base.user.email||'').includes('demo@poulettos.local')) base.user={name:'',email:''};
    base.hens=(s?.hens||[]).map(h=>({...h,status:h.status||'active',deceasedAt:h.deceasedAt||null,updatedAt:h.updatedAt||null,deleted:!!h.deleted})).filter(h=>!h.deleted);
    base.entries=(s?.entries||[]).flatMap(e=>{
      const hid=e.henId ?? (e.henIds?.[0]||UNKNOWN);
      return [{id:e.id||uid(),date:e.date,weight:Number(e.weight)||0,henId:hid,note:e.note||'',updatedAt:e.updatedAt||null,deleted:!!e.deleted}];
    }).filter(e=>!e.deleted);
    base.deletedEntryIds=[...(s?.deletedEntryIds||[])];
    base.deletedHenIds=[...(s?.deletedHenIds||[])];
    base.meta={...(s?.meta||{}),deviceId};
    if(s?.meta?.serverKnownIds) base.meta.serverKnownIds={hens:[...(s.meta.serverKnownIds.hens||[])],entries:[...(s.meta.serverKnownIds.entries||[])]};
    else if(s?.meta?.knownIds) base.meta.serverKnownIds={hens:[...(s.meta.knownIds.hens||[])],entries:[...(s.meta.knownIds.entries||[])]};
    return base;
  }
  function openLocalDB(){
    return new Promise((resolve,reject)=>{
      if(!('indexedDB' in window)) return resolve(null);
      const req=indexedDB.open(DB_NAME,1);
      req.onupgradeneeded=()=>{try{req.result.createObjectStore(DB_STORE);}catch{}};
      req.onsuccess=()=>resolve(req.result);
      req.onerror=()=>resolve(null);
    });
  }
  async function readIndexedState(){
    const db=await openLocalDB();
    if(!db)return null;
    return new Promise(resolve=>{
      try{const tx=db.transaction(DB_STORE,'readonly'),req=tx.objectStore(DB_STORE).get(DB_KEY);req.onsuccess=()=>resolve(req.result||null);req.onerror=()=>resolve(null);}catch{resolve(null)}
    });
  }
  async function writeIndexedState(value){
    const db=await openLocalDB();
    if(!db)return;
    await new Promise(resolve=>{try{const tx=db.transaction(DB_STORE,'readwrite');tx.objectStore(DB_STORE).put(value,DB_KEY);tx.oncomplete=()=>resolve();tx.onerror=()=>resolve();}catch{resolve()}});
  }
  function snapshotScore(s){
    const n=(s?.hens?.length||0)+(s?.entries?.length||0);
    const t=new Date(s?.meta?.updatedAt||0).getTime()||0;
    return {count:n,time:t,hasUser:!!s?.user?.email};
  }
  function chooseLocalSnapshot(a,b){
    if(!a)return b||null;
    if(!b)return a;
    const A=snapshotScore(a),B=snapshotScore(b);
    // Prefer the most recently saved snapshot. If timestamps are equal/missing,
    // never replace a populated snapshot with an empty one.
    if(B.time>A.time)return b;
    if(A.time>B.time)return a;
    if(B.count>A.count)return b;
    if(A.count>B.count)return a;
    if(B.hasUser&&!A.hasUser)return b;
    return a;
  }
  async function restoreLocalState(){
    // localStorage is synchronous and is available immediately; IndexedDB is the
    // durable database. Reconcile both instead of blindly letting an older/empty
    // IndexedDB snapshot overwrite valid local data.
    const current=load();
    const indexed=await readIndexedState();
    const chosen=chooseLocalSnapshot(normalizeState(current),indexed?normalizeState(indexed):null) || structuredClone(defaultState);
    state=normalizeState(chosen);
    localReady=snapshotScore(state).count>0 || !!state.user?.email || !!state.meta?.lastSyncAt;
    if(localReady){
      const serialized=JSON.stringify(state);
      try{localStorage.setItem(KEY,serialized);localStorage.setItem(CACHE_KEY,serialized);}catch{}
      await writeIndexedState(state);
    }
    return localReady;
  }
  function load(){
    try{
      const raw=JSON.parse(localStorage.getItem(KEY)||localStorage.getItem(CACHE_KEY)||'null');
      if(raw)return normalizeState(raw);
      const old=JSON.parse(localStorage.getItem('poulettos-state-v2')||localStorage.getItem('pouleco-state-v1')||'null');
      if(old)return normalizeState(old);
    }catch{}
    return structuredClone(defaultState);
  }
  function hasLocalData(){
    try{return !!(localStorage.getItem(KEY)||localStorage.getItem(CACHE_KEY)||localStorage.getItem('poulettos-state-v2')||localStorage.getItem('pouleco-state-v1'));}catch{return false}
  }
  async function save(){
    state.meta.updatedAt=new Date().toISOString();
    state.meta.deviceId=deviceId;
    const serialized=JSON.stringify(state);
    try{localStorage.setItem(KEY,serialized);localStorage.setItem(CACHE_KEY,serialized);}catch{}
    localReady=true;
    await writeIndexedState(state);
  }
  function isoDate(d){const z=n=>String(n).padStart(2,'0');return `${d.getFullYear()}-${z(d.getMonth()+1)}-${z(d.getDate())}`}
  function dateOnly(s){return new Date(`${s}T12:00:00`)}
  function entryDayKey(v){if(typeof v==='string'){const m=v.match(/^(\d{4}-\d{2}-\d{2})/);if(m)return m[1]}return isoDate(new Date(v))}
  function fmt(d){return new Intl.DateTimeFormat('fr-FR',{day:'2-digit',month:'2-digit',year:'numeric'}).format(d)}
  function grams(es){return es.reduce((a,e)=>a+(Number(e.weight)||0),0)}
  function total(es){return es.length}
  function henName(id){return id===UNKNOWN?'Poule inconnue':state.hens.find(h=>h.id===id)?.name||'Poule inconnue'}
  function escapeHtml(s){return String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]))}
  function henVisual(h){return h.photo?`<img src="${h.photo}" alt="">`:escapeHtml(h.emoji||'🐔')}
  function toast(msg){const t=$('#toast');t.textContent=msg;t.classList.add('show');setTimeout(()=>t.classList.remove('show'),1800)}

  function rangeFor(p,anchorStr=periodAnchor){
    const a=dateOnly(anchorStr);a.setHours(0,0,0,0);
    if(p==='week'){const day=(a.getDay()+6)%7,start=new Date(a);start.setDate(a.getDate()-day);const end=new Date(start);end.setDate(start.getDate()+7);return {start,end};}
    if(p==='month'){return {start:new Date(a.getFullYear(),a.getMonth(),1),end:new Date(a.getFullYear(),a.getMonth()+1,1)}}
    if(p==='year'){return {start:new Date(a.getFullYear(),0,1),end:new Date(a.getFullYear()+1,0,1)}}
    if(totalRangeEnabled&&totalRangeStart&&totalRangeEnd){const start=dateOnly(totalRangeStart),end=dateOnly(totalRangeEnd);end.setDate(end.getDate()+1);return {start,end}}
    if(!state.entries.length)return {start:new Date(0),end:new Date(a.getFullYear(),a.getMonth()+1,1)};
    const sorted=state.entries.map(e=>dateOnly(entryDayKey(e.date))).sort((x,y)=>x-y),start=new Date(sorted[0].getFullYear(),sorted[0].getMonth(),1),end=new Date(a.getFullYear(),a.getMonth()+1,1);return {start,end};
  }
  function entriesFor(p){const r=rangeFor(p);return state.entries.filter(e=>{const d=dateOnly(entryDayKey(e.date));return d>=r.start&&d<r.end})}

  function navigate(view){$$('.view').forEach(v=>v.classList.toggle('active',v.id==='view-'+view));$$('.bottom-nav button').forEach(b=>b.classList.toggle('active',b.dataset.nav===view));if(view==='home')renderHome();if(view==='stats')renderStats();if(view==='hens')renderHens();if(view==='history')renderHistory();if(view==='entry'&&!editingId)prepareNewEntry()}
  $$('[data-nav]').forEach(b=>b.addEventListener('click',()=>navigate(b.dataset.nav)));

  function henOptions(selected=''){return `<option value="${UNKNOWN}" ${selected===UNKNOWN?'selected':''}>❓ Poule inconnue</option>`+state.hens.filter(h=>h.status!=='dead').map(h=>`<option value="${h.id}" ${selected===h.id?'selected':''}>${escapeHtml(h.emoji||'🐔')} ${escapeHtml(h.name)}</option>`).join('')}
  function addEggRow(data={weight:'',henId:UNKNOWN}){const row=document.createElement('div');row.className='egg-row';row.dataset.rowId=uid();row.innerHTML=`<div class="egg-index">🥚</div><div class="egg-fields"><div><label>Poids</label><div class="weight-input"><input class="egg-weight" type="number" min="0" max="200" step="0.1" inputmode="decimal" placeholder="ex. 62" value="${data.weight??''}" required><span>g</span></div></div><div><label>Poule</label><select class="egg-hen">${henOptions(data.henId||UNKNOWN)}</select></div></div><button type="button" class="remove-egg" aria-label="Retirer">×</button>`;$('.egg-rows').appendChild(row);updateEggIndexes()}
  function updateEggIndexes(){$$('.egg-row').forEach((r,i)=>r.querySelector('.egg-index').textContent='🥚 '+(i+1))}
  function prepareNewEntry(){editingId=null;$('#entryEyebrow').textContent='Saisie rapide';$('#entryTitle').textContent='Nouveaux œufs';$('#saveEntryBtn').textContent='Enregistrer';$('#cancelEdit').classList.add('hidden');$('#entryDate').value=isoDate(new Date());$('#note').value='';$('.egg-rows').innerHTML='';addEggRow()}
  function editEntry(id){const e=state.entries.find(x=>x.id===id);if(!e)return;editingId=id;$('#entryEyebrow').textContent='Modification';$('#entryTitle').textContent='Modifier l’œuf';$('#saveEntryBtn').textContent='Enregistrer les changements';$('#cancelEdit').classList.remove('hidden');$('#entryDate').value=entryDayKey(e.date);$('#note').value=e.note||'';$('.egg-rows').innerHTML='';addEggRow({weight:e.weight,henId:e.henId||UNKNOWN});navigate('entry')}
  $('#addEggRow').onclick=()=>addEggRow();$('.egg-rows').addEventListener('click',e=>{const b=e.target.closest('.remove-egg');if(!b)return;const rows=$$('.egg-row');if(rows.length===1){toast('Il faut au moins un œuf');return}b.closest('.egg-row').remove();updateEggIndexes()});$('#cancelEdit').onclick=()=>{editingId=null;navigate('history')};

  $('#entryForm').addEventListener('submit',e=>{e.preventDefault();const date=$('#entryDate').value,rows=$$('.egg-row').map(r=>({weight:Number(r.querySelector('.egg-weight').value),henId:r.querySelector('.egg-hen').value}));if(!date||rows.some(x=>!Number.isFinite(x.weight)||x.weight<=0)){toast('Renseignez le poids de chaque œuf');return}const stamp=`${date}T12:00:00`;
    if(editingId){const e0=state.entries.find(x=>x.id===editingId);if(e0){Object.assign(e0,{date:stamp,weight:rows[0].weight,henId:rows[0].henId,note:$('#note').value.trim(),updatedAt:new Date().toISOString()});save();toast('Œuf modifié ✓')}editingId=null;navigate('history');syncSoon();return}
    rows.forEach(x=>state.entries.unshift({id:uid(),date:stamp,weight:x.weight,henId:x.henId,note:$('#note').value.trim(),updatedAt:new Date().toISOString()}));state.entries.sort((a,b)=>entryDayKey(b.date).localeCompare(entryDayKey(a.date)));save();toast(`${rows.length} œuf${rows.length>1?'s':''} enregistré${rows.length>1?'s':''} ✓`);navigate('home');syncSoon();
  });

  function renderHome(){const today=entryDayKey(new Date()),todayE=state.entries.filter(e=>entryDayKey(e.date)===today),week=entriesFor('week'),month=entriesFor('month');$('#todayCount').textContent=total(todayE);$('#weekCount').textContent=total(week);$('#monthCount').textContent=total(month);$('#totalCount').textContent=total(state.entries);$('#helloName').textContent=(state.user.name||'vous').split(' ')[0]||'vous';$('#recentList').innerHTML=state.entries.slice(0,6).map(e=>`<button class="list-row clickable" data-edit="${e.id}"><div class="row-left"><span class="egg-mini">🥚</span><div><strong>${Number(e.weight).toFixed(0)} g</strong><div class="row-time">${fmt(dateOnly(entryDayKey(e.date)))} · ${escapeHtml(henName(e.henId))}</div></div></div><span class="chev">›</span></button>`).join('')||'<div class="list-row"><span class="row-time">Aucune saisie pour le moment.</span></div>';$$('[data-edit]').forEach(b=>b.onclick=()=>editEntry(b.dataset.edit))}

  function currentStreak(hid,referenceDate,sourceEntries=state.entries){const days=new Set(sourceEntries.filter(e=>e.henId===hid).map(e=>entryDayKey(e.date)));let ref;if(referenceDate instanceof Date)ref=isoDate(referenceDate);else if(typeof referenceDate==='string'&&/^\d{4}-\d{2}-\d{2}$/.test(referenceDate))ref=referenceDate;else return 0;let d=dateOnly(ref),streak=0;while(days.has(isoDate(d))){streak++;d.setDate(d.getDate()-1)}return streak}
  function longestStreak(hid,sourceEntries=state.entries){const days=[...new Set(sourceEntries.filter(e=>e.henId===hid).map(e=>entryDayKey(e.date)))].sort();let best=0,run=0,prev=null;for(const day of days){const d=dateOnly(day);if(prev){const diff=Math.round((d-prev)/86400000);run=diff===1?run+1:1}else run=1;best=Math.max(best,run);prev=d}return best}

  function renderStats(){const es=entriesFor(period),t=total(es),w=grams(es),anchor=dateOnly(periodAnchor),r=rangeFor(period),today=dateOnly(isoDate(new Date()));$('#periodTotal').textContent=t;$('#averageWeight').textContent=t?`${(w/t).toFixed(1).replace('.',',')} g`:'0 g';$('#periodDate').value=periodAnchor;$('#periodDate').disabled=period==='all';
    if(period==='week')$('#chartCaption').textContent=r.start.getTime()===rangeFor('week',isoDate(today)).start.getTime()?'cette semaine':`semaine du ${r.start.toLocaleDateString('fr-FR',{day:'numeric',month:'short'})}`;else if(period==='month')$('#chartCaption').textContent=anchor.getFullYear()===today.getFullYear()&&anchor.getMonth()===today.getMonth()?'ce mois':anchor.toLocaleDateString('fr-FR',{month:'long',year:'numeric'});else if(period==='year')$('#chartCaption').textContent=anchor.getFullYear()===today.getFullYear()?'cette année':String(anchor.getFullYear());else $('#chartCaption').textContent='depuis le début';
    $('#periodHint').textContent=period==='week'?'La date choisit la semaine correspondante.':period==='month'?'La date choisit le mois correspondant.':period==='year'?'La date choisit l’année correspondante.':(totalRangeEnabled?'Analyse personnalisée entre les deux dates.':'Le total regroupe tout l’historique.');$('#totalRangeOptions').classList.toggle('hidden',period!=='all');$('#totalRangePicker').classList.toggle('hidden',period!=='all'||!totalRangeEnabled);$('#totalRangeToggle').checked=totalRangeEnabled;if(totalRangeEnabled){$('#totalStart').value=totalRangeStart;$('#totalEnd').value=totalRangeEnd}
    let buckets=[];if(period==='week'||period==='month'){for(let d=new Date(r.start);d<r.end;d.setDate(d.getDate()+1))buckets.push(new Date(d))}else if(period==='year'){for(let d=new Date(r.start);d<r.end;d.setMonth(d.getMonth()+1))buckets.push(new Date(d))}else{const months=Math.max(1,(r.end.getFullYear()-r.start.getFullYear())*12+r.end.getMonth()-r.start.getMonth());if(months>18){for(let y=r.start.getFullYear();y<r.end.getFullYear();y++)buckets.push(new Date(y,0,1));if(!buckets.length)buckets=[new Date(anchor.getFullYear(),0,1)]}else for(let d=new Date(r.start);d<r.end;d.setMonth(d.getMonth()+1))buckets.push(new Date(d))}
    const vals=buckets.map(d=>state.entries.filter(e=>{const x=dateOnly(entryDayKey(e.date));if(period==='week'||period==='month')return x.getFullYear()===d.getFullYear()&&x.getMonth()===d.getMonth()&&x.getDate()===d.getDate();if(period==='year')return x.getFullYear()===d.getFullYear()&&x.getMonth()===d.getMonth();const months=Math.max(1,(r.end.getFullYear()-r.start.getFullYear())*12+r.end.getMonth()-r.start.getMonth());return months>18?x.getFullYear()===d.getFullYear():x.getFullYear()===d.getFullYear()&&x.getMonth()===d.getMonth()}).length);
    const labels=buckets.map(d=>period==='week'?d.toLocaleDateString('fr-FR',{weekday:'short',day:'numeric'}).replace('.',''):period==='month'?String(d.getDate()):period==='year'?d.toLocaleDateString('fr-FR',{month:'short'}).replace('.',''):(buckets.length>18?String(d.getFullYear()):d.toLocaleDateString('fr-FR',{month:'short',year:'2-digit'}).replace('.','')));
    const titles=buckets.map(d=>d.toLocaleDateString('fr-FR',{weekday:'long',day:'numeric',month:'long',year:'numeric'})),max=Math.max(1,...vals),chart=$('#barChart'),chartCard=chart.closest('.chart-card');if(period==='all'){chartCard.classList.add('hidden');$('#totalSummary').classList.remove('hidden');renderTotalSummary(es)}else{chartCard.classList.remove('hidden');$('#totalSummary').classList.add('hidden');chart.innerHTML=vals.map((v,i)=>`<div class="bar-wrap" title="${titles[i]} : ${v} œuf${v>1?'s':''}"><div class="bar" style="height:${v?Math.max(6,v/max*100):2}%">${v?`<small>${v}</small>`:''}</div><span class="bar-label">${labels[i]}</span></div>`).join('')}
    const showStreak=period!=='all'||totalRangeEnabled;const streakReference=period==='all'?(totalRangeEnabled?totalRangeEnd:null):periodAnchor;const streakSource=period==='all'?es:state.entries;const counts=state.hens.map(h=>({h,n:es.filter(e=>e.henId===h.id).length,streak:showStreak&&h.status!=='dead'?currentStreak(h.id,streakReference,streakSource):0}));const unknown=es.filter(e=>e.henId===UNKNOWN).length;if(unknown)counts.push({h:{name:'Poule inconnue',emoji:'❓'},n:unknown,streak:showStreak?currentStreak(UNKNOWN,streakReference,streakSource):0});const maxHen=Math.max(1,...counts.map(x=>x.n));$('#henStats').innerHTML=counts.sort((a,b)=>b.n-a.n).map(x=>`<div class="hen-stat"><div class="hen-photo">${x.h.photo?`<img src="${x.h.photo}" alt="">`:escapeHtml(x.h.emoji||'🐔')}</div><div><div class="hen-stat-name">${escapeHtml(x.h.name)}${x.h.status==='dead'?` <span class="archived-badge">archivée</span>`:''}</div><div class="progress"><i style="width:${x.n/maxHen*100}%"></i></div>${showStreak?`<small class="streak">🔥 ${x.streak} jour${x.streak>1?'s':''} de suite</small>`:''}</div><strong>${x.n}</strong></div>`).join('')||'<div class="row-time">Pas encore de données.</div>';
  }

  function renderTotalSummary(es){const dates=es.map(e=>entryDayKey(e.date)).sort(),first=dates[0]?dateOnly(dates[0]):null,last=dates.at(-1)?dateOnly(dates.at(-1)):null;$('#totalSince').textContent=first?(totalRangeEnabled?`${fmt(first)} → ${fmt(last)}`:`Depuis le ${first.toLocaleDateString('fr-FR',{day:'numeric',month:'long',year:'numeric'})}`):'Aucune donnée';$('#totalActiveHens').textContent=state.hens.filter(h=>h.status!=='dead').length;const best=Math.max(0,...state.hens.map(h=>longestStreak(h.id,es)),longestStreak(UNKNOWN,es));$('#totalBestStreak').textContent=`${best} jour${best>1?'s':''}`;const rows=state.hens.map(h=>{const eggs=es.filter(e=>e.henId===h.id);return{name:h.name,emoji:h.emoji||'🐔',photo:h.photo,n:eggs.length,avg:eggs.length?grams(eggs)/eggs.length:0,status:h.status}});const unknown=es.filter(e=>e.henId===UNKNOWN);if(unknown.length)rows.push({name:'Poule inconnue',emoji:'❓',n:unknown.length,avg:grams(unknown)/unknown.length,status:'unknown'});rows.sort((a,b)=>b.n-a.n);$('#totalHenRows').innerHTML=rows.filter(x=>x.n>0).map(x=>`<div class="total-hen-row"><div class="total-hen-id">${x.photo?`<img src="${x.photo}" alt="">`:escapeHtml(x.emoji)}</div><div class="total-hen-name"><strong>${escapeHtml(x.name)}</strong>${x.status==='dead'?'<span class="archived-badge">archivée</span>':''}</div><strong>${x.n}</strong><span>${x.avg.toFixed(1).replace('.',',')} g</span></div>`).join('')||'<div class="row-time">Pas encore de données.</div>'}

  $$('[data-period]').forEach(b=>b.addEventListener('click',()=>{$$('[data-period]').forEach(x=>x.classList.remove('active'));b.classList.add('active');period=b.dataset.period;renderStats()}));$('#periodDate').addEventListener('change',e=>{if(e.target.value){periodAnchor=e.target.value;renderStats()}});$('#totalRangeToggle').addEventListener('change',e=>{totalRangeEnabled=e.target.checked;if(totalRangeEnabled){const sorted=state.entries.map(x=>entryDayKey(x.date)).sort();totalRangeStart=sorted[0]||isoDate(new Date());totalRangeEnd=isoDate(new Date())}renderStats()});$('#totalStart').addEventListener('change',e=>{totalRangeStart=e.target.value;validateTotalRange()});$('#totalEnd').addEventListener('change',e=>{totalRangeEnd=e.target.value;validateTotalRange()});function validateTotalRange(){if(totalRangeStart&&totalRangeEnd&&totalRangeStart>totalRangeEnd){toast('La date de début doit précéder la date de fin');return}renderStats()}

  function renderHens(){const active=state.hens.filter(h=>h.status!=='dead'),dead=state.hens.filter(h=>h.status==='dead');const activeHtml=active.map(h=>{const es=state.entries.filter(e=>e.henId===h.id);return `<div class="hen-item" data-hen-edit="${h.id}"><div class="hen-avatar">${henVisual(h)}</div><div><strong>${escapeHtml(h.name)}</strong><small>${escapeHtml(h.breed||'')}<br>${es.length} œuf${es.length>1?'s':''} · ${grams(es).toFixed(0)} g</small></div><button type="button" class="edit-hen" data-hen-edit="${h.id}" aria-label="Modifier ${escapeHtml(h.name)}">✎</button></div>`}).join('');const deadHtml=dead.length?`<div class="archive-head"><h3>Poules archivées</h3><span>${dead.length}</span></div><div class="hen-list archived-list">${dead.map(h=>{const es=state.entries.filter(e=>e.henId===h.id);return `<div class="hen-item archived" data-hen-edit="${h.id}"><div class="hen-avatar">${henVisual(h)}</div><div><strong>${escapeHtml(h.name)}</strong><small>${escapeHtml(h.breed||'')}<br>Décédée${h.deceasedAt?' le '+fmt(dateOnly(h.deceasedAt)):''} · ${es.length} œuf${es.length>1?'s':''}</small></div><button type="button" class="edit-hen" data-hen-edit="${h.id}" aria-label="Voir ${escapeHtml(h.name)}">✎</button></div>`}).join('')}</div>`:'';$('#henList').innerHTML=(activeHtml||'<div class="row-time">Aucune poule active.</div>')+deadHtml;$$('[data-hen-edit]').forEach(b=>b.onclick=()=>openHenModal(b.dataset.henEdit))}
  function openHenModal(id){const h=state.hens.find(x=>x.id===id);$('#henForm').dataset.id=id||'';$('#henModalEyebrow').textContent=id?'Modifier une poule':'Nouvelle poule';$('#henModalTitle').textContent=id?'Modifier la poule':'Ajouter une poule';$('#henName').value=h?.name||'';$('#henBreed').value=h?.breed||'';$('#henEmoji').value=h?.emoji||'🐔';$('#henImage').value='';$('#deceasedDate').value=h?.deceasedAt||'';$('#deceasedDateWrap').classList.toggle('hidden',!h||h.status!=='dead');const preview=$('#henPreview');preview.innerHTML=h?.photo?`<img src="${h.photo}" alt="">`:escapeHtml(h?.emoji||'🐔');delete preview.dataset.photo;const archiveBtn=$('#archiveHenBtn');if(h){archiveBtn.classList.remove('hidden');archiveBtn.textContent=h.status==='dead'?'↩ Réactiver cette poule':'🐔 Marquer comme décédée'}else archiveBtn.classList.add('hidden');$('#henModal').classList.remove('hidden')}
  function closeHenModal(){$('#henModal').classList.add('hidden')}$('#addHen').onclick=()=>openHenModal('');$('#closeHenModal').onclick=closeHenModal;$('#henModal').addEventListener('click',e=>{if(e.target.id==='henModal')closeHenModal()});$('#archiveHenBtn').onclick=()=>{const id=$('#henForm').dataset.id,h=state.hens.find(x=>x.id===id);if(!h)return;if(h.status==='dead'){h.status='active';h.deceasedAt=null;h.updatedAt=new Date().toISOString();toast('Poule réactivée')}else{if(!confirm(`Archiver ${h.name} comme décédée ? Ses anciennes données seront conservées.`))return;h.status='dead';h.deceasedAt=isoDate(new Date());h.updatedAt=new Date().toISOString();toast('Poule archivée')}save();closeHenModal();renderHens();renderStats();syncSoon()};$('#henImage').addEventListener('change',()=>{const file=$('#henImage').files[0];if(!file)return;const reader=new FileReader();reader.onload=()=>{const img=new Image();img.onload=()=>{const max=500,scale=Math.min(1,max/img.width,max/img.height),c=document.createElement('canvas');c.width=Math.round(img.width*scale);c.height=Math.round(img.height*scale);c.getContext('2d').drawImage(img,0,0,c.width,c.height);$('#henPreview').innerHTML=`<img src="${c.toDataURL('image/jpeg',.82)}" alt="">`;$('#henPreview').dataset.photo=c.toDataURL('image/jpeg',.82)};img.src=reader.result};reader.readAsDataURL(file)});$('#henForm').addEventListener('submit',e=>{e.preventDefault();const id=e.currentTarget.dataset.id,name=$('#henName').value.trim();if(!name){toast('Renseignez un nom');return}let h=state.hens.find(x=>x.id===id);if(!h){h={id:uid(),name:'',breed:'',emoji:'🐔',status:'active',deceasedAt:null};state.hens.push(h)}h.name=name;h.breed=$('#henBreed').value.trim();h.emoji=$('#henEmoji').value.trim()||'🐔';if($('#henPreview').dataset.photo)h.photo=$('#henPreview').dataset.photo;if(h.status==='dead')h.deceasedAt=$('#deceasedDate').value||h.deceasedAt||isoDate(new Date());h.updatedAt=new Date().toISOString();save();closeHenModal();renderHens();renderStats();toast(id?'Poule modifiée ✓':'Poule ajoutée ✓');syncSoon()});

  function renderHistory(){
    const sortSel=$('#historySort'),orderSel=$('#historyOrder');
    if(sortSel)sortSel.value=historySort;if(orderSel)orderSel.value=historyOrder;
    $('#historyCount').textContent=`${state.entries.length} œuf${state.entries.length>1?'s':''}`;
    const rows=[...state.entries].sort((a,b)=>{
      let av,bv;
      if(historySort==='hen'){av=henName(a.henId).toLocaleLowerCase('fr');bv=henName(b.henId).toLocaleLowerCase('fr');}
      else if(historySort==='weight'){av=Number(a.weight)||0;bv=Number(b.weight)||0;}
      else {av=entryDayKey(a.date);bv=entryDayKey(b.date);}
      if(av<bv)return historyOrder==='asc'?-1:1;if(av>bv)return historyOrder==='asc'?1:-1;
      return entryDayKey(b.date).localeCompare(entryDayKey(a.date));
    });
    $('#historyList').innerHTML=rows.map(e=>`<div class="history-row"><button class="history-main" data-edit="${e.id}"><div class="history-egg">🥚</div><div><strong>${Number(e.weight).toFixed(0)} g · ${escapeHtml(henName(e.henId))}</strong><div class="row-time">${fmt(dateOnly(entryDayKey(e.date)))}${e.note?' · '+escapeHtml(e.note):''}</div></div></button><button class="delete-btn" data-delete="${e.id}" aria-label="Supprimer">🗑️</button></div>`).join('')||'<div class="history-row"><span class="row-time">Historique vide.</span></div>';
    $$('[data-edit]').forEach(b=>b.onclick=()=>editEntry(b.dataset.edit));
    $$('[data-delete]').forEach(b=>b.onclick=()=>{if(!confirm('Supprimer cet œuf ?'))return;const deleted=state.entries.find(e=>e.id===b.dataset.delete);if(deleted){deleted.deleted=true;deleted.updatedAt=new Date().toISOString();}state.entries=state.entries.filter(e=>e.id!==b.dataset.delete);state.deletedEntryIds=state.deletedEntryIds||[];if(deleted)state.deletedEntryIds.push(deleted.id);save();renderHistory();renderHome();toast('Entrée supprimée');syncSoon()});
  }
  $('#historySort')?.addEventListener('change',e=>{historySort=e.target.value;renderHistory()});
  $('#historyOrder')?.addEventListener('change',e=>{historyOrder=e.target.value;renderHistory()});

  function logout(){
    try{ if(window.google?.accounts?.id) google.accounts.id.disableAutoSelect(); }catch{}
    googleIdToken='';
    try{localStorage.removeItem(TOKEN_KEY);}catch{}
    localStorage.removeItem(KEY);
    localStorage.removeItem(CACHE_KEY);
    state=structuredClone(defaultState);
    state.meta.deviceId=deviceId;
    $('#profileMenu')?.classList.add('hidden');
    $('#app')?.classList.add('hidden');
    $('#loginScreen')?.classList.remove('hidden');
    toast('Déconnexion effectuée');
  }

  function toggleProfileMenu(){
    const menu=$('#profileMenu');
    if(!menu)return;
    const open=menu.classList.toggle('hidden');
    if(!open) $('#profileEmail').textContent=state.user?.email||'';
  }

  function startApp(user, token){
    state.user={name:user.name||user.email.split('@')[0],email:String(user.email).toLowerCase()};
    if(token){googleIdToken=token;try{localStorage.setItem(TOKEN_KEY,token);}catch{}}
    save();
    $('#loginScreen').classList.add('hidden');
    $('#app').classList.remove('hidden');
    navigate('home');
    setSync(state.meta?.lastSyncAt?'✓ Synchronisé':'● Local',state.meta?.lastSyncAt);
    // Local first: the UI is immediately usable; synchronization runs in background.
    setTimeout(()=>syncNow(),80);
  }
  function decodeJwtPayload(token){
    try{const part=token.split('.')[1].replace(/-/g,'+').replace(/_/g,'/');return JSON.parse(atob(part.padEnd(part.length+((4-part.length%4)%4),'=')));}catch{return null}
  }
  let authRefreshPending=false;
  function initAuth(forcePrompt=false){
    const id=window.POULETTOS_CONFIG?.GOOGLE_CLIENT_ID;if(!id)return;
    const stored=googleIdToken?decodeJwtPayload(googleIdToken):null;
    if(!forcePrompt && stored?.email && stored.exp && stored.exp*1000>Date.now()+30000){
      if(!state.user?.email) state.user={name:stored.name||stored.email.split('@')[0],email:String(stored.email).toLowerCase()};
      if(state.user?.email){startApp({name:state.user.name||stored.name,email:state.user.email},googleIdToken);return;}
    }
    const render=()=>{
      if(!window.google?.accounts?.id)return;
      window.google.accounts.id.initialize({
        client_id:id,auto_select:true,use_fedcm_for_prompt:true,
        callback:r=>{try{
          const p=decodeJwtPayload(r.credential);if(!p?.email)throw new Error('invalid');
          authRefreshPending=false;
          startApp({name:p.name||p.email.split('@')[0],email:p.email},r.credential);
        }catch{authRefreshPending=false;toast('Connexion impossible')}}
      });
      const btn=$('#googleBtn');
      if(btn)window.google.accounts.id.renderButton(btn,{theme:'outline',size:'large',shape:'pill',width:290});
      // With a local session but an expired token, ask Google for a fresh credential.
      // auto_select avoids forcing the account chooser when Google can identify the user.
      try{window.google.accounts.id.prompt()}catch(e){}
    };
    if(window.google?.accounts?.id) render();
    else {
      let tries=0;
      const waitForGoogle=()=>{
        if(window.google?.accounts?.id){render();return;}
        if(++tries<20)setTimeout(waitForGoogle,500);
        else { const hint=$('#authHint'); if(hint)hint.textContent='Connexion Google indisponible. Vérifiez votre connexion puis rechargez la page.'; }
      };
      setTimeout(waitForGoogle,250);
    }
  }
  function ensureFreshGoogleToken(){
    const payload=googleIdToken?decodeJwtPayload(googleIdToken):null;
    if(payload?.exp && payload.exp*1000>Date.now()+60000)return true;
    if(authRefreshPending)return false;
    authRefreshPending=true;
    initAuth(true);
    return false;
  }

  function formatSyncDate(value){if(!value)return 'Jamais';const d=new Date(value);if(Number.isNaN(d.getTime()))return 'Jamais';return d.toLocaleDateString('fr-FR',{day:'2-digit',month:'2-digit'})+' '+d.toLocaleTimeString('fr-FR',{hour:'2-digit',minute:'2-digit'});}
  function setSync(text,lastAt){const p=$('#syncPill');if(p)p.textContent=text;const icon=$('#syncIcon');if(icon)icon.textContent=text.includes('Synchro')?'↻':text.startsWith('✓')?'✓':text.startsWith('!')?'!':'↻';const last=$('#syncLast');if(last)last.textContent=lastAt?formatSyncDate(lastAt):formatSyncDate(state.meta?.lastSyncAt);}
  async function syncNow(retryAfterAuth=false){const cfg=window.POULETTOS_CONFIG||{};if(syncInProgress)return;if(!cfg.API_URL){setSync('! API non configurée');return}if(!navigator.onLine){setSync('! Hors ligne');return}
    if(!googleIdToken || !ensureFreshGoogleToken()){setSync('! Connexion Google…');return}
    syncInProgress=true;setSync('↻ Synchro…');const btn=$('#syncBtn');if(btn)btn.disabled=true;
    try{
      const syncState=structuredClone(state);
      syncState.entries=[...state.entries,...(state.deletedEntryIds||[]).map(id=>({id,deleted:true,updatedAt:new Date().toISOString()}))];
      syncState.hens=[...state.hens,...(state.deletedHenIds||[]).map(id=>({id,deleted:true,updatedAt:new Date().toISOString()}))];
      syncState.meta=syncState.meta||{};syncState.meta.serverKnownIds=state.meta?.serverKnownIds||{hens:[],entries:[]};
      const payload={action:'sync',idToken:googleIdToken,clientState:syncState,deviceId};
      const r=await fetch(cfg.API_URL,{method:'POST',headers:{'Content-Type':'text/plain;charset=utf-8'},body:JSON.stringify(payload)});
      const raw=await r.text();
      let data;try{data=JSON.parse(raw)}catch{throw new Error('Réponse serveur invalide ('+r.status+')')}
      if(!data.ok){
        const msg=String(data.error||'sync failed');
        if(!retryAfterAuth && /token|audience|connexion|auth|invalid/i.test(msg)){
          googleIdToken='';try{localStorage.removeItem(TOKEN_KEY)}catch{}
          syncInProgress=false;if(btn)btn.disabled=false;
          ensureFreshGoogleToken();
          setSync('! Reconnexion Google…');
          return;
        }
        if(!retryAfterAuth && /temporairement occupée|lock timeout|holding the lock/i.test(msg)){
          setSync('↻ Nouvelle tentative…');
          syncInProgress=false;
          if(btn)btn.disabled=false;
          setTimeout(()=>syncNow(true),3500);
          return;
        }
        throw new Error(msg);
      }
      if(data.state){const returned=normalizeState(data.state);returned.meta.lastSyncAt=new Date().toISOString();returned.meta.serverKnownIds=data.state.meta?.serverKnownIds||data.state.meta?.knownIds||returned.meta.serverKnownIds||{hens:[],entries:[]};state=returned;await save();}
      else{state.meta.lastSyncAt=new Date().toISOString();await save();}
      setSync('✓ Synchronisé',state.meta.lastSyncAt);renderHome();renderHistory();renderHens();renderStats();
    }catch(err){setSync('! Échec');console.warn('Poulettos sync',err);toast('Synchronisation impossible : '+(err?.message||'erreur'))}
    finally{syncInProgress=false;if(btn)btn.disabled=false}
  }
  function syncSoon(){setTimeout(()=>syncNow(),0);}
  $('#syncBtn').addEventListener('click',syncNow);
  $('#profileBtn').addEventListener('click',toggleProfileMenu);
  $('#logoutBtn').addEventListener('click',logout);
  document.addEventListener('click',e=>{
    const menu=$('#profileMenu'),btn=$('#profileBtn');
    if(menu&&!menu.classList.contains('hidden')&&!menu.contains(e.target)&&!btn.contains(e.target))menu.classList.add('hidden');
  });

  window.addEventListener('pagehide',()=>{
    try{
      const serialized=JSON.stringify(state);
      localStorage.setItem(KEY,serialized);
      localStorage.setItem(CACHE_KEY,serialized);
    }catch{}
  });

  window.addEventListener('online',()=>syncNow());
  window.addEventListener('load',async()=>{
    // Restore the local database BEFORE rendering or authenticating. IndexedDB is the durable
    // offline source of truth; localStorage remains a compatibility backup.
    await restoreLocalState();
    const hasLocalUser=!!state.user?.email;
    const localCacheExists=localReady && (!!state.user?.email || state.hens.length>0 || state.entries.length>0 || !!state.meta?.lastSyncAt);
    if(localCacheExists){
      $('#loginScreen').classList.add('hidden');
      $('#app').classList.remove('hidden');
      navigate('home');
      setSync(state.meta?.lastSyncAt?'✓ Synchronisé':'● Local',state.meta?.lastSyncAt);
    }
    if(navigator.onLine) initAuth();
    else if(localCacheExists) setSync('● Hors ligne',state.meta?.lastSyncAt);
  });
})();
