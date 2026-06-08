// ─── CONFIGURAZIONE SUPABASE ──────────────────────────────────────────────
// Inserisci qui le tue credenziali Supabase (Settings → API)
// Project URL preconfigurato
var SUPABASE_URL = localStorage.getItem('sb_url') || 'https://rtyqvvjrzfjsjywxlnle.supabase.co';
var SUPABASE_KEY = localStorage.getItem('sb_key') || '';
// Salva l'URL se non era già salvato
if(!localStorage.getItem('sb_url')) localStorage.setItem('sb_url','https://rtyqvvjrzfjsjywxlnle.supabase.co');

// Client Supabase minimale (senza libreria esterna)
var SB = {
  headers: function(){
    return {
      'Content-Type': 'application/json',
      'apikey': SUPABASE_KEY,
      'Authorization': 'Bearer ' + SUPABASE_KEY,
      'Prefer': 'return=representation'
    };
  },
  url: function(table, params){
    return SUPABASE_URL + '/rest/v1/' + table + (params ? '?' + params : '');
  },
  get: async function(table, params){
    var r = await fetch(SB.url(table, params), { headers: SB.headers() });
    return r.json();
  },
  upsert: async function(table, data, conflictCol){
    var prefer = 'resolution=merge-duplicates,return=minimal';
    var url = SB.url(table) + (conflictCol ? '?on_conflict=' + conflictCol : '');
    var r = await fetch(url, {
      method: 'POST',
      headers: Object.assign({}, SB.headers(), {'Prefer': prefer}),
      body: JSON.stringify(Array.isArray(data) ? data : [data])
    });
    return r;
  },
  update: async function(table, match, data){
    var params = Object.entries(match).map(function(e){ return 'eq.'+e[0]+'=eq.'+e[1]; }).join('&');
    // Usa la sintassi corretta Supabase: ?colonna=eq.valore
    var queryParams = Object.entries(match).map(function(e){ return e[0]+'=eq.'+e[1]; }).join('&');
    var r = await fetch(SB.url(table, queryParams), {
      method: 'PATCH',
      headers: SB.headers(),
      body: JSON.stringify(data)
    });
    return r;
  },
  isConfigured: function(){
    return SUPABASE_URL && SUPABASE_KEY && SUPABASE_URL.startsWith('https://');
  }
};

function setupSupabase(){
  var key = prompt('Inserisci la Supabase anon public key\n\nDove trovarla:\n1. Vai su supabase.com → il tuo progetto\n2. Settings → API Keys\n3. Tab "Legacy anon, service_role API keys"\n4. Copia la chiave "anon public"');
  if(!key || key.length < 50) return false;
  localStorage.setItem('sb_key', key.trim());
  SUPABASE_KEY = key.trim();
  return true;
}

const EPOCH = new Date(Date.UTC(1899, 11, 30)); // EPOCH UTC esatta per Excel

// Parsa una stringa data (YYYY-MM-DD o ISO) come data locale senza offset UTC
function parseLocalDate(s) {
  if(!s) return null;
  // Se è YYYY-MM-DD (nostro formato salvato)
  var m = String(s).match(/^(\d{4})-(\d{2})-(\d{2})/);
  if(m) return new Date(parseInt(m[1]), parseInt(m[2])-1, parseInt(m[3]));
  return null;
}
let bookings = [];
let chatHistory = [];

document.addEventListener('DOMContentLoaded', function() {
  bindEvents();
  loadFromStorage();
  initPreventivi();
});

function bindEvents() {
  document.getElementById('sync-btn').addEventListener('click', doSync);
  document.getElementById('file-input').addEventListener('change', function() { loadFile(this); });
  document.querySelectorAll('.tab').forEach(function(t) {
    t.addEventListener('click', function() { showTab(t.dataset.tab); });
  });
  document.querySelectorAll('.qbtn').forEach(function(b) {
    b.addEventListener('click', function() { askQ(b.dataset.q); });
  });
  document.getElementById('chat-input').addEventListener('keydown', function(e) {
    if (e.key === 'Enter') sendMsg();
  });
  document.getElementById('send-btn').addEventListener('click', sendMsg);
  document.querySelectorAll('.rcard').forEach(function(c) {
    c.addEventListener('click', function() { openReport(c.dataset.report); });
  });
  document.getElementById('rmodal-back').addEventListener('click', closeReport);
  document.getElementById('rmodal-share').addEventListener('click', exportReport);
  document.getElementById('rmodal-print').addEventListener('click', function(){
    var title = document.getElementById('rmodal-title').textContent;
    var bodyHtml = document.getElementById('rmodal-body').innerHTML;
    printReport(title, bodyHtml);
  });
  document.getElementById('rmodal-body').addEventListener('click', function(e) {
    if (e.target.classList.contains('cb')) e.target.classList.toggle('checked');
    // Gestione badge pagamento (data-toggle-pag)
    var pag = e.target.closest ? e.target.closest('[data-toggle-pag]') : null;
    if(!pag && e.target.dataset && e.target.dataset.togglePag) pag = e.target;
    if(pag) togglePagamento(pag.getAttribute('data-toggle-pag'));
  });
  document.getElementById('rmodal-body').addEventListener('focusout', function(e) {
    var el = e.target;
    if(el && el.dataset && el.dataset.saveNote) saveNote(el.dataset.saveNote, el.value);
    if(el && el.dataset && el.dataset.saveColNote) saveColNote(el.dataset.saveColNote, el.value);
  });
  document.getElementById('rmodal-body').addEventListener('change', function(e) {
    if(e.target && e.target.id === 'cam-start-date') refreshReportCamere();
    if(e.target && e.target.id === 'col-start-date') refreshReportColazione();
  });
  document.getElementById('input-row').style.display = 'none';
  var cancelBtn = document.getElementById('file-panel-cancel');
  if (cancelBtn) cancelBtn.addEventListener('click', hideFilePanel);
  var dbBtn = document.getElementById('db-btn');
  if (dbBtn) dbBtn.addEventListener('click', configureDatabase);
  var resetApiBtn = document.getElementById('reset-api-btn');
  if (resetApiBtn) resetApiBtn.addEventListener('click', resetApiKey);
}

// ─── STORAGE ─────────────────────────────────────────────
function saveToStorage(data, lastSync) {
  try {
    localStorage.setItem('bnb_bookings', JSON.stringify(data));
    localStorage.setItem('bnb_lastSync', lastSync);
  } catch(e) {}
}

function loadFromStorage() {
  // Se Supabase è configurato, carica da lì
  if(SB.isConfigured()){
    loadFromSupabase();
    return;
  }
  // Altrimenti usa localStorage
  try {
    var raw = localStorage.getItem('bnb_bookings');
    var ls = localStorage.getItem('bnb_lastSync');
    if (raw) {
      var parsed = JSON.parse(raw);
      if (parsed && parsed.length) {
        bookings = parsed.map(function(b) {
          return Object.assign({}, b, {
            checkin: b.checkin ? parseLocalDate(b.checkin) : null,
            checkout: b.checkout ? parseLocalDate(b.checkout) : null
          });
        });
        var d = new Date(ls);
        setStatus('ok', 'Aggiornato ' + d.toLocaleTimeString('it-IT',{hour:'2-digit',minute:'2-digit'}) + ' · ' + d.toLocaleDateString('it-IT'));
        document.getElementById('app-sub').textContent = bookings.length + ' prenotazioni';
        showMainApp();
        renderAll();
        return;
      }
    }
  } catch(e) {}
  showOnboarding();
}

async function loadFromSupabase(){
  setStatus('loading', 'Caricamento dal database…');
  try {
    // Carica prenotazioni non cancellate (attive + storiche)
    var today = ds(new Date());
    var rows = await SB.get('bookings',
      'or=(checkout.gte.' + today + ',stato.eq.Attiva)&order=checkin.desc&limit=500'
    );
    if(!Array.isArray(rows) || rows.length === 0){
      // Nessun dato su Supabase → mostra onboarding
      showOnboarding();
      setStatus('loading', 'Nessuna prenotazione nel database. Carica il file Excel.');
      return;
    }

    // Carica le note
    var notes = await SB.get('notes', 'select=codice,pagamento_ok,no_show,note,note_colazione');
    var notesMap = {};
    if(Array.isArray(notes)){
      notes.forEach(function(n){ notesMap[n.codice] = n; });
    }
    localStorage.setItem('bnb_notes', JSON.stringify(
      Object.fromEntries(Object.entries(notesMap).map(function(e){
        return [e[0], {
          pagamentoOk: e[1].pagamento_ok,
          noShow: e[1].no_show,
          note: e[1].note||'',
          noteColazione: e[1].note_colazione||''
        }];
      }))
    ));

    bookings = rows.map(function(r){
      return {
        codice: r.codice, canale: r.canale||'',
        checkin: r.checkin ? parseLocalDate(r.checkin) : null,
        checkout: r.checkout ? parseLocalDate(r.checkout) : null,
        nome: r.nome||'', cognome: r.cognome||'', paese: r.paese||'',
        importo: parseFloat(r.importo)||0, commissioni: parseFloat(r.commissioni)||0,
        tassa: parseFloat(r.tassa_soggiorno)||0, camera: r.camera||'',
        adulti: parseInt(r.adulti)||0, bambini: parseInt(r.bambini)||0,
        stato: r.stato||'Attiva'
      };
    });

    var d = new Date();
    setStatus('ok', rows.length + ' pren. · DB ' + d.toLocaleTimeString('it-IT',{hour:'2-digit',minute:'2-digit'}));
    document.getElementById('app-sub').textContent = rows.length + ' prenotazioni';
    showMainApp();
    renderAll();
  } catch(e){
    setStatus('err', 'Errore DB: ' + e.message);
    // Fallback a localStorage
    loadFromStorage();
  }
}

// ─── SYNC ────────────────────────────────────────────────
function doSync() {
  // Mostra il pannello — il link al gestionale è dentro il pannello stesso
  showFilePanel();
}

function showFilePanel() {
  document.getElementById('file-panel').style.display = 'flex';
}

function hideFilePanel() {
  document.getElementById('file-panel').style.display = 'none';
}

function loadFile(input) {
  var file = input.files[0];
  if (!file) return;
  setStatus('loading', 'Elaborazione file…');
  var reader = new FileReader();
  reader.onload = function(e) {
    var u8 = new Uint8Array(e.target.result);
    parseXlsx(u8).then(function(rows) {
      processData(rows);
      hideFilePanel();
      input.value = '';
    }).catch(function(err) {
      setStatus('err', 'Errore lettura file: ' + err.message);
    });
  };
  reader.readAsArrayBuffer(file);
}

// ─── PARSE XLSX ──────────────────────────────────────────
async function parseXlsx(u8) {
  var files = await readZip(u8);
  var ssRaw = files['xl/sharedStrings.xml'];
  var strs = ssRaw ? parseSharedStrings(new TextDecoder().decode(ssRaw)) : [];
  var sheetKey = Object.keys(files).find(function(k) { return k.match(/xl\/worksheets\/sheet\d+\.xml/); });
  var sheetRaw = sheetKey ? files[sheetKey] : null;
  return sheetRaw ? parseSheet(new TextDecoder().decode(sheetRaw), strs) : [];
}

async function readZip(u8) {
  var view = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
  var eocd = -1;
  for (var i = u8.length - 22; i >= 0; i--) {
    if (u8[i]===0x50&&u8[i+1]===0x4b&&u8[i+2]===0x05&&u8[i+3]===0x06){eocd=i;break;}
  }
  if (eocd<0) return {};
  var cdCount = view.getUint16(eocd+8,true);
  var cdOffset = view.getUint32(eocd+16,true);
  var pos = cdOffset, entries = [];
  for (var e=0;e<cdCount;e++){
    if(view.getUint32(pos,true)!==0x02014b50) break;
    var method=view.getUint16(pos+10,true),compSize=view.getUint32(pos+20,true);
    var fnLen=view.getUint16(pos+28,true),exLen=view.getUint16(pos+30,true),cmLen=view.getUint16(pos+32,true);
    var lhOffset=view.getUint32(pos+42,true);
    var fname=new TextDecoder().decode(u8.slice(pos+46,pos+46+fnLen));
    entries.push({fname,method,compSize,lhOffset});
    pos+=46+fnLen+exLen+cmLen;
  }
  var files={};
  for(var entry of entries){
    var lh=entry.lhOffset,lfnLen=view.getUint16(lh+26,true),lexLen=view.getUint16(lh+28,true);
    var ds=lh+30+lfnLen+lexLen,cb=u8.slice(ds,ds+entry.compSize);
    if(entry.method===0){files[entry.fname]=cb;}
    else if(entry.method===8){try{files[entry.fname]=await inflate(cb);}catch(e){}}
  }
  return files;
}

async function inflate(data) {
  var ds = new DecompressionStream('deflate-raw');
  var w=ds.writable.getWriter(),r=ds.readable.getReader();
  w.write(data);w.close();
  var chunks=[],total=0;
  while(true){var x=await r.read();if(x.done)break;chunks.push(x.value);total+=x.value.length;}
  var out=new Uint8Array(total),off=0;
  for(var c of chunks){out.set(c,off);off+=c.length;}
  return out;
}

function parseSharedStrings(xml) {
  var strs=[],re=/<si>([\s\S]*?)<\/si>/g,m;
  while((m=re.exec(xml))!==null){
    var tRe=/<t[^>]*>([\s\S]*?)<\/t>/g,tm,val='';
    while((tm=tRe.exec(m[1]))!==null) val+=tm[1];
    strs.push(dx(val));
  }
  return strs;
}

function parseSheet(xml,strs){
  var rows=[],headers=null,rowRe=/<row\b[^>]*>([\s\S]*?)<\/row>/g,rm;
  while((rm=rowRe.exec(xml))!==null){
    var cellRe=/<c\b([^>]*)>([\s\S]*?)<\/c>/g,cm,cells={};
    while((cm=cellRe.exec(rm[1]))!==null){
      var attrs=cm[1],inner=cm[2],rM=/\br="([A-Z]+)\d+"/.exec(attrs);
      if(!rM) continue;
      var col=rM[1],vM=/<v>([\s\S]*?)<\/v>/.exec(inner),val;
      if(attrs.indexOf('t="s"')>=0&&vM) val=strs[parseInt(vM[1])]||'';
      else if(vM){var n=parseFloat(vM[1]);val=isNaN(n)?vM[1]:n;}
      else continue;
      cells[col]=val;
    }
    if(!Object.keys(cells).length) continue;
    var sorted=Object.keys(cells).sort(function(a,b){return a.length!==b.length?a.length-b.length:a<b?-1:1;});
    if(!headers){headers={};sorted.forEach(function(c){headers[c]=String(cells[c]);});}
    else{var obj={};sorted.forEach(function(c){if(headers[c])obj[headers[c]]=cells[c];});if(Object.keys(obj).length)rows.push(obj);}
  }
  return rows;
}

function dx(s){return s.replace(/&amp;/g,'&').replace(/&lt;/g,'<').replace(/&gt;/g,'>').replace(/&quot;/g,'"').replace(/&#39;/g,"'");}

// ─── PROCESS DATA ────────────────────────────────────────
function processData(rows) {
  bookings = rows.map(function(r) {
    return {
      codice: String(r['Codice']||''), canale: r['Canale']||'',
      checkin: excelDate(r['Check-in']), checkout: excelDate(r['Check-out']),
      nome: r['Nome']||'', cognome: r['Cognome']||'', paese: r['Paese']||'',
      importo: parseFloat(r['Importo'])||0, commissioni: parseFloat(r['Commissioni'])||0,
      tassa: parseFloat(r['Tassa soggiorno'])||0, camera: r['Camera']||'',
      adulti: parseInt(r['Numero adulti'])||0, bambini: parseInt(r['Numero bambini'])||0,
      stato: r['Stato']||''
    };
  });

  var now = new Date().toISOString();
  // Salva sempre in locale (funziona anche senza Supabase)
  saveToStorage(bookings.map(function(b){
    return Object.assign({},b,{checkin:b.checkin?ds(b.checkin):null,checkout:b.checkout?ds(b.checkout):null});
  }), now);

  var d = new Date(now);
  setStatus('ok','Aggiornato alle ' + d.toLocaleTimeString('it-IT',{hour:'2-digit',minute:'2-digit'}));
  document.getElementById('app-sub').textContent = bookings.length + ' prenotazioni';
  showMainApp();
  renderAll();

  // Sync con Supabase (se configurato)
  if(SB.isConfigured()){
    setStatus('loading', 'Sincronizzazione database…');
    syncToSupabase(bookings).then(function(result){
      setStatus('ok', 'Database aggiornato · ' + result);
    }).catch(function(e){
      setStatus('err', 'Sync DB fallita: ' + e.message);
    });
  }
}

async function syncToSupabase(newBookings){
  var today = ds(new Date());

  // 1. Carica dal DB le prenotazioni con checkout >= oggi
  var existing = await SB.get('bookings', 'checkout=gte.' + today + '&select=codice,stato,checkin,checkout');
  if(!Array.isArray(existing)) throw new Error('Risposta DB non valida');

  // 2. Costruisci mappa codice → prenotazione esistente
  var existingMap = {};
  existing.forEach(function(b){ existingMap[b.codice] = b; });

  // 3. Mappa dei codici nel nuovo file Excel (solo quelli con checkout >= oggi)
  var newMap = {};
  newBookings.forEach(function(b){
    if(b.checkout && ds(b.checkout) >= today) newMap[b.codice] = b;
  });

  // 4. Prenotazioni da inserire/aggiornare (nel file Excel)
  var toUpsert = newBookings.map(function(b){
    return {
      codice:         b.codice,
      canale:         b.canale,
      checkin:        b.checkin ? ds(b.checkin) : null,
      checkout:       b.checkout ? ds(b.checkout) : null,
      cognome:        b.cognome,
      nome:           b.nome,
      paese:          b.paese,
      importo:        b.importo,
      commissioni:    b.commissioni,
      tassa_soggiorno: b.tassa,
      camera:         b.camera,
      adulti:         b.adulti,
      bambini:        b.bambini,
      stato:          b.stato || 'Attiva'
    };
  });

  // 5. Prenotazioni nel DB con checkout futuro ma NON nel nuovo file → Cancellata
  var toCancelCodici = Object.keys(existingMap).filter(function(cod){
    return !newMap[cod] && existingMap[cod].stato !== 'Cancellata';
  });

  // 6. Prenotazioni già "Cancellata" che ricompaiono nel file → Attiva
  var toRestoreCodici = Object.keys(existingMap).filter(function(cod){
    return newMap[cod] && existingMap[cod].stato === 'Cancellata';
  });

  // Esegui upsert delle prenotazioni del file
  if(toUpsert.length > 0){
    // Upsert a batch di 50
    for(var i=0; i<toUpsert.length; i+=50){
      await SB.upsert('bookings', toUpsert.slice(i, i+50));
    }
  }

  // Marca come Cancellata le prenotazioni sparite
  for(var j=0; j<toCancelCodici.length; j++){
    await SB.update('bookings', {codice: toCancelCodici[j]}, {
      stato: 'Cancellata',
      stato_precedente: existingMap[toCancelCodici[j]].stato,
      cancellata_il: new Date().toISOString()
    });
  }

  // Ripristina le prenotazioni ricomparse
  for(var k=0; k<toRestoreCodici.length; k++){
    await SB.update('bookings', {codice: toRestoreCodici[k]}, { stato: 'Attiva' });
  }

  var msg = toUpsert.length + ' aggiornate';
  if(toCancelCodici.length) msg += ', ' + toCancelCodici.length + ' cancellate';
  if(toRestoreCodici.length) msg += ', ' + toRestoreCodici.length + ' ripristinate';
  return msg;
}

function excelDate(n){
  if(!n||isNaN(n))return null;
  // Convertiamo il numero seriale Excel in data locale (senza offset UTC)
  var d = new Date(EPOCH.getTime() + n * 86400000);
  // Creiamo una data usando solo anno/mese/giorno locali per evitare offset fuso orario
  return new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate());
}
function fmtDate(d){if(!d)return'—';return d.toLocaleDateString('it-IT',{day:'2-digit',month:'2-digit',year:'numeric'});}
function ds(d){
  if(!d||!(d instanceof Date))return'';
  // Usa ora LOCALE (non UTC) per evitare offset fuso orario
  var y=d.getFullYear(), m=String(d.getMonth()+1).padStart(2,'0'), day=String(d.getDate()).padStart(2,'0');
  return y+'-'+m+'-'+day;
}
// ─── UI ──────────────────────────────────────────────────
function showOnboarding(){
  document.getElementById('onboarding').style.display='flex';
  document.getElementById('main-app').style.display='none';
}
function showMainApp(){
  document.getElementById('onboarding').style.display='none';
  var ma=document.getElementById('main-app');
  ma.style.display='flex';ma.style.flexDirection='column';ma.style.flex='1';ma.style.overflow='hidden';
}
function setStatus(type,msg){
  var bar=document.getElementById('status-bar');
  bar.className=type;
  document.getElementById('status-text').textContent=msg;
}
function showTab(name){
  document.querySelectorAll('.tab').forEach(function(t){t.classList.toggle('active',t.dataset.tab===name);});
  document.querySelectorAll('.tab-content').forEach(function(c){c.classList.remove('active');});
  var tc=document.getElementById('tab-'+name);if(tc)tc.classList.add('active');
  document.getElementById('input-row').style.display=name==='chat'?'flex':'none';
  if(name==='chat')setTimeout(scrollChat,50);
}
function renderAll(){renderStats();renderToday();}
function renderStats(){
  var today=ds(new Date()),now=new Date();now.setHours(0,0,0,0);
  var we=new Date(now);we.setDate(we.getDate()+7);
  var ins=bookings.filter(function(b){return ds(b.checkin)===today;});
  var outs=bookings.filter(function(b){return ds(b.checkout)===today;});
  var wk=bookings.filter(function(b){return b.checkin>=now&&b.checkin<we;});
  var active=bookings.filter(function(b){return b.stato==='Attiva'||b.stato==='Modificata';});
  var rev=active.reduce(function(s,b){return s+b.importo;},0);
  set('s-ci',ins.length);set('s-ci-n',ins.length?ins.map(function(b){return b.cognome;}).join(', '):'—');
  set('s-co',outs.length);set('s-co-n',outs.length?outs.map(function(b){return b.cognome;}).join(', '):'—');
  set('s-wk',wk.length);set('s-wk-n',wk.length?wk.map(function(b){return b.cognome;}).join(', '):'—');
  set('s-rev','€'+Math.round(rev).toLocaleString('it-IT'));set('s-rev-n',active.length+' prenotazioni');
}
function renderToday(){
  var today=ds(new Date());
  var ins=bookings.filter(function(b){return ds(b.checkin)===today;});
  var outs=bookings.filter(function(b){return ds(b.checkout)===today;});
  var list=document.getElementById('today-list');
  if(!ins.length&&!outs.length){list.innerHTML='<div class="empty">Nessun movimento oggi</div>';return;}
  var html='';
  ins.forEach(function(b){html+=brow(b,'in');});
  outs.forEach(function(b){html+=brow(b,'out');});
  list.innerHTML=html;
}
function brow(b,type){
  var ini=((b.nome||'')[0]||'').toUpperCase()+((b.cognome||'')[0]||'').toUpperCase();
  var av=type==='in'?'av-in':'av-out';
  var badge=type==='in'?'<span class="bbadge badge-in">Check-in</span>':'<span class="bbadge badge-out">Check-out</span>';
  return '<div class="brow"><div class="bavatar '+av+'">'+ini+'</div><div class="binfo"><div class="bname">'+b.nome+' '+b.cognome+'</div><div class="broom">'+b.camera+' · '+b.adulti+' adulti</div></div>'+badge+'</div>';
}
function set(id,val){var el=document.getElementById(id);if(el)el.textContent=val;}

// ─── CHAT ────────────────────────────────────────────────
function buildCtx(){
  var today=ds(new Date());
  return 'Oggi è '+today+'.\n\nPrenotazioni:\n'+bookings.map(function(b){
    return '- '+b.nome+' '+b.cognome+' | '+b.canale+' | CI:'+fmtDate(b.checkin)+' | CO:'+fmtDate(b.checkout)+' | '+b.camera+' | '+b.adulti+' adulti | €'+b.importo+' | comm:€'+b.commissioni+' | tassa:€'+b.tassa+' | '+b.stato+' | '+b.paese;
  }).join('\n');
}
function addMsg(role,text){
  var c=document.getElementById('messages'),d=document.createElement('div');
  d.className='msg '+(role==='user'?'user':'ai');
  d.innerHTML=role==='ai'?'<div class="micon">🤖</div><div class="bubble">'+text.replace(/\n/g,'<br>')+'</div>':'<div class="bubble">'+text+'</div>';
  c.appendChild(d);scrollChat();
}
function scrollChat(){var c=document.getElementById('messages');if(c)c.scrollTop=c.scrollHeight;}
function showTyping(){
  var c=document.getElementById('messages'),d=document.createElement('div');
  d.id='tdot';d.className='msg ai';
  d.innerHTML='<div class="micon">🤖</div><div class="bubble"><div class="typing"><span></span><span></span><span></span></div></div>';
  c.appendChild(d);scrollChat();
}
function removeTyping(){var t=document.getElementById('tdot');if(t)t.remove();}
function getApiKey(){
  var k = localStorage.getItem('bnb_apikey');
  if(!k){
    k = prompt('Inserisci la tua API key Groq (GRATUITA):\n\n1. Vai su console.groq.com\n2. Registrati gratis\n3. Crea una API Key\n4. Incollala qui\n\nViene salvata solo sul tuo dispositivo.');
    if(k) localStorage.setItem('bnb_apikey', k.trim());
  }
  return k ? k.trim() : null;
}

function sendMsg(){
  if(!bookings.length){addMsg('ai','Prima sincronizza i dati.');return;}
  var inp=document.getElementById('chat-input'),text=inp.value.trim();
  if(!text)return;
  var apiKey = getApiKey();
  if(!apiKey){addMsg('ai','API key non inserita. Riprova e inserisci la chiave.');return;}
  inp.value='';inp.blur();
  addMsg('user',text);
  chatHistory.push({role:'user',content:text});
  showTyping();
  var sys='Sei un assistente esperto per un B&B italiano. Rispondi in italiano, conciso e pratico. Date: gg/mm/aaaa. Importi con €.\n\n'+buildCtx();
  // Usa Groq API (gratuita)
  var msgs = [{role:'system',content:sys}].concat(chatHistory.slice(-12));
  fetch('https://api.groq.com/openai/v1/chat/completions',{
    method:'POST',
    headers:{'Content-Type':'application/json','Authorization':'Bearer '+apiKey},
    body:JSON.stringify({model:'llama-3.3-70b-versatile',max_tokens:1000,messages:msgs})
  }).then(function(r){return r.json();}).then(function(data){
    removeTyping();
    if(data.error){
      addMsg('ai','Errore API: '+data.error.message);
      if(data.error.code==='invalid_api_key'||data.error.code==='401') localStorage.removeItem('bnb_apikey');
      return;
    }
    var reply=data.choices&&data.choices[0]?data.choices[0].message.content:'Nessuna risposta.';
    addMsg('ai',reply);chatHistory.push({role:'assistant',content:reply});
  }).catch(function(e){removeTyping();addMsg('ai','Errore di connessione: '+e.message);});
}
function askQ(q){document.getElementById('chat-input').value=q;showTab('chat');sendMsg();}
function togglePagamento(codice){
  var notes = JSON.parse(localStorage.getItem('bnb_notes')||'{}');
  if(!notes[codice]) notes[codice] = {};
  notes[codice].pagamentoOk = !notes[codice].pagamentoOk;
  var ok = notes[codice].pagamentoOk;
  localStorage.setItem('bnb_notes', JSON.stringify(notes));

  // Aggiorna badge
  document.querySelectorAll('[id="badge-pag-'+codice+'"]').forEach(function(el){
    el.textContent = ok ? '✓ Regolare' : '⚠ Da regolarizzare';
    el.className = 'rsbadge ' + (ok ? 'b-ok' : 'b-ko');
  });

  // Aggiorna colori quota/tassa nella card (arrivi e partenze)
  var card = document.getElementById('rec-'+codice);
  if(card){
    card.querySelectorAll('.rec-val span[style*="color"]').forEach(function(el){
      el.style.color = ok ? 'var(--accent-d)' : 'var(--coral-d)';
      el.textContent = el.textContent.replace(ok ? 'Da Incassare' : 'Saldato', ok ? 'Saldato' : 'Da Incassare');
    });
  }

  // Ricalcola totali reception in tempo reale
  var totBox = document.getElementById('totali-reception');
  if(totBox){
    var today = new Date(); today.setHours(0,0,0,0);
    var todayStr = ds(today);
    var cis = bookings.filter(function(b){return ds(b.checkin)===todayStr&&(b.stato==='Attiva'||b.stato==='Modificata');});
    var cos = bookings.filter(function(b){return ds(b.checkout)===todayStr&&(b.stato==='Attiva'||b.stato==='Modificata');});
    var inCasaT = bookings.filter(function(b){return b.checkin<today&&b.checkout>today&&ds(b.checkin)!==todayStr&&(b.stato==='Attiva'||b.stato==='Modificata');});
    var tuttiCodici = {};
    cis.forEach(function(b){ tuttiCodici[b.codice]=b; });
    cos.forEach(function(b){ tuttiCodici[b.codice]=b; });
    inCasaT.forEach(function(b){ tuttiCodici[b.codice]=b; });
    var freshNotes = JSON.parse(localStorage.getItem('bnb_notes')||'{}');
    // Ricalcola totali inline
    var _tutti = Object.values(tuttiCodici);
    var _da_inc = _tutti.filter(function(b){ return !(freshNotes[b.codice]||{}).pagamentoOk; });
    var _sald = _tutti.filter(function(b){ return (freshNotes[b.codice]||{}).pagamentoOk; });
    var _totQ = _da_inc.reduce(function(s,b){return s+b.importo;},0);
    var _totT = _da_inc.reduce(function(s,b){return s+calcTassa(b);},0);
    var _newHtml = '<div class="totali" id="totali-reception" style="margin-top:14px">' +
      '<h3>Totali reception oggi</h3>' +
      (_sald.length ? '<div class="tot-row" style="color:var(--accent-d)"><span>✓ Ospiti saldati</span><span>'+_sald.length+'</span></div>' : '') +
      '<div class="tot-row"><span>Ospiti da regolarizzare</span><span>'+_da_inc.length+'</span></div>' +
      '<div class="tot-row"><span>Quota soggiorni</span><span>€'+_totQ.toFixed(2)+'</span></div>' +
      '<div class="tot-row"><span>Tasse soggiorno</span><span>€'+_totT.toFixed(2)+'</span></div>' +
      '<div class="tot-row" style="font-weight:700;border-top:1px solid rgba(0,0,0,.1);margin-top:6px;padding-top:6px"><span>Totale da incassare</span><span style="color:var(--coral-d)">€'+(_totQ+_totT).toFixed(2)+'</span></div>' +
    '</div>';
    totBox.outerHTML = _newHtml;
  }

  if(SB.isConfigured()){
    SB.upsert('notes', {codice: codice,
      pagamento_ok: ok,
      note: (notes[codice]||{}).note||'',
      note_colazione: (notes[codice]||{}).noteColazione||''
    }, 'codice').catch(function(){});
  }
}

function saveNote(codice, valore){
  var notes = JSON.parse(localStorage.getItem('bnb_notes')||'{}');
  if(!notes[codice]) notes[codice] = {};
  var val = valore !== undefined ? valore : '';
  notes[codice].note = val;
  localStorage.setItem('bnb_notes', JSON.stringify(notes));
  var ta = document.querySelector('[data-save-note="'+codice+'"]');
  if(ta){ ta.style.fontWeight = val.trim() ? '700' : ''; ta.style.color = val.trim() ? '#dc2626' : ''; }
  // Sync Supabase
  if(SB.isConfigured()){
    SB.upsert('notes', {codice: codice, note: val,
      pagamento_ok: !!(notes[codice]||{}).pagamentoOk,
      note_colazione: (notes[codice]||{}).noteColazione||''
    }, 'codice').catch(function(){});
  }
}

function saveColNote(codice, valore){
  var notes = JSON.parse(localStorage.getItem('bnb_notes')||'{}');
  if(!notes[codice]) notes[codice] = {};
  var val = valore !== undefined ? valore : '';
  notes[codice].noteColazione = val;
  localStorage.setItem('bnb_notes', JSON.stringify(notes));
  var ta = document.querySelector('[data-save-col-note="'+codice+'"]');
  if(ta){ ta.style.fontWeight = val.trim() ? '700' : ''; ta.style.color = val.trim() ? '#dc2626' : ''; }
  if(SB.isConfigured()){
    SB.upsert('notes', {codice: codice, note_colazione: val,
      pagamento_ok: !!(notes[codice]||{}).pagamentoOk,
      note: (notes[codice]||{}).note||''
    }, 'codice').catch(function(){});
  }
}

function resetApiKey(){localStorage.removeItem('bnb_apikey');addMsg('ai','API key rimossa. Alla prossima domanda ti verrà chiesta di nuovo.');}

function configureDatabase(){
  if(SB.isConfigured()){
    var confirm = window.confirm('Database già configurato.\nVuoi riconfigurarlo?');
    if(!confirm) return;
  }
  if(setupSupabase()){
    alert('Database configurato! Al prossimo caricamento del file Excel i dati verranno sincronizzati.');
    loadFromSupabase();
  }
}


// ─── PREVENTIVI ───────────────────────────────────────────────────────────

function initPreventivi(){
  document.getElementById('btn-nuovo-preventivo').addEventListener('click', openNuovoPreventivo);
  document.getElementById('btn-lista-preventivi').addEventListener('click', openListaPreventivi);
  document.getElementById('prev-back').addEventListener('click', closePreventivo);
  document.getElementById('pv-add-cam').addEventListener('click', addCameraBlock);
  document.getElementById('pv-open-site').addEventListener('click', openSiteConDate);
  document.getElementById('pv-genera').addEventListener('click', generaPreventivo);
  document.getElementById('pv-condividi').addEventListener('click', condividiPreventivo);

  // Scadenza default: +7 giorni
  var d = new Date(); d.setDate(d.getDate()+7);
  document.getElementById('pv-scadenza').value = ds(d);
}

var pvCamereCount = 0;
var ultimoIdGenerato = null;

function openNuovoPreventivo(){
  document.getElementById('prev-modal').classList.add('open');
  document.getElementById('prev-title').textContent = 'Nuovo Preventivo';
  document.getElementById('pv-link-box').style.display = 'none';
  pvCamereCount = 0;
  document.getElementById('pv-camere-list').innerHTML = '';
  addCameraBlock();
}

function openListaPreventivi(){
  document.getElementById('prev-modal').classList.add('open');
  document.getElementById('prev-title').textContent = 'Preventivi inviati';
  document.getElementById('prev-scroll').innerHTML = '<div style="padding:20px;text-align:center;color:var(--text2);font-size:14px">Caricamento…</div>';
  loadListaPreventivi();
}

function closePreventivo(){
  document.getElementById('prev-modal').classList.remove('open');
  // Ripristina scroll se era lista
  if(document.getElementById('prev-title').textContent === 'Preventivi inviati'){
    location.reload();
  }
}

async function loadListaPreventivi(){
  if(!SB.isConfigured()){ 
    document.getElementById('prev-scroll').innerHTML = '<div style="padding:20px;color:var(--text2);font-size:14px">Configura Supabase per vedere i preventivi.</div>';
    return;
  }
  var rows = await SB.get('preventivi','select=id,nome_ospite,checkin,checkout,scadenza,creato_il&order=creato_il.desc&limit=30');
  var oggi = ds(new Date());
  var html = '<div style="padding:0 0 16px">';
  if(!rows || !rows.length){
    html += '<div style="padding:20px;text-align:center;color:var(--text2);font-size:14px">Nessun preventivo ancora.</div>';
  } else {
    rows.forEach(function(r){
      var scaduto = r.scadenza && r.scadenza < oggi;
      var url = location.origin + location.pathname.replace('index.html','') + 'preventivo/?id=' + r.id;
      html += '<div style="background:var(--bg2);border-radius:var(--r);padding:14px;margin-bottom:10px;border-left:3px solid '+(scaduto?'var(--border2)':'var(--accent)')+'">' +
        '<div style="font-size:15px;font-weight:600;margin-bottom:4px">' + r.nome_ospite + '</div>' +
        '<div style="font-size:12px;color:var(--text2)">' + fmtDate(new Date(r.checkin+'T00:00:00')) + ' → ' + fmtDate(new Date(r.checkout+'T00:00:00')) + '</div>' +
        '<div style="font-size:11px;color:'+(scaduto?'var(--coral-d)':'var(--accent-d)')+';margin-top:4px">' + (scaduto ? '⏰ Scaduto' : '✓ Attivo') + (r.scadenza ? ' · fino al ' + r.scadenza : '') + '</div>' +
        '<button data-prev-id="' + r.id + '" style="margin-top:8px;padding:7px 14px;border-radius:20px;border:0.5px solid var(--border2);background:var(--bg);font-size:12px;cursor:pointer;font-family:inherit">📤 Condividi</button>' +
      '</div>';
    });
  }
  html += '</div>';
  // Aggiungi bottone nuovo preventivo in fondo
  html += '<button id="pv-nuovo-dalla-lista" style="width:100%;padding:13px;border-radius:var(--r);border:none;background:var(--accent);color:#fff;font-size:14px;font-weight:600;cursor:pointer;font-family:inherit">+ Nuovo Preventivo</button>';
  document.getElementById('prev-scroll').innerHTML = html;

  // Bind condividi buttons
  document.querySelectorAll('[data-prev-id]').forEach(function(btn){
    btn.addEventListener('click', function(){
      var id = btn.dataset.prevId;
      var url = location.origin + location.pathname.replace(/\/[^\/]*$/, '') + '/preventivo/?id=' + id;
      condividiUrl(url, 'Preventivo Le Stanze dei Tesori');
    });
  });
  var btnNuovo = document.getElementById('pv-nuovo-dalla-lista');
  if(btnNuovo) btnNuovo.addEventListener('click', openNuovoPreventivo);
}

function addCameraBlock(){
  pvCamereCount++;
  var idx = pvCamereCount;
  var div = document.createElement('div');
  div.className = 'cam-block';
  div.id = 'cam-block-' + idx;
  div.innerHTML = '<div class="cam-block-hdr">' +
    '<span class="cam-block-title">🛏 Camera ' + idx + '</span>' +
    (idx > 1 ? '<button class="cam-del-btn" data-cam-del="' + idx + '">✕ Rimuovi</button>' : '') +
  '</div>' +
  '<div class="prev-field"><label class="prev-label">Nome camera</label>' +
    '<select class="prev-input" id="cam-nome-' + idx + '">' +
      '<option>1.Porta Carini</option>' +
      '<option>2.Porta S. Agata</option>' +
      '<option>3.Porta Reale</option>' +
      '<option>4.Suite Deluxe Con Vasca Idromassaggio</option>' +
    '</select></div>' +
  '<div class="prev-input-row">' +
    '<div class="prev-field"><label class="prev-label">Prezzo soggiorno (€)</label><input class="prev-input" id="cam-prezzo-' + idx + '" type="number" min="0" placeholder="es. 180"></div>' +
    '<div class="prev-field"><label class="prev-label">Tassa soggiorno (€)</label><input class="prev-input" id="cam-tassa-' + idx + '" type="number" min="0" placeholder="auto" id="cam-tassa-' + idx + '"></div>' +
  '</div>';
  document.getElementById('pv-camere-list').appendChild(div);

  // Bind rimuovi
  div.querySelectorAll('[data-cam-del]').forEach(function(btn){
    btn.addEventListener('click', function(){
      document.getElementById('cam-block-' + btn.dataset.camDel).remove();
    });
  });

  // Auto-calcolo tassa
  var prezzoInput = document.getElementById('cam-prezzo-' + idx);
  var tassaInput  = document.getElementById('cam-tassa-' + idx);
  function autoTassa(){
    var adulti = parseInt(document.getElementById('pv-adulti').value)||1;
    var ci = document.getElementById('pv-checkin').value;
    var co = document.getElementById('pv-checkout').value;
    if(ci && co){
      var n = Math.round((new Date(co) - new Date(ci))/86400000);
      tassaInput.value = adulti * Math.min(n,4) * 4;
    }
  }
  document.getElementById('pv-checkin').addEventListener('change', autoTassa);
  document.getElementById('pv-checkout').addEventListener('change', autoTassa);
  document.getElementById('pv-adulti').addEventListener('change', autoTassa);
}

function openSiteConDate(){
  var ci = document.getElementById('pv-checkin').value;
  var co = document.getElementById('pv-checkout').value;
  var url = 'https://www.bed-and-breakfast.it/it/booking/sicilia/le-stanze-dei-tesori-palermo/10505';
  if(ci && co) url += '?checkin=' + ci + '&checkout=' + co;
  window.open(url, '_blank');
}

function genId(){
  // Genera ID univoco tipo "a1b2c-d3e4f-g5h6i"
  var s = '';
  var chars = 'abcdefghijklmnopqrstuvwxyz0123456789';
  for(var i=0;i<15;i++){
    if(i===5||i===10) s+='-';
    s += chars[Math.floor(Math.random()*chars.length)];
  }
  return s;
}

async function generaPreventivo(){
  var nome = document.getElementById('pv-nome').value.trim();
  var checkin  = document.getElementById('pv-checkin').value;
  var checkout = document.getElementById('pv-checkout').value;
  var adulti   = parseInt(document.getElementById('pv-adulti').value)||1;
  var bambini  = parseInt(document.getElementById('pv-bambini').value)||0;
  var scadenza = document.getElementById('pv-scadenza').value;
  var canc     = document.getElementById('pv-cancellazione').value;
  var note     = document.getElementById('pv-note').value.trim();
  var ctaLabel = document.getElementById('pv-cta-label').value.trim() || 'Prenota ora';

  if(!nome){ toast('Inserisci il nome dell'ospite'); return; }
  if(!checkin || !checkout){ toast('Inserisci le date del soggiorno'); return; }

  // Raccogli camere
  var camere = [];
  document.querySelectorAll('.cam-block').forEach(function(block){
    var idx = block.id.replace('cam-block-','');
    var nomeEl = document.getElementById('cam-nome-'+idx);
    var prezzoEl = document.getElementById('cam-prezzo-'+idx);
    var tassaEl = document.getElementById('cam-tassa-'+idx);
    if(!nomeEl) return;
    var prezzo = parseFloat(prezzoEl?.value)||0;
    if(prezzo === 0){ toast('Inserisci il prezzo per ogni camera'); return; }
    camere.push({
      nome: nomeEl.value,
      prezzo: prezzo,
      tassa: parseFloat(tassaEl?.value)||0,
      descrizione: descrizioneCamera(nomeEl.value),
      servizi: serviziCamera(nomeEl.value)
    });
  });
  if(!camere.length){ toast('Aggiungi almeno una camera'); return; }

  var id = genId();
  var ctaUrl = 'https://www.bed-and-breakfast.it/it/booking/sicilia/le-stanze-dei-tesori-palermo/10505?checkin=' + checkin + '&checkout=' + checkout;

  var data = {id, nome_ospite:nome, checkin, checkout, adulti, bambini,
    camere: JSON.stringify(camere), note, cancellazione:canc,
    scadenza: scadenza||null, cta_url:ctaUrl, cta_label:ctaLabel, creato_da:'app'};

  if(SB.isConfigured()){
    document.getElementById('pv-genera').textContent = 'Salvataggio…';
    var r = await SB.upsert('preventivi', data, 'id');
    document.getElementById('pv-genera').textContent = '✨ Genera preventivo';
    if(!r.ok){ toast('Errore salvataggio: ' + r.status); return; }
  } else {
    // Fallback: salva in localStorage
    var prevs = JSON.parse(localStorage.getItem('bnb_preventivi')||'[]');
    prevs.unshift(data);
    localStorage.setItem('bnb_preventivi', JSON.stringify(prevs.slice(0,50)));
  }

  ultimoIdGenerato = id;
  var baseUrl = location.origin + location.pathname.replace(/[^\/]*$/, '') + 'preventivo/?id=' + id;
  document.getElementById('pv-link-url').textContent = baseUrl;
  document.getElementById('pv-link-box').style.display = 'block';
  document.getElementById('prev-scroll').scrollTo({top: document.getElementById('prev-scroll').scrollHeight, behavior:'smooth'});
  toast('✓ Preventivo creato!');
}

function condividiPreventivo(){
  if(!ultimoIdGenerato) return;
  var url = location.origin + location.pathname.replace(/[^\/]*$/, '') + 'preventivo/?id=' + ultimoIdGenerato;
  var nome = document.getElementById('pv-nome').value.trim();
  condividiUrl(url, 'Preventivo per ' + nome + ' — Le Stanze dei Tesori');
}

function condividiUrl(url, titolo){
  var testo = titolo + '\n' + url;
  if(navigator.share){
    navigator.share({title: titolo, text: testo, url: url}).catch(function(){
      copyToClipboard(url);
    });
  } else {
    copyToClipboard(url);
  }
}

function copyToClipboard(text){
  navigator.clipboard.writeText(text).then(function(){
    toast('Link copiato!');
  }).catch(function(){
    toast('Copia: ' + text);
  });
}

function descrizioneCamera(nome){
  var desc = {
    '1.Porta Carini': 'Camera matrimoniale elegante con balcone. Arredata in stile siciliano con pavimenti in cotto, letto king size, bagno en-suite con doccia e bidet.',
    '2.Porta S. Agata': 'Spaziosa camera matrimoniale con vista sui tetti di Palermo. Arredi curati, aria condizionata, TV smart e bagno privato.',
    '3.Porta Reale': 'Camera matrimoniale con balcone panoramico. Ideale per chi vuole godere dell'autenticità del centro storico palermitano.',
    '4.Suite Deluxe Con Vasca Idromassaggio': 'Suite di lusso con vasca idromassaggio privata. L'esperienza più esclusiva delle Stanze dei Tesori, per un soggiorno indimenticabile.'
  };
  return desc[nome] || 'Camera matrimoniale con bagno privato, aria condizionata e tutti i comfort.';
}

function serviziCamera(nome){
  var base = ['WiFi gratuito','Colazione inclusa','Aria condizionata','Bagno privato','TV Smart','Bidet'];
  if(nome.includes('Suite')) base.push('Vasca idromassaggio','Prodotti da bagno premium','Vista panoramica');
  if(nome.includes('Carini') || nome.includes('Reale')) base.push('Balcone');
  return base;
}

// ─── REPORTS ─────────────────────────────────────────────
function openReport(type){
  if(!bookings.length){alert('Sincronizza prima i dati.');return;}
  var modal=document.getElementById('rmodal'),body=document.getElementById('rmodal-body'),title=document.getElementById('rmodal-title');
  modal.classList.add('open');
  var today=new Date();today.setHours(0,0,0,0);
  var todayStr=ds(today),label=today.toLocaleDateString('it-IT',{weekday:'long',day:'numeric',month:'long',year:'numeric'});
  var map={colazione:['Colazioni',reportColazione],reception:['Reception',reportReception],camere:['Report Camere',reportCamere],settimanale:['Riepilogo settimanale',reportSettimanale],fabiola:['Report Fabiola',reportFabiola]};
  if(map[type]){title.textContent=map[type][0];body.innerHTML=map[type][1](today,todayStr,label);}
}
function closeReport(){document.getElementById('rmodal').classList.remove('open');}
function printReport(){
  var b=document.getElementById('rmodal-body').innerHTML,t=document.getElementById('rmodal-title').textContent;
  var w=window.open('','_blank');
  w.document.write('<!DOCTYPE html><html><head><title>'+t+'</title><style>body{font-family:sans-serif;padding:20px;font-size:13px}.rstitle{font-weight:600;font-size:11px;text-transform:uppercase;color:#666;margin:14px 0 6px;border-bottom:1px solid #ddd;padding-bottom:4px}.rsrow{display:flex;justify-content:space-between;padding:6px 0;border-bottom:1px solid #eee}.rsbadge{padding:2px 8px;border-radius:10px;font-size:11px}.b-bk{background:#E6F1FB;color:#185FA5}.b-ok{background:#E1F5EE;color:#0F6E56}.b-ko{background:#FAECE7;color:#993C1D}.totbox{background:#f5f5f2;padding:10px;margin-top:8px}.totrow{display:flex;justify-content:space-between;padding:3px 0}.totrow.main{font-weight:700;border-top:1px solid #ddd;margin-top:6px;padding-top:6px}</style></head><body><h2>'+t+'</h2>'+b+'</body></html>');
  w.document.close();w.print();
}

// ─── HELPERS REPORT ───────────────────────────────────────────────────────

function getCanale(b){
  var c = (b.canale||'').toLowerCase();
  var n = ((b.nome||'')+(b.cognome||'')).toLowerCase();
  if(n.includes('air bb')||n.includes('airbnb')) return 'Airbnb';
  if(c.includes('booking engine')||c.includes('sitoweb')||c.includes('sito web')) return 'Sito Web';
  if(c.includes('booking')) return 'Booking.com';
  if(c.includes('ireservation')||c.includes('diretto')) return 'Diretto';
  return b.canale||'Diretto';
}

function getCanaleLabel(b){
  var c = getCanale(b);
  var cls = {'Booking.com':'b-bk','Sito Web':'b-si','Airbnb':'b-ab','Diretto':'b-dir'};
  return '<span class="rsbadge '+(cls[c]||'b-dir')+'">'+ c +'</span>';
}

function getNotti(b){
  if(!b.checkin||!b.checkout) return 0;
  return Math.round((b.checkout - b.checkin)/86400000);
}

function calcTassa(b){
  // 4€/adulto/notte max 4 notti; bambini ≥12 anni pagano (non sappiamo età, nota)
  var notti = Math.min(getNotti(b), 4);
  return b.adulti * notti * 4;
}

function dotazioniCamera(b, refDay){
  var notti = getNotti(b);
  var checkinDay = Math.round((refDay - b.checkin)/86400000); // 0=giorno arrivo
  var isCheckout = ds(b.checkout) === ds(refDay);
  var isCheckin  = ds(b.checkin)  === ds(refDay);

  if(isCheckout || isCheckin){
    return '<div class="dotaz">🔄 <strong>Cambio totale 100%</strong><br>' +
           '<span class="dotaz-items">2 lenzuola · 2 federe · 2 teli doccia · 2 teli viso · 2 teli bidet · 1 tappetino · 2 spazzolini · 2 cuffie doccia</span></div>';
  }

  var giornoFermata = checkinDay;
  // È l'ultima notte se domani è il checkout
  var domani = new Date(refDay); domani.setDate(domani.getDate()+1);
  var isUltimaNotte = ds(domani) === ds(b.checkout);

  var cambioLetto = giornoFermata > 0 && giornoFermata % 4 === 0;
  var cambioBagno = giornoFermata > 0 && giornoFermata % 2 === 0;

  // Caso speciale: cambio previsto MA è l'ultima notte
  if((cambioLetto || cambioBagno) && isUltimaNotte){
    var cosa = cambioLetto && cambioBagno ? 'Lenzuola e Set Bagno' :
               cambioLetto ? 'Lenzuola' : 'Set Bagno';
    return '<div class="dotaz" style="border-left:3px solid #f59e0b;padding-left:8px">' +
      '<strong style="color:#92400e">⚠ Cambio '+cosa+' — ULTIMA NOTTE<br>' +
      'Cambiare solo se in pessime condizioni per evitare sprechi</strong></div>';
  }

  if(!cambioLetto && !cambioBagno){
    return '<div class="dotaz">✨ Solo riassetto — nessun cambio biancheria oggi</div>';
  }

  var items = [];
  if(cambioLetto) items.push('🛏 Set Letto: 2 lenzuola · 2 federe');
  if(cambioBagno) items.push('🚿 Set Bagno: 2 teli doccia · 2 teli viso · 2 teli bidet · 1 tappetino');
  return '<div class="dotaz">' + items.join('<br>') + '</div>';
}


// ─── EXPORT / CONDIVISIONE REPORT ─────────────────────────────────────────

function buildStandaloneHtml(title, bodyHtml){
  var styles = `
    body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',sans-serif;padding:16px;font-size:13px;max-width:800px;margin:0 auto;color:#1a1a18;background:#fff}
    h1{font-size:17px;margin-bottom:4px;color:#0F6E56}
    h2{font-size:14px;font-weight:600;margin:18px 0 8px;color:#0F6E56;border-bottom:2px solid #1D9E75;padding-bottom:5px}
    .rec-card{border:1px solid #ddd;border-radius:10px;padding:12px;margin-bottom:12px}
    .checkout-card{opacity:.9}
    .rec-header{display:flex;gap:8px;flex-wrap:wrap;margin-bottom:8px;align-items:center}
    .rec-name{font-size:15px;font-weight:700;flex:1}
    .rec-grid{display:grid;grid-template-columns:1fr 1fr;gap:6px;margin-bottom:8px}
    .rec-item{display:flex;flex-direction:column;gap:1px}
    .rec-lbl{font-size:10px;color:#888;text-transform:uppercase;letter-spacing:.04em}
    .rec-val{font-size:13px;font-weight:500}
    .rsbadge{font-size:11px;padding:2px 8px;border-radius:10px;font-weight:500;display:inline-block}
    .b-bk{background:#E6F1FB;color:#185FA5}.b-si{background:#FAEEDA;color:#854F0B}
    .b-ab{background:#f5e9fb;color:#7c3aed}.b-dir{background:#fef3c7;color:#92400e}
    .b-ok{background:#E1F5EE;color:#0F6E56}.b-ko{background:#FAECE7;color:#993C1D}
    .b-out{background:#FAECE7;color:#993C1D}
    .cam-day{margin-bottom:14px}
    .cam-day-hdr{font-weight:700;font-size:13px;border-bottom:2px solid #1D9E75;padding-bottom:4px;margin-bottom:8px;color:#0F6E56}
    .cam-row{padding:7px 0;border-bottom:1px solid #eee}
    .cam-row-hdr{display:flex;align-items:center;gap:8px;margin-bottom:2px}
    .cam-nome{font-weight:600}
    .dotaz{font-size:11px;margin-top:5px;padding:6px 8px;background:#f5f5f2;border-radius:6px;color:#555}
    .dotaz-items{font-size:11px;color:#888}
    .totali{background:#f0f9f5;border:1px solid #1D9E75;border-radius:8px;padding:12px;margin-top:14px}
    .totali h3{margin:0 0 8px;font-size:13px;font-weight:600;color:#0F6E56}
    .tot-row{display:flex;justify-content:space-between;font-size:13px;padding:3px 0}
    .nodata{font-size:13px;color:#aaa;text-align:center;padding:16px 0}
    table{width:100%;border-collapse:collapse;margin-bottom:8px}
    th{background:#f0f9f5;font-size:11px;padding:5px 7px;text-align:left;border:1px solid #ddd}
    td{font-size:12px;padding:5px 7px;border:1px solid #ddd}
    .print-bar{display:flex;gap:8px;margin-bottom:16px;flex-wrap:wrap}
    .btn{padding:9px 16px;border:none;border-radius:8px;cursor:pointer;font-size:13px;font-weight:500}
    .btn-print{background:#1D9E75;color:#fff}.btn-share{background:#25D366;color:#fff}
    .rdate{font-size:11px;color:#888;margin-bottom:10px}
    textarea{display:none}
    @media print{.print-bar{display:none}}
  `;

  var now = new Date().toLocaleString('it-IT');
  var shareScript = `
    function doShare(){
      if(navigator.share){
        var htmlContent = document.documentElement.outerHTML;
        var blob = new Blob([htmlContent],{type:'text/html'});
        var file = new File([blob],'report-bnb.html',{type:'text/html'});
        if(navigator.canShare && navigator.canShare({files:[file]})){
          navigator.share({files:[file],title:'${title}',text:'Report B&B del ${now}'})
            .catch(function(){doDownload();});
        } else {
          navigator.share({title:'${title}',text:'Report B&B del ${now} - aprire il file allegato'})
            .catch(function(){doDownload();});
        }
      } else { doDownload(); }
    }
    function doDownload(){
      var htmlContent = document.documentElement.outerHTML;
      var blob = new Blob([htmlContent],{type:'text/html'});
      var a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = 'report-bnb-${now.replace(/[/:, ]/g,"-")}.html';
      a.click();
    }
  `;

  return '<!DOCTYPE html><html lang="it"><head>' +
    '<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">' +
    '<title>'+title+'</title>' +
    '<style>'+styles+'</style>' +
    '<script>'+shareScript+'<\/script>' +
    '</head><body>' +
    '<div class="print-bar">' +
      '<button class="btn btn-print" onclick="window.print()">🖨 Stampa / PDF</button>' +
      '<button class="btn btn-share" onclick="doShare()">📤 Condividi (WhatsApp/Email)</button>' +
    '</div>' +
    '<h1>'+title+'</h1>' +
    '<p style="font-size:11px;color:#888;margin-bottom:14px">Generato il '+now+'</p>' +
    bodyHtml +
    '</body></html>';
}

function exportReport(){
  var title = document.getElementById('rmodal-title').textContent;
  var bodyHtml = document.getElementById('rmodal-body').innerHTML;
  var standalone = buildStandaloneHtml(title, bodyHtml);

  // Prova Web Share API con file (Android Chrome, iOS Safari 15.1+)
  var blob = new Blob([standalone], {type:'text/html'});
  var file = new File([blob], 'report-bnb.html', {type:'text/html'});
  var now = new Date().toLocaleString('it-IT');

  if(navigator.share && navigator.canShare && navigator.canShare({files:[file]})){
    navigator.share({
      files: [file],
      title: title,
      text: 'Report B&B — ' + now
    }).catch(function(e){
      // Fallback: download diretto
      downloadHtml(standalone, title);
    });
  } else if(navigator.share){
    // Share senza file (link/testo)
    navigator.share({title: title, text: 'Report B&B — ' + now})
      .catch(function(){ downloadHtml(standalone, title); });
  } else {
    // Fallback PC: download diretto
    downloadHtml(standalone, title);
  }
}

function downloadHtml(htmlContent, title){
  var blob = new Blob([htmlContent], {type:'text/html'});
  var a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  var safe = title.replace(/[^a-zA-Z0-9À-ɏ ]/g,'-').trim();
  var d = new Date();
  a.download = 'report-' + safe + '-' + d.getDate() + '-' + (d.getMonth()+1) + '-' + d.getFullYear() + '.html';
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
  URL.revokeObjectURL(a.href);
}

function printReport(title, bodyHtml){
  var w = window.open('','_blank');
  w.document.write('<!DOCTYPE html><html><head><meta charset="UTF-8"><title>'+title+'</title>'+
    '<style>@media print{body{margin:0}}body{font-family:sans-serif;padding:20px;font-size:13px;max-width:800px;margin:0 auto}'+
    'h1{font-size:18px;margin-bottom:4px}h2{font-size:14px;font-weight:600;margin:16px 0 6px;color:#333}'+
    '.print-btn{position:fixed;top:10px;right:10px;padding:8px 16px;background:#1D9E75;color:#fff;border:none;border-radius:8px;cursor:pointer;font-size:14px}'+
    '.print-btn:hover{background:#0F6E56}@media print{.print-btn{display:none}}'+
    '.rec-card{border:1px solid #ddd;border-radius:8px;padding:12px;margin-bottom:12px}'+
    '.rec-header{display:flex;gap:8px;flex-wrap:wrap;margin-bottom:8px;align-items:center}'+
    '.rec-name{font-size:15px;font-weight:700;flex:1}'+
    '.rec-grid{display:grid;grid-template-columns:1fr 1fr;gap:6px;margin-bottom:8px}'+
    '.rec-lbl{font-size:10px;color:#888;text-transform:uppercase}'+
    '.rec-val{font-size:13px;font-weight:500}'+
    '.rsbadge{font-size:11px;padding:2px 8px;border-radius:10px;font-weight:500}'+
    '.b-bk{background:#E6F1FB;color:#185FA5}.b-si{background:#FAEEDA;color:#854F0B}'+
    '.b-ab{background:#f5e9fb;color:#7c3aed}.b-dir{background:#fef3c7;color:#92400e}'+
    '.b-ok{background:#E1F5EE;color:#0F6E56}.b-ko{background:#FAECE7;color:#993C1D}'+
    '.cam-day{margin-bottom:14px}.cam-day-hdr{font-weight:700;font-size:13px;border-bottom:2px solid #1D9E75;padding-bottom:4px;margin-bottom:8px;color:#0F6E56}'+
    '.cam-row{padding:7px 0;border-bottom:1px solid #eee}.cam-nome{font-weight:600}'+
    '.dotaz{font-size:12px;color:#444;margin-top:4px;padding:5px;background:#f5f5f2;border-radius:5px}'+
    '.dotaz-items{font-size:11px;color:#666}'+
    '.totali{background:#f0f9f5;border:1px solid #1D9E75;border-radius:8px;padding:12px;margin-top:16px}'+
    '.totali h3{margin:0 0 8px;font-size:14px;color:#0F6E56}.tot-row{display:flex;justify-content:space-between;font-size:13px;padding:2px 0}'+
    '.note-box{background:#fffbeb;border:1px solid #fde68a;border-radius:6px;padding:8px;margin-top:6px;font-size:12px}'+
    'table{width:100%;border-collapse:collapse;margin-bottom:12px}th{background:#f0f9f5;font-size:12px;padding:6px;text-align:left;border:1px solid #ddd}td{font-size:12px;padding:6px;border:1px solid #ddd}'+
    '</style></head><body>'+
    '<button class="print-btn" onclick="window.print()">🖨 Stampa / Salva PDF</button>'+
    '<h1>'+title+'</h1>'+
    bodyHtml+
    '</body></html>');
  w.document.close();
}

function reportColazione(today,todayStr,label){
  var notes = JSON.parse(localStorage.getItem('bnb_notes')||'{}');

  // Selettore data
  var startSel = document.getElementById('col-start-date');
  var refDay = today;
  if(startSel && startSel.value){
    var p = startSel.value.split('-');
    refDay = new Date(parseInt(p[0]), parseInt(p[1])-1, parseInt(p[2]));
  }
  var refStr = ds(refDay);
  var refLabel = refDay.toLocaleDateString('it-IT',{weekday:'long',day:'numeric',month:'long',year:'numeric'});

  // Colazione: dal giorno DOPO arrivo fino al giorno checkout incluso
  var presenti = bookings.filter(function(b){
    return b.checkin < refDay && b.checkout >= refDay &&
           (b.stato==='Attiva'||b.stato==='Modificata');
  });

  var dateSelector = '<div style="display:flex;align-items:center;gap:10px;margin-bottom:14px">' +
    '<label style="font-size:12px;color:var(--text2)">Data:</label>' +
    '<input type="date" id="col-start-date" value="'+ds(refDay)+'" ' +
    'style="padding:6px 10px;border-radius:8px;border:0.5px solid var(--border2);background:var(--bg2);color:var(--text);font-size:13px;outline:none">' +
  '</div>';

  if(!presenti.length) return '<div class="rdate">'+refLabel+'</div>' + dateSelector +
    '<div class="nodata">Nessun ospite a colazione in questa data</div>';

  var totAdulti=0, totBimbi=0;
  var rows = presenti.map(function(b){
    totAdulti += b.adulti; totBimbi += (b.bambini||0);
    var nd = notes[b.codice]||{};
    var noteCol = nd.noteColazione||'';
    var hasNote = noteCol.trim().length > 0;
    var ospiti = b.adulti+(b.bambini||0);
    return '<tr>' +
      '<td style="font-weight:600;padding:7px 8px;border:0.5px solid var(--border)">'+b.camera+'</td>' +
      '<td style="padding:7px 8px;border:0.5px solid var(--border)">'+b.cognome+' '+b.nome+'</td>' +
      '<td style="text-align:center;padding:7px 8px;border:0.5px solid var(--border)">'+b.adulti+(b.bambini?' +'+b.bambini:'')+'</td>' +
      '<td style="padding:7px 8px;border:0.5px solid var(--border)">' +
        '<textarea data-save-col-note="'+b.codice+'" placeholder="Note…" style="width:100%;min-width:120px;padding:4px 6px;border-radius:6px;border:0.5px solid var(--border2);background:var(--bg2);font-size:11px;resize:none;min-height:32px;font-family:inherit;'+(hasNote?'font-weight:700;color:#dc2626;':'color:var(--text)')+'">'+noteCol+'</textarea>' +
      '</td>' +
    '</tr>';
  }).join('');

  return '<div class="rdate">'+refLabel+'</div>' +
    dateSelector +
    '<table style="width:100%;border-collapse:collapse;margin-bottom:12px">' +
    '<thead><tr>' +
      '<th style="background:var(--bg3);padding:6px 8px;text-align:left;border:0.5px solid var(--border);font-size:11px">Camera</th>' +
      '<th style="background:var(--bg3);padding:6px 8px;text-align:left;border:0.5px solid var(--border);font-size:11px">Ospite</th>' +
      '<th style="background:var(--bg3);padding:6px 8px;text-align:center;border:0.5px solid var(--border);font-size:11px">Ad+Bim</th>' +
      '<th style="background:var(--bg3);padding:6px 8px;text-align:left;border:0.5px solid var(--border);font-size:11px">Note colazione</th>' +
    '</tr></thead><tbody>'+rows+'</tbody></table>' +
    '<div class="totali">' +
      '<h3>Riepilogo colazioni</h3>' +
      '<div class="tot-row"><span>Camere</span><span>'+presenti.length+'</span></div>' +
      '<div class="tot-row"><span>Adulti / Bambini</span><span>'+totAdulti+' / '+totBimbi+'</span></div>' +
      '<div class="tot-row" style="font-weight:700;border-top:1px solid rgba(0,0,0,.1);margin-top:6px;padding-top:6px"><span>Tavoli da allestire</span><span>'+presenti.length+' (max 2 pp.)</span></div>' +
    '</div>';
}

function reportPulizie(today,todayStr,label){
  var cos=bookings.filter(function(b){return ds(b.checkout)===todayStr&&(b.stato==='Attiva'||b.stato==='Modificata');});
  var cis=bookings.filter(function(b){return ds(b.checkin)===todayStr&&(b.stato==='Attiva'||b.stato==='Modificata');});
  var ins=bookings.filter(function(b){return b.checkin<today&&b.checkout>today&&ds(b.checkout)!==todayStr&&(b.stato==='Attiva'||b.stato==='Modificata');});
  function sec(title,items,tipo){
    if(!items.length)return'';
    var ico=tipo==='lib'?'🔴':tipo==='pre'?'🟢':'🔵';
    return'<div class="rsec"><div class="rstitle">'+ico+' '+title+'</div>'+items.map(function(b){
      return'<div class="cbrow"><div class="cb"></div><div><div class="cblabel">'+b.camera+'</div><div class="cbdetail">'+b.nome+' '+b.cognome+(tipo==='ins'?' · CO '+fmtDate(b.checkout):'')+'</div></div></div>';
    }).join('')+'</div>';
  }
  return'<div class="rdate">'+label+'</div>'+sec('Libera e pulizia a fondo',cos,'lib')+sec('Pulizia ordinaria (in-stay)',ins,'ins')+sec('Prepara per nuovi arrivi',cis,'pre')+(cos.length+ins.length+cis.length===0?'<div class="nodata">Nessuna pulizia oggi</div>':'')+'<div class="totbox"><div class="totrow"><span class="tlab">Pulizie a fondo</span><span>'+cos.length+'</span></div><div class="totrow"><span class="tlab">Ordinarie</span><span>'+ins.length+'</span></div><div class="totrow"><span class="tlab">Da preparare</span><span>'+cis.length+'</span></div><div class="totrow main"><span>Totale camere</span><span>'+(cos.length+ins.length+cis.length)+'</span></div></div>';
}
function reportReception(today,todayStr,label){
  var notes = JSON.parse(localStorage.getItem('bnb_notes')||'{}');

  // 1. Arrivi di oggi
  var cis = bookings.filter(function(b){
    return ds(b.checkin)===todayStr && (b.stato==='Attiva'||b.stato==='Modificata');
  });

  // 2. Partenze di oggi
  var cos = bookings.filter(function(b){
    return ds(b.checkout)===todayStr && (b.stato==='Attiva'||b.stato==='Modificata');
  });

  // 3. Ospiti in casa (checkin < oggi, checkout > oggi — già presenti, non partono oggi)
  var inCasa = bookings.filter(function(b){
    return b.checkin < today && b.checkout > today &&
           ds(b.checkin) !== todayStr &&
           (b.stato==='Attiva'||b.stato==='Modificata');
  });

  // Tutti per il totale (arrivi + in casa + partenze, senza duplicati)
  var tuttiMap = {};
  cis.forEach(function(b){ tuttiMap[b.codice]=b; });
  cos.forEach(function(b){ tuttiMap[b.codice]=b; });
  inCasa.forEach(function(b){ tuttiMap[b.codice]=b; });
  var tutti = Object.values(tuttiMap);

  // ── Card arrivo (check-in oggi) ───────────────────────────────────────
  function rowArrivo(b){
    var notti=getNotti(b), tassa=calcTassa(b);
    var nd=notes[b.codice]||{}, pagamentoOk=nd.pagamentoOk||false, noteText=nd.note||'';
    var hasBimbi=b.bambini>0;
    var quotaColor = pagamentoOk ? 'var(--accent-d)' : 'var(--coral-d)';
    var quotaLabel = pagamentoOk ? 'Saldato' : 'Da Incassare';
    return '<div class="rec-card" id="rec-'+b.codice+'">' +
      '<div class="rec-header">' +
        '<div class="rec-name">'+b.cognome+' '+b.nome+'</div>' +
        getCanaleLabel(b) +
        '<span class="rsbadge '+(pagamentoOk?'b-ok':'b-ko')+' pagamento-badge" style="cursor:pointer" data-toggle-pag="'+b.codice+'" id="badge-pag-'+b.codice+'">'+(pagamentoOk?'✓ Regolare':'⚠ Da regolarizzare')+'</span>' +
      '</div>' +
      '<div class="rec-grid">' +
        '<div class="rec-item"><span class="rec-lbl">Arrivo</span><span class="rec-val">'+fmtDate(b.checkin)+'</span></div>' +
        '<div class="rec-item"><span class="rec-lbl">Partenza</span><span class="rec-val">'+fmtDate(b.checkout)+'</span></div>' +
        '<div class="rec-item"><span class="rec-lbl">Notti</span><span class="rec-val">'+notti+'</span></div>' +
        '<div class="rec-item"><span class="rec-lbl">Camera</span><span class="rec-val">'+b.camera+'</span></div>' +
        '<div class="rec-item"><span class="rec-lbl">Adulti</span><span class="rec-val">'+b.adulti+'</span></div>' +
        '<div class="rec-item"><span class="rec-lbl">Bambini</span><span class="rec-val">'+(b.bambini||0)+'</span></div>' +
        '<div class="rec-item" style="grid-column:1/-1"><span class="rec-lbl">Quota soggiorno</span>' +
          '<div class="rec-val"><span style="color:'+quotaColor+';font-weight:600">€'+b.importo+' — '+quotaLabel+'</span></div></div>' +
        '<div class="rec-item" style="grid-column:1/-1"><span class="rec-lbl">Tassa soggiorno</span>' +
          '<div class="rec-val"><span style="color:'+quotaColor+';font-weight:600">€'+tassa+(hasBimbi?' <small style="color:var(--text2);font-weight:400">(verif. età bambini)</small>':'')+' — '+quotaLabel+'</span></div></div>' +
      '</div>' +
      '<textarea class="rec-note" data-save-note="'+b.codice+'" placeholder="Note ospite…" style="width:100%;padding:8px;border-radius:8px;border:0.5px solid var(--border2);background:var(--bg2);font-size:12px;resize:none;min-height:44px;font-family:inherit;'+(noteText.trim()?'font-weight:700;color:#dc2626':'color:var(--text)')+'">'+noteText+'</textarea>' +
    '</div>';
  }

  // ── Card ospite in casa (soggiorno in corso) ──────────────────────────
  function rowInCasa(b){
    var notti=getNotti(b), tassa=calcTassa(b);
    var nd=notes[b.codice]||{}, pagamentoOk=nd.pagamentoOk||false, noteText=nd.note||'';
    var giornoAttuale = Math.round((today - b.checkin)/86400000);
    var quotaColor = pagamentoOk ? 'var(--accent-d)' : 'var(--coral-d)';
    var quotaLabel = pagamentoOk ? 'Saldato' : 'Da Verificare';
    return '<div class="rec-card" id="rec-'+b.codice+'" style="border-left:3px solid var(--accent)">' +
      '<div class="rec-header">' +
        '<div class="rec-name">'+b.cognome+' '+b.nome+'</div>' +
        '<span class="rsbadge b-dir" style="font-size:10px">Notte '+giornoAttuale+'/'+notti+'</span>' +
        getCanaleLabel(b) +
        '<span class="rsbadge '+(pagamentoOk?'b-ok':'b-ko')+' pagamento-badge" style="cursor:pointer" data-toggle-pag="'+b.codice+'" id="badge-pag-'+b.codice+'">'+(pagamentoOk?'✓ Regolare':'⚠ Da regolarizzare')+'</span>' +
      '</div>' +
      '<div class="rec-grid">' +
        '<div class="rec-item"><span class="rec-lbl">Camera</span><span class="rec-val">'+b.camera+'</span></div>' +
        '<div class="rec-item"><span class="rec-lbl">Partenza</span><span class="rec-val">'+fmtDate(b.checkout)+'</span></div>' +
        '<div class="rec-item"><span class="rec-lbl">Adulti</span><span class="rec-val">'+b.adulti+'</span></div>' +
        '<div class="rec-item"><span class="rec-lbl">Notti tot.</span><span class="rec-val">'+notti+'</span></div>' +
        '<div class="rec-item" style="grid-column:1/-1"><span class="rec-lbl">Stato pagamento</span>' +
          '<div class="rec-val"><span style="color:'+quotaColor+';font-weight:600">€'+b.importo+' (+€'+tassa+' tassa) — '+quotaLabel+'</span></div></div>' +
      '</div>' +
      (noteText ? '<div style="font-size:12px;font-weight:700;color:#dc2626;padding:6px 8px;background:var(--bg3);border-radius:6px;margin-top:4px">📝 '+noteText+'</div>' : '') +
      '<textarea class="rec-note" data-save-note="'+b.codice+'" placeholder="Note ospite…" style="width:100%;padding:8px;border-radius:8px;border:0.5px solid var(--border2);background:var(--bg2);font-size:12px;resize:none;min-height:36px;font-family:inherit;margin-top:6px;'+(noteText.trim()?'font-weight:700;color:#dc2626':'color:var(--text)')+'">'+noteText+'</textarea>' +
    '</div>';
  }

  // ── Card partenza (check-out oggi) ────────────────────────────────────
  function rowPartenza(b){
    var notti=getNotti(b), tassa=calcTassa(b);
    var nd=notes[b.codice]||{}, pagamentoOk=nd.pagamentoOk||false;
    var quotaColor = pagamentoOk ? 'var(--accent-d)' : 'var(--coral-d)';
    var quotaLabel = pagamentoOk ? 'Saldato' : 'Da Saldare';
    return '<div class="rec-card checkout-card" id="rec-'+b.codice+'">' +
      '<div class="rec-header">' +
        '<div class="rec-name">'+b.cognome+' '+b.nome+'</div>' +
        '<span class="rsbadge b-out">Check-out</span>' +
        getCanaleLabel(b) +
        '<span class="rsbadge '+(pagamentoOk?'b-ok':'b-ko')+' pagamento-badge" style="cursor:pointer" data-toggle-pag="'+b.codice+'" id="badge-pag-'+b.codice+'">'+(pagamentoOk?'✓ Regolare':'⚠ Da regolarizzare')+'</span>' +
      '</div>' +
      '<div class="rec-grid">' +
        '<div class="rec-item"><span class="rec-lbl">Camera</span><span class="rec-val">'+b.camera+'</span></div>' +
        '<div class="rec-item"><span class="rec-lbl">Notti</span><span class="rec-val">'+notti+'</span></div>' +
        '<div class="rec-item"><span class="rec-lbl">Adulti</span><span class="rec-val">'+b.adulti+'</span></div>' +
        '<div class="rec-item"><span class="rec-lbl">Canale</span><span class="rec-val">'+getCanale(b)+'</span></div>' +
        '<div class="rec-item" style="grid-column:1/-1"><span class="rec-lbl">Quota soggiorno</span>' +
          '<div class="rec-val"><span style="color:'+quotaColor+';font-weight:600">€'+b.importo+' — '+quotaLabel+'</span></div></div>' +
        '<div class="rec-item" style="grid-column:1/-1"><span class="rec-lbl">Tassa soggiorno</span>' +
          '<div class="rec-val"><span style="color:'+quotaColor+';font-weight:600">€'+tassa+' — '+quotaLabel+'</span></div></div>' +
      '</div></div>';
  }

  var html = '<div class="rdate">'+label+'</div>';

  // Sezione arrivi
  if(cis.length){
    html += '<div class="rstitle" style="margin-bottom:10px">🟢 Arrivi oggi ('+cis.length+')</div>';
    html += cis.map(rowArrivo).join('');
  } else {
    html += '<div class="nodata" style="margin-bottom:12px">Nessun arrivo oggi</div>';
  }

  // Sezione ospiti in casa
  if(inCasa.length){
    html += '<div class="rstitle" style="margin:14px 0 10px">🏠 Ospiti in casa ('+inCasa.length+')</div>';
    // Ordina per data checkout (chi parte prima in cima)
    inCasa.sort(function(a,b){ return a.checkout - b.checkout; });
    html += inCasa.map(rowInCasa).join('');
  }

  // Sezione partenze
  if(cos.length){
    html += '<div class="rstitle" style="margin:14px 0 10px">🔴 Partenze oggi ('+cos.length+')</div>';
    html += cos.map(rowPartenza).join('');
  } else {
    html += '<div class="nodata" style="margin-bottom:12px">Nessuna partenza oggi</div>';
  }

  // Totali inline (senza dipendenza esterna)
  var da_inc = tutti.filter(function(b){ return !(notes[b.codice]||{}).pagamentoOk; });
  var saldati_n = tutti.filter(function(b){ return (notes[b.codice]||{}).pagamentoOk; });
  var totQ = da_inc.reduce(function(s,b){return s+b.importo;},0);
  var totT = da_inc.reduce(function(s,b){return s+calcTassa(b);},0);
  html += '<div class="totali" id="totali-reception" style="margin-top:14px">' +
    '<h3>Totali reception oggi</h3>' +
    (saldati_n.length ? '<div class="tot-row" style="color:var(--accent-d)"><span>✓ Ospiti saldati</span><span>'+saldati_n.length+'</span></div>' : '') +
    '<div class="tot-row"><span>Ospiti da regolarizzare</span><span>'+da_inc.length+'</span></div>' +
    '<div class="tot-row"><span>Quota soggiorni</span><span>€'+totQ.toFixed(2)+'</span></div>' +
    '<div class="tot-row"><span>Tasse soggiorno</span><span>€'+totT.toFixed(2)+'</span></div>' +
    '<div class="tot-row" style="font-weight:700;border-top:1px solid rgba(0,0,0,.1);margin-top:6px;padding-top:6px"><span>Totale da incassare</span><span style="color:var(--coral-d)">€'+(totQ+totT).toFixed(2)+'</span></div>' +
  '</div>';
  return html;
}


function reportCamere(today,todayStr,label){
  var notes = JSON.parse(localStorage.getItem('bnb_notes')||'{}');

  var startSel = document.getElementById('cam-start-date');
  var startDay = today;
  if(startSel && startSel.value){
    var parts = startSel.value.split('-');
    startDay = new Date(parseInt(parts[0]),parseInt(parts[1])-1,parseInt(parts[2]));
  }

  var days = [];
  for(var i=0;i<7;i++){
    var d = new Date(startDay);
    d.setDate(d.getDate()+i);
    days.push(d);
  }

  var totFermata=0, totPartenza=0, totArrivo=0;
  var incassoLordo=0, incassoComm=0, incassoTasse=0;
  var prenotazioniSettimana = new Set();

  var html = '<div class="rdate">'+label+'</div>';

  // Selettore data
  html += '<div style="display:flex;align-items:center;gap:10px;margin-bottom:14px">' +
    '<label style="font-size:12px;color:var(--text2)">Data inizio:</label>' +
    '<input type="date" id="cam-start-date" value="'+ds(startDay)+'" ' +
    'style="padding:6px 10px;border-radius:8px;border:0.5px solid var(--border2);background:var(--bg2);color:var(--text);font-size:13px;outline:none">' +
  '</div>';

  // ── SEZIONE CHECK-IN ──────────────────────────────────────────────────
  var tuttiArrivi = [];
  days.forEach(function(day){
    var dayStr = ds(day);
    var arrivals = bookings.filter(function(b){
      return (b.stato==='Attiva'||b.stato==='Modificata') && ds(b.checkin)===dayStr;
    });
    arrivals.forEach(function(b){ tuttiArrivi.push({b:b, day:day}); totArrivo++; });
  });

  html += '<div class="rstitle" style="margin-bottom:10px;margin-top:4px">🟢 Check-in settimana ('+tuttiArrivi.length+')</div>';
  if(tuttiArrivi.length){
    html += '<table style="width:100%;border-collapse:collapse;margin-bottom:16px">' +
      '<thead><tr>' +
      '<th style="background:var(--bg3);padding:6px 8px;text-align:left;border:0.5px solid var(--border);font-size:11px">Data</th>' +
      '<th style="background:var(--bg3);padding:6px 8px;text-align:left;border:0.5px solid var(--border);font-size:11px">Camera</th>' +
      '<th style="background:var(--bg3);padding:6px 8px;text-align:left;border:0.5px solid var(--border);font-size:11px">Ospite</th>' +
      '<th style="background:var(--bg3);padding:6px 8px;text-align:center;border:0.5px solid var(--border);font-size:11px">Ospiti</th>' +
      '<th style="background:var(--bg3);padding:6px 8px;text-align:left;border:0.5px solid var(--border);font-size:11px">CO</th>' +
      '<th style="background:var(--bg3);padding:6px 8px;text-align:left;border:0.5px solid var(--border);font-size:11px">Canale</th>' +
      '</tr></thead><tbody>';
    tuttiArrivi.forEach(function(item){
      var b=item.b, nd=notes[b.codice]||{};
      var pagOk = nd.pagamentoOk;
      html += '<tr style="background:rgba(29,158,117,.04)">' +
        '<td style="padding:6px 8px;border:0.5px solid var(--border);font-size:12px;font-weight:600">'+
          item.day.toLocaleDateString('it-IT',{weekday:'short',day:'numeric',month:'short'})+'</td>' +
        '<td style="padding:6px 8px;border:0.5px solid var(--border);font-size:12px;font-weight:600">'+b.camera+'</td>' +
        '<td style="padding:6px 8px;border:0.5px solid var(--border);font-size:12px">'+b.cognome+' '+b.nome+'</td>' +
        '<td style="padding:6px 8px;border:0.5px solid var(--border);font-size:12px;text-align:center">'+(b.adulti+(b.bambini||0))+'</td>' +
        '<td style="padding:6px 8px;border:0.5px solid var(--border);font-size:12px">'+fmtDate(b.checkout)+'</td>' +
        '<td style="padding:6px 8px;border:0.5px solid var(--border);font-size:11px">'+getCanale(b)+'</td>' +
      '</tr>';
    });
    html += '</tbody></table>';
  } else {
    html += '<div class="nodata" style="margin-bottom:16px">Nessun check-in questa settimana</div>';
  }

  // ── SEZIONE PULIZIE GIORNALIERE (solo fermata e partenza) ────────────
  html += '<div class="rstitle" style="margin-bottom:10px">🧹 Pulizie giornaliere</div>';

  days.forEach(function(day){
    var dayStr = ds(day);
    var dayLabel = day.toLocaleDateString('it-IT',{weekday:'long',day:'numeric',month:'long'});

    // Solo fermata e partenza (NO arrivi)
    var fermata = bookings.filter(function(b){
      return (b.stato==='Attiva'||b.stato==='Modificata') &&
             b.checkin < day && b.checkout > day;
    });
    var partenza = bookings.filter(function(b){
      return (b.stato==='Attiva'||b.stato==='Modificata') &&
             ds(b.checkout) === dayStr;
    });

    // Aggiorna contatori e incassi (una volta per prenotazione)
    partenza.forEach(function(b){ totPartenza++; });
    fermata.forEach(function(b){ totFermata++; });

    // Incassi settimana
    var tutteDelGiorno = fermata.concat(partenza);
    tutteDelGiorno.forEach(function(b){
      if(!prenotazioniSettimana.has(b.codice)){
        prenotazioniSettimana.add(b.codice);
        incassoLordo  += b.importo;
        incassoComm   += b.commissioni;
        incassoTasse  += calcTassa(b);
      }
    });

    if(!fermata.length && !partenza.length) return;

    html += '<div class="cam-day"><div class="cam-day-hdr">'+dayLabel+'</div>';

    // Prima le partenze, poi le fermate
    partenza.forEach(function(b){
      html += camRow(b, day, 'partenza', notes);
    });
    fermata.forEach(function(b){
      html += camRow(b, day, 'fermata', notes);
    });

    html += '</div>';
  });

  // Totali
  var incassoNetto = incassoLordo - incassoComm - incassoTasse;
  html += '<div class="totali" style="margin-top:6px">' +
    '<h3>Totali settimana</h3>' +
    '<div class="tot-row"><span>🟢 Check-in</span><span>'+totArrivo+'</span></div>' +
    '<div class="tot-row"><span>🔴 Partenze (pulizia a fondo)</span><span>'+totPartenza+'</span></div>' +
    '<div class="tot-row"><span>🔵 Fermate (riassetto)</span><span>'+totFermata+'</span></div>' +
    '<div class="tot-row" style="margin-top:8px"><span>Incasso lordo</span><span>€'+incassoLordo.toFixed(2)+'</span></div>' +
    '<div class="tot-row"><span>Commissioni</span><span>− €'+incassoComm.toFixed(2)+'</span></div>' +
    '<div class="tot-row"><span>Tasse soggiorno</span><span>€'+incassoTasse.toFixed(2)+'</span></div>' +
    '<div class="tot-row" style="font-weight:700;border-top:1px solid rgba(0,0,0,.1);margin-top:6px;padding-top:6px"><span>Incasso netto</span><span>€'+incassoNetto.toFixed(2)+'</span></div>' +
  '</div>';

  return html;
}

function camRow(b, day, tipo, notes){
  var isCheckout = tipo === 'partenza';
  var checkinDay = Math.round((day - b.checkin)/86400000);
  var notti = getNotti(b);
  var nd = notes[b.codice]||{};
  var pagamentoOk = nd.pagamentoOk||false;
  var noteText = nd.note||'';
  var ospiti = b.adulti + (b.bambini||0);
  var tipoPulizia = isCheckout ? '🔴 In partenza — Pulizia a fondo' : '🔵 In fermata — Riassetto';

  return '<div class="cam-row">' +
    '<div class="cam-row-hdr">' +
      '<span class="cam-nome">'+b.camera+'</span>' +
      '<span class="rsbadge b-dir" style="font-size:10px">'+b.cognome+' '+b.nome+'</span>' +
      '<span class="rsbadge '+(pagamentoOk?'b-ok':'b-ko')+'">'+(pagamentoOk?'✓':'!')+'</span>' +
    '</div>' +
    '<div style="font-size:12px;color:var(--text2);margin:2px 0">Notte '+(isCheckout?notti:checkinDay+1)+'/'+notti+' · '+getCanale(b)+'</div>' +
    '<div style="font-size:12px;margin:3px 0"><strong>'+tipoPulizia+'</strong></div>' +
    '<div style="font-size:11px;color:var(--text2)">👥 '+ospiti+' ('+b.adulti+' ad.'+(b.bambini?' + '+b.bambini+' bim.':'')+')</div>' +
    dotazioniCamera(b, day) +
    (noteText?'<div style="font-size:12px;font-weight:700;color:#dc2626;margin-top:4px;padding:5px;background:var(--bg3);border-radius:6px">📝 '+noteText+'</div>':'') +
  '</div>';
}

function refreshReportColazione(){
  var today = new Date(); today.setHours(0,0,0,0);
  var body = document.getElementById('rmodal-body');
  if(body) body.innerHTML = reportColazione(today, ds(today),
    today.toLocaleDateString('it-IT',{weekday:'long',day:'numeric',month:'long',year:'numeric'}));
}

function refreshReportCamere(){
  var today = new Date(); today.setHours(0,0,0,0);
  var todayStr = ds(today);
  var label = today.toLocaleDateString('it-IT',{weekday:'long',day:'numeric',month:'long',year:'numeric'});
  var body = document.getElementById('rmodal-body');
  if(body) body.innerHTML = reportCamere(today,todayStr,label);
}

function reportSettimanale(today,todayStr,label){
  var notes = JSON.parse(localStorage.getItem('bnb_notes')||'{}');
  var we = new Date(today); we.setDate(we.getDate()+7);
  var active = bookings.filter(function(b){return b.stato==='Attiva'||b.stato==='Modificata';});

  // Raggruppa per data check-in nella settimana
  var giorni = {};
  var totFermata=0, totPartenza=0;
  var rev=0, comm=0, tasse=0;

  // Raccogli tutte le date uniche nei prossimi 7 giorni
  var days = [];
  for(var i=0;i<7;i++){
    var d = new Date(today); d.setDate(d.getDate()+i);
    days.push(d);
  }

  var html = '<div class="rdate">'+label+'</div>';

  // Tabella calendario
  days.forEach(function(day){
    var dayStr = ds(day);
    var dayLabel = day.toLocaleDateString('it-IT',{weekday:'long',day:'numeric',month:'long'});

    // Prenotazioni che toccano questo giorno
    var arrivals = active.filter(function(b){ return ds(b.checkin)===dayStr; });
    var departures = active.filter(function(b){ return ds(b.checkout)===dayStr; });
    var staying = active.filter(function(b){ return b.checkin<day && b.checkout>day && ds(b.checkin)!==dayStr; });

    if(!arrivals.length && !departures.length && !staying.length) return;

    html += '<div class="cam-day">';
    html += '<div class="cam-day-hdr">'+dayLabel+'</div>';

    html += '<table style="width:100%;border-collapse:collapse;font-size:12px;margin-bottom:4px">';
    html += '<thead><tr><th style="background:var(--bg3);padding:5px 8px;text-align:left;border:0.5px solid var(--border)">Camera</th>' +
            '<th style="background:var(--bg3);padding:5px 8px;text-align:left;border:0.5px solid var(--border)">Ospite</th>' +
            '<th style="background:var(--bg3);padding:5px 8px;text-align:center;border:0.5px solid var(--border)">Ospiti</th>' +
            '<th style="background:var(--bg3);padding:5px 8px;text-align:left;border:0.5px solid var(--border)">Stato</th>' +
            '<th style="background:var(--bg3);padding:5px 8px;text-align:right;border:0.5px solid var(--border)">€</th></tr></thead><tbody>';

    function rowCal(b, tipo){
      var ospiti = b.adulti + (b.bambini||0);
      var nd = notes[b.codice]||{};
      var pag = nd.pagamentoOk ? '✓' : '!';
      var pagColor = nd.pagamentoOk ? '#0F6E56' : '#993C1D';
      var tipoLabel = tipo==='arrivo' ? '🟢 Arr.' : tipo==='partenza' ? '🔴 Part.' : '🔵 Stay';
      var bgRow = tipo==='arrivo' ? 'rgba(29,158,117,.06)' : tipo==='partenza' ? 'rgba(153,60,29,.06)' : '';
      if(tipo==='arrivo'){ totFermata++; rev+=b.importo; comm+=b.commissioni; tasse+=calcTassa(b); }
      if(tipo==='partenza') totPartenza++;
      return '<tr style="background:'+bgRow+'">' +
        '<td style="padding:5px 8px;border:0.5px solid var(--border);font-weight:600">'+b.camera+'</td>' +
        '<td style="padding:5px 8px;border:0.5px solid var(--border)">'+b.cognome+' '+b.nome+'</td>' +
        '<td style="padding:5px 8px;border:0.5px solid var(--border);text-align:center">'+ospiti+'</td>' +
        '<td style="padding:5px 8px;border:0.5px solid var(--border)">'+tipoLabel+' <span style="color:'+pagColor+'">'+pag+'</span></td>' +
        '<td style="padding:5px 8px;border:0.5px solid var(--border);text-align:right">'+b.importo+'</td></tr>';
    }

    departures.forEach(function(b){ html += rowCal(b,'partenza'); });
    arrivals.forEach(function(b){ html += rowCal(b,'arrivo'); });
    staying.forEach(function(b){ html += rowCal(b,'stay'); });

    html += '</tbody></table></div>';
  });

  var netto = rev - comm - tasse;
  html += '<div class="totali">' +
    '<h3>Totali settimana</h3>' +
    '<div class="tot-row"><span>🟢 Arrivi</span><span>'+totFermata+'</span></div>' +
    '<div class="tot-row"><span>🔴 Partenze</span><span>'+totPartenza+'</span></div>' +
    '<div class="tot-row" style="margin-top:8px"><span>Incasso lordo</span><span>€'+rev.toFixed(2)+'</span></div>' +
    '<div class="tot-row"><span>Commissioni</span><span>− €'+comm.toFixed(2)+'</span></div>' +
    '<div class="tot-row"><span>Tasse soggiorno</span><span>€'+tasse.toFixed(2)+'</span></div>' +
    '<div class="tot-row" style="font-weight:700;border-top:1px solid rgba(0,0,0,.1);margin-top:6px;padding-top:6px"><span>Incasso netto</span><span>€'+netto.toFixed(2)+'</span></div>' +
  '</div>';

  return html;
}

function reportFabiola(today,todayStr,label){
  var notes = JSON.parse(localStorage.getItem('bnb_notes')||'{}');
  var html = '';

  // Sezione 1: Reception
  html += '<h2 style="margin:0 0 10px;font-size:16px;color:var(--accent-d,#0F6E56);border-bottom:2px solid var(--accent-d,#0F6E56);padding-bottom:6px">📋 RECEPTION</h2>';
  html += reportReception(today,todayStr,label);

  // Sezione 2: Colazioni
  html += '<h2 style="margin:20px 0 10px;font-size:16px;color:var(--accent-d,#0F6E56);border-bottom:2px solid var(--accent-d,#0F6E56);padding-bottom:6px">☕ COLAZIONI</h2>';
  html += reportColazione(today,todayStr,label);

  // Sezione 3: Camere (7 giorni da oggi)
  html += '<h2 style="margin:20px 0 10px;font-size:16px;color:var(--accent-d,#0F6E56);border-bottom:2px solid var(--accent-d,#0F6E56);padding-bottom:6px">🛏 CAMERE — prossimi 7 giorni</h2>';
  html += reportCamere(today,todayStr,label);

  return html;
}

