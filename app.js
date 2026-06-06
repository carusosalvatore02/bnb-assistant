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
  });
  document.getElementById('input-row').style.display = 'none';
  var cancelBtn = document.getElementById('file-panel-cancel');
  if (cancelBtn) cancelBtn.addEventListener('click', hideFilePanel);
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
  saveToStorage(bookings.map(function(b){
    return Object.assign({},b,{checkin:b.checkin?ds(b.checkin):null,checkout:b.checkout?ds(b.checkout):null});
  }), now);
  var d = new Date(now);
  setStatus('ok','Aggiornato alle ' + d.toLocaleTimeString('it-IT',{hour:'2-digit',minute:'2-digit'}));
  document.getElementById('app-sub').textContent = bookings.length + ' prenotazioni';
  showMainApp();
  renderAll();
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
  localStorage.setItem('bnb_notes', JSON.stringify(notes));
  // Aggiorna tutti i badge con questo codice
  document.querySelectorAll('#badge-pag-'+codice).forEach(function(el){
    el.textContent = notes[codice].pagamentoOk ? '✓ Regolare' : '⚠ Da regolarizzare';
    el.className = 'rsbadge ' + (notes[codice].pagamentoOk ? 'b-ok' : 'b-ko');
  });
}

function saveNote(codice, valore){
  var notes = JSON.parse(localStorage.getItem('bnb_notes')||'{}');
  if(!notes[codice]) notes[codice] = {};
  // Accetta valore diretto o lo legge dal DOM
  var val = valore !== undefined ? valore : (document.getElementById('note-'+codice)||{}).value || '';
  notes[codice].note = val;
  localStorage.setItem('bnb_notes', JSON.stringify(notes));
}

function saveColNote(codice, valore){
  var notes = JSON.parse(localStorage.getItem('bnb_notes')||'{}');
  if(!notes[codice]) notes[codice] = {};
  notes[codice].noteColazione = valore !== undefined ? valore : '';
  localStorage.setItem('bnb_notes', JSON.stringify(notes));
}

function resetApiKey(){localStorage.removeItem('bnb_apikey');addMsg('ai','API key rimossa. Alla prossima domanda ti verrà chiesta di nuovo.');}

// ─── REPORTS ─────────────────────────────────────────────
function openReport(type){
  if(!bookings.length){alert('Sincronizza prima i dati.');return;}
  var modal=document.getElementById('rmodal'),body=document.getElementById('rmodal-body'),title=document.getElementById('rmodal-title');
  modal.classList.add('open');
  var today=new Date();today.setHours(0,0,0,0);
  var todayStr=ds(today),label=today.toLocaleDateString('it-IT',{weekday:'long',day:'numeric',month:'long',year:'numeric'});
  var map={colazione:['Colazioni',reportColazione],pulizie:['Pulizie',reportPulizie],reception:['Reception',reportReception],camere:['Report Camere',reportCamere],settimanale:['Riepilogo settimanale',reportSettimanale],fabiola:['Report Fabiola',reportFabiola]};
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
  // refDay: Date del giorno di riferimento
  var notti = getNotti(b);
  var checkinDay = Math.round((refDay - b.checkin)/86400000); // 0=giorno arrivo
  var isCheckout = ds(b.checkout) === ds(refDay);
  var isCheckin  = ds(b.checkin)  === ds(refDay);

  if(isCheckout || isCheckin){
    return '<div class="dotaz">🔄 <strong>Cambio totale 100%</strong><br>' +
           '<span class="dotaz-items">2 lenzuola · 2 federe · 2 teli doccia · 2 teli viso · 2 teli bidet · 1 tappetino · 2 spazzolini · 2 cuffie doccia</span></div>';
  }

  // Giorno di fermata: checkinDay = giorni trascorsi dall'arrivo (1 = prima notte finita)
  var giornoFermata = checkinDay; // 1-based dal giorno dopo l'arrivo
  var items = [];

  // Set letto: ogni 4 giorni (giorni 4, 8, 12...)
  var cambioLetto = giornoFermata > 0 && giornoFermata % 4 === 0;
  // Set bagno: ogni 2 giorni (giorni 2, 4, 6...)
  var cambioBagno = giornoFermata > 0 && giornoFermata % 2 === 0;

  if(cambioLetto){
    items.push('🛏 Set Letto: 2 lenzuola · 2 federe');
  }
  if(cambioBagno){
    items.push('🚿 Set Bagno: 2 teli doccia · 2 teli viso · 2 teli bidet · 1 tappetino');
  }
  if(!cambioLetto && !cambioBagno){
    return '<div class="dotaz">✨ Solo riassetto — nessun cambio biancheria oggi</div>';
  }
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
  var presenti = bookings.filter(function(b){
    return b.checkin<=today && b.checkout>today &&
           (b.stato==='Attiva'||b.stato==='Modificata');
  });

  if(!presenti.length) return '<div class="rdate">'+label+'</div><div class="nodata">Nessun ospite presente oggi</div>';

  var totAdulti = presenti.reduce(function(s,b){return s+b.adulti;},0);
  var totBimbi  = presenti.reduce(function(s,b){return s+(b.bambini||0);},0);
  var totTavoli = presenti.length; // 1 tavolo per camera

  var rows = presenti.map(function(b){
    var nd = notes[b.codice]||{};
    var noteCol = nd.noteColazione||'';
    var ospiti = b.adulti + (b.bambini||0);
    return '<div class="rec-card">' +
      '<div class="rec-header"><div class="rec-name">'+b.camera+'</div>' +
      '<span class="rsbadge b-dir">'+b.cognome+' '+b.nome+'</span></div>' +
      '<div class="rec-grid">' +
        '<div class="rec-item"><span class="rec-lbl">Adulti</span><span class="rec-val">'+b.adulti+'</span></div>' +
        '<div class="rec-item"><span class="rec-lbl">Bambini</span><span class="rec-val">'+(b.bambini||0)+'</span></div>' +
        '<div class="rec-item"><span class="rec-lbl">Coperti</span><span class="rec-val">'+ospiti+'</span></div>' +
        '<div class="rec-item"><span class="rec-lbl">CO previsto</span><span class="rec-val">'+fmtDate(b.checkout)+'</span></div>' +
      '</div>' +
      '<textarea class="rec-note" data-save-col-note="'+b.codice+'" placeholder="Note colazione (allergie, dieta, richieste…)" style="width:100%;padding:8px;border-radius:8px;border:0.5px solid var(--border2);background:var(--bg2);color:var(--text);font-size:12px;resize:none;min-height:44px;font-family:inherit">'+noteCol+'</textarea>' +
    '</div>';
  }).join('');

  var totali = '<div class="totali" style="margin-top:14px">' +
    '<h3>Totali colazioni</h3>' +
    '<div class="tot-row"><span>Camere presenti</span><span>'+presenti.length+'</span></div>' +
    '<div class="tot-row"><span>Totale adulti</span><span>'+totAdulti+'</span></div>' +
    (totBimbi?'<div class="tot-row"><span>Totale bambini</span><span>'+totBimbi+'</span></div>':'') +
    '<div class="tot-row" style="font-weight:700;border-top:1px solid rgba(0,0,0,.1);margin-top:6px;padding-top:6px"><span>Tavoli da allestire</span><span>'+totTavoli+' (max 2 persone per tavolo)</span></div>' +
  '</div>';

  return '<div class="rdate">'+label+'</div>' + rows + totali;
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
  var cis = bookings.filter(function(b){return ds(b.checkin)===todayStr&&(b.stato==='Attiva'||b.stato==='Modificata');});
  var cos = bookings.filter(function(b){return ds(b.checkout)===todayStr&&(b.stato==='Attiva'||b.stato==='Modificata');});

  function rowArrivo(b){
    var notti = getNotti(b);
    var tassa = calcTassa(b);
    var nd = notes[b.codice]||{};
    var pagamentoOk = nd.pagamentoOk||false;
    var noteText = nd.note||'';
    var hasBimbi = b.bambini > 0;
    var canale = getCanale(b);

    // Tutti i canali: incasso fisico sul posto
    // Booking.com: incassa sempre tu, tassa sempre visibile
    var quotaStr = '<span style="color:var(--coral-d);font-weight:600">€'+b.importo+' — Da incassare sul posto</span>';

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
        '<div class="rec-item" style="grid-column:1/-1"><span class="rec-lbl">Quota soggiorno</span><div class="rec-val">'+quotaStr+'</div></div>' +
        '<div class="rec-item" style="grid-column:1/-1"><span class="rec-lbl">Tassa soggiorno</span><div class="rec-val" style="color:var(--coral-d);font-weight:600">€'+tassa+(hasBimbi?' <small style="color:var(--text2);font-weight:400">(verificare età bambini per tassa)</small>':'')+'</div></div>' +
      '</div>' +
      '<textarea class="rec-note" data-save-note="'+b.codice+'" placeholder="Note ospite…" style="width:100%;padding:8px;border-radius:8px;border:0.5px solid var(--border2);background:var(--bg2);color:var(--text);font-size:12px;resize:none;min-height:44px;font-family:inherit">'+noteText+'</textarea>' +
    '</div>';
  }

  function rowPartenza(b){
    var nd = notes[b.codice]||{};
    var pagamentoOk = nd.pagamentoOk||false;
    var notti = getNotti(b);
    return '<div class="rec-card checkout-card">' +
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
      '</div></div>';
  }

  var html = '<div class="rdate">'+label+'</div>';
  html += cis.length ? '<div class="rstitle" style="margin-bottom:10px">🟢 Arrivi ('+cis.length+')</div>' + cis.map(rowArrivo).join('') : '<div class="nodata" style="margin-bottom:12px">Nessun arrivo oggi</div>';
  html += cos.length ? '<div class="rstitle" style="margin:14px 0 10px">🔴 Partenze ('+cos.length+')</div>' + cos.map(rowPartenza).join('') : '<div class="nodata">Nessuna partenza oggi</div>';

  // Totale da incassare oggi
  var totIncasso = cis.reduce(function(s,b){return s+b.importo;},0);
  var totTasse   = cis.reduce(function(s,b){return s+calcTassa(b);},0);
  html += '<div class="totali" style="margin-top:14px">' +
    '<h3>Totali da incassare oggi</h3>' +
    '<div class="tot-row"><span>Quota soggiorni</span><span>€'+totIncasso.toFixed(2)+'</span></div>' +
    '<div class="tot-row"><span>Tasse soggiorno</span><span>€'+totTasse.toFixed(2)+'</span></div>' +
    '<div class="tot-row" style="font-weight:700;border-top:1px solid rgba(0,0,0,.1);margin-top:6px;padding-top:6px"><span>Totale da incassare</span><span>€'+(totIncasso+totTasse).toFixed(2)+'</span></div>' +
  '</div>';

  return html;
}

function reportCamere(today,todayStr,label){
  var notes = JSON.parse(localStorage.getItem('bnb_notes')||'{}');

  // Leggi data di partenza dal selettore (se presente)
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

  var totFermata = 0, totPartenza = 0, totArrivo = 0;
  var incassoLordo=0, incassoComm=0, incassoTasse=0;
  var prenotazioniSettimana = new Set();

  var html = '<div class="rdate">'+label+'</div>';

  // Selettore data
  html += '<div style="display:flex;align-items:center;gap:10px;margin-bottom:14px">' +
    '<label style="font-size:12px;color:var(--text2)">Data inizio:</label>' +
    '<input type="date" id="cam-start-date" value="'+ds(startDay)+'" ' +
    'style="padding:6px 10px;border-radius:8px;border:0.5px solid var(--border2);background:var(--bg2);color:var(--text);font-size:13px" ' +
    'onchange="refreshReportCamere()">' +
  '</div>';

  days.forEach(function(day){
    var dayStr = ds(day);
    var dayLabel = day.toLocaleDateString('it-IT',{weekday:'long',day:'numeric',month:'long'});

    var active = bookings.filter(function(b){
      return (b.stato==='Attiva'||b.stato==='Modificata') &&
             b.checkin <= day && b.checkout > day;
    });

    if(!active.length){
      html += '<div class="cam-day"><div class="cam-day-hdr">'+dayLabel+'</div>' +
              '<div class="nodata" style="padding:6px 0;font-size:12px">Nessuna camera occupata</div></div>';
      return;
    }

    html += '<div class="cam-day"><div class="cam-day-hdr">'+dayLabel+'</div>';

    active.forEach(function(b){
      var isCheckout = ds(b.checkout) === dayStr;
      var isCheckin  = ds(b.checkin)  === dayStr;
      var checkinDay = Math.round((day - b.checkin)/86400000);
      var notti = getNotti(b);
      var nd = notes[b.codice]||{};
      var pagamentoOk = nd.pagamentoOk||false;
      var noteText = nd.note||'';
      var ospiti = b.adulti + (b.bambini||0);

      var tipoPulizia, tipoCls;
      if(isCheckout){ tipoPulizia='🔴 In partenza — Pulizia a fondo'; tipoCls='coral'; totPartenza++; }
      else if(isCheckin){ tipoPulizia='🟢 Arrivo — Rifare tutta la camera 100%'; tipoCls='green'; totArrivo++; }
      else { tipoPulizia='🔵 In fermata — Riassetto'; tipoCls='blue'; totFermata++; }

      // Incassi settimana (conta ogni prenotazione una volta sola)
      if(!prenotazioniSettimana.has(b.codice) && (b.checkin >= days[0] && b.checkin < days[days.length-1] || b.checkout > days[0] && b.checkout <= days[days.length-1])){
        prenotazioniSettimana.add(b.codice);
        incassoLordo  += b.importo;
        incassoComm   += b.commissioni;
        incassoTasse  += calcTassa(b);
      }

      html += '<div class="cam-row">' +
        '<div class="cam-row-hdr">' +
          '<span class="cam-nome">'+b.camera+'</span>' +
          '<span class="rsbadge b-dir" style="font-size:10px">'+b.cognome+' '+b.nome+'</span>' +
          '<span class="rsbadge '+(pagamentoOk?'b-ok':'b-ko')+'">'+(pagamentoOk?'✓':'⚠')+'</span>' +
        '</div>' +
        '<div style="font-size:12px;color:var(--text2);margin:2px 0">Notte '+(isCheckout?notti:checkinDay+1)+'/'+notti+' · '+getCanale(b)+'</div>' +
        '<div style="font-size:12px;margin:3px 0"><strong>'+tipoPulizia+'</strong></div>' +
        '<div style="font-size:11px;color:var(--text2)">👥 '+ospiti+' ospiti ('+b.adulti+' adulti'+(b.bambini?' + '+b.bambini+' bimbi':'')+')</div>' +
        dotazioniCamera(b, day) +
        (noteText?'<div style="font-size:11px;color:var(--text2);margin-top:4px;padding:5px;background:var(--bg3);border-radius:6px">📝 '+noteText+'</div>':'') +
      '</div>';
    });

    html += '</div>';
  });

  // Totali
  var incassoNetto = incassoLordo - incassoComm - incassoTasse;
  html += '<div class="totali" style="margin-top:6px">' +
    '<h3>Totali settimana</h3>' +
    '<div class="tot-row"><span>🔴 Camere in partenza</span><span>'+totPartenza+'</span></div>' +
    '<div class="tot-row"><span>🔵 Camere in fermata</span><span>'+totFermata+'</span></div>' +
    '<div class="tot-row"><span>🟢 Camere in arrivo</span><span>'+totArrivo+'</span></div>' +
    '<div class="tot-row" style="margin-top:8px"><span>Incasso lordo</span><span>€'+incassoLordo.toFixed(2)+'</span></div>' +
    '<div class="tot-row"><span>Commissioni</span><span>− €'+incassoComm.toFixed(2)+'</span></div>' +
    '<div class="tot-row"><span>Tasse soggiorno</span><span>€'+incassoTasse.toFixed(2)+'</span></div>' +
    '<div class="tot-row" style="font-weight:700;border-top:1px solid rgba(0,0,0,.1);margin-top:6px;padding-top:6px"><span>Incasso netto</span><span>€'+incassoNetto.toFixed(2)+'</span></div>' +
  '</div>';

  return html;
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
      var pag = nd.pagamentoOk ? '✓' : '⚠';
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

