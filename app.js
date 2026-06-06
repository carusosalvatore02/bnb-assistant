const EPOCH = new Date(1899, 11, 30);

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
  document.getElementById('rmodal-print').addEventListener('click', printReport);
  document.getElementById('rmodal-body').addEventListener('click', function(e) {
    if (e.target.classList.contains('cb')) e.target.classList.toggle('checked');
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
  // Apre il gestionale in una nuova tab per scaricare l'Excel
  window.open('https://www.bed-and-breakfast.it/ar/prenotazioni.cfm', '_blank');
  // Mostra il pannello di caricamento file
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
    k = prompt('Inserisci la tua API key Anthropic (da console.anthropic.com):\n\nViene salvata solo sul tuo dispositivo.');
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
  fetch('https://api.anthropic.com/v1/messages',{
    method:'POST',
    headers:{'Content-Type':'application/json','x-api-key':apiKey,'anthropic-version':'2023-06-01','anthropic-dangerous-direct-browser-access':'true'},
    body:JSON.stringify({model:'claude-sonnet-4-20250514',max_tokens:1000,system:sys,messages:chatHistory.slice(-12)})
  }).then(function(r){return r.json();}).then(function(data){
    removeTyping();
    if(data.error){
      addMsg('ai','Errore API: '+data.error.message+'. Controlla la chiave in Impostazioni.');
      if(data.error.type==='authentication_error') localStorage.removeItem('bnb_apikey');
      return;
    }
    var reply=data.content&&data.content[0]?data.content[0].text:'Nessuna risposta.';
    addMsg('ai',reply);chatHistory.push({role:'assistant',content:reply});
  }).catch(function(e){removeTyping();addMsg('ai','Errore di connessione: '+e.message);});
}
function askQ(q){document.getElementById('chat-input').value=q;showTab('chat');sendMsg();}
function resetApiKey(){localStorage.removeItem('bnb_apikey');addMsg('ai','API key rimossa. Alla prossima domanda ti verrà chiesta di nuovo.');}

// ─── REPORTS ─────────────────────────────────────────────
function openReport(type){
  if(!bookings.length){alert('Sincronizza prima i dati.');return;}
  var modal=document.getElementById('rmodal'),body=document.getElementById('rmodal-body'),title=document.getElementById('rmodal-title');
  modal.classList.add('open');
  var today=new Date();today.setHours(0,0,0,0);
  var todayStr=ds(today),label=today.toLocaleDateString('it-IT',{weekday:'long',day:'numeric',month:'long',year:'numeric'});
  var map={colazione:['Colazioni',reportColazione],pulizie:['Pulizie',reportPulizie],reception:['Reception',reportReception],settimanale:['Riepilogo settimanale',reportSettimanale]};
  if(map[type]){title.textContent=map[type][0];body.innerHTML=map[type][1](today,todayStr,label);}
}
function closeReport(){document.getElementById('rmodal').classList.remove('open');}
function printReport(){
  var b=document.getElementById('rmodal-body').innerHTML,t=document.getElementById('rmodal-title').textContent;
  var w=window.open('','_blank');
  w.document.write('<!DOCTYPE html><html><head><title>'+t+'</title><style>body{font-family:sans-serif;padding:20px;font-size:13px}.rstitle{font-weight:600;font-size:11px;text-transform:uppercase;color:#666;margin:14px 0 6px;border-bottom:1px solid #ddd;padding-bottom:4px}.rsrow{display:flex;justify-content:space-between;padding:6px 0;border-bottom:1px solid #eee}.rsbadge{padding:2px 8px;border-radius:10px;font-size:11px}.b-bk{background:#E6F1FB;color:#185FA5}.b-ok{background:#E1F5EE;color:#0F6E56}.b-ko{background:#FAECE7;color:#993C1D}.totbox{background:#f5f5f2;padding:10px;margin-top:8px}.totrow{display:flex;justify-content:space-between;padding:3px 0}.totrow.main{font-weight:700;border-top:1px solid #ddd;margin-top:6px;padding-top:6px}</style></head><body><h2>'+t+'</h2>'+b+'</body></html>');
  w.document.close();w.print();
}
function reportColazione(today,todayStr,label){
  var presenti=bookings.filter(function(b){return b.checkin<=today&&b.checkout>today&&(b.stato==='Attiva'||b.stato==='Modificata');});
  if(!presenti.length)return'<div class="rdate">'+label+'</div><div class="nodata">Nessun ospite presente oggi</div>';
  var totA=presenti.reduce(function(s,b){return s+b.adulti;},0),totB=presenti.reduce(function(s,b){return s+b.bambini;},0);
  var rows=presenti.map(function(b){
    var nights=Math.round((b.checkout-b.checkin)/86400000),done=Math.round((today-b.checkin)/86400000);
    return'<div class="rsrow"><div class="rsleft"><div class="rsname">'+b.nome+' '+b.cognome+'</div><div class="rsdetail">'+b.camera+' · CO: '+fmtDate(b.checkout)+'</div></div><div class="rsright"><div class="rsval">'+b.adulti+' adulti'+(b.bambini?' + '+b.bambini+' bimbi':'')+'</div><div class="rssub">Notte '+done+' di '+nights+'</div></div></div>';
  }).join('');
  return'<div class="rdate">'+label+'</div><div class="rsec"><div class="rstitle">Camere da preparare</div>'+rows+'</div><div class="totbox"><div class="totrow"><span class="tlab">Camere</span><span>'+presenti.length+'</span></div><div class="totrow"><span class="tlab">Adulti</span><span>'+totA+'</span></div>'+(totB?'<div class="totrow"><span class="tlab">Bambini</span><span>'+totB+'</span></div>':'')+'<div class="totrow main"><span>Coperti totali</span><span>'+(totA+totB)+'</span></div></div>';
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
  var cis=bookings.filter(function(b){return ds(b.checkin)===todayStr&&(b.stato==='Attiva'||b.stato==='Modificata');});
  var cos=bookings.filter(function(b){return ds(b.checkout)===todayStr&&(b.stato==='Attiva'||b.stato==='Modificata');});
  var totInc=cis.filter(function(b){return!b.canale.toLowerCase().includes('booking');}).reduce(function(s,b){return s+b.importo+b.tassa;},0);
  var totTasse=cis.reduce(function(s,b){return s+b.tassa;},0);
  var arrRows=cis.map(function(b){
    var isB=b.canale.toLowerCase().includes('booking');
    return'<div class="rsrow"><div class="rsleft"><div class="rsname">'+b.nome+' '+b.cognome+'</div><div class="rsdetail">'+b.camera+' · '+b.adulti+' adulti · CO: '+fmtDate(b.checkout)+'</div><span class="rsbadge '+(isB?'b-bk':'b-si')+'">'+(isB?'Booking — pagato online':'Sito diretto')+'</span></div><div class="rsright"><div class="rsval">€'+b.importo+'</div><div class="rssub">'+(isB?'Tassa: €'+b.tassa:'Da incassare: €'+(b.importo+b.tassa))+'</div></div></div>';
  }).join('');
  var coRows=cos.map(function(b){return'<div class="rsrow"><div class="rsleft"><div class="rsname">'+b.nome+' '+b.cognome+'</div><div class="rsdetail">'+b.camera+'</div></div><div class="rsright"><div class="rsval" style="color:var(--accent-d)">Check-out</div></div></div>';}).join('');
  return'<div class="rdate">'+label+'</div><div class="rsec"><div class="rstitle">Arrivi oggi ('+cis.length+')</div>'+(arrRows||'<div class="nodata">Nessun arrivo</div>')+'</div><div class="rsec"><div class="rstitle">Partenze oggi ('+cos.length+')</div>'+(coRows||'<div class="nodata">Nessuna partenza</div>')+'</div><div class="totbox"><div class="totrow"><span class="tlab">Tasse soggiorno da incassare</span><span>€'+totTasse.toFixed(2)+'</span></div><div class="totrow main"><span>Totale da incassare</span><span>€'+Math.round(totInc).toLocaleString('it-IT')+'</span></div></div>';
}
function reportSettimanale(today,todayStr,label){
  var we=new Date(today);we.setDate(we.getDate()+7);
  var wk=bookings.filter(function(b){return b.checkin>=today&&b.checkin<we&&(b.stato==='Attiva'||b.stato==='Modificata');});
  var active=bookings.filter(function(b){return b.stato==='Attiva'||b.stato==='Modificata';});
  var rev=active.reduce(function(s,b){return s+b.importo;},0),comm=active.reduce(function(s,b){return s+b.commissioni;},0),tasse=active.reduce(function(s,b){return s+b.tassa;},0);
  var netto=rev-comm-tasse;
  var byC={};active.forEach(function(b){byC[b.canale]=(byC[b.canale]||0)+1;});
  var cRows=Object.entries(byC).sort(function(a,b){return b[1]-a[1];}).map(function(e){return'<div class="totrow"><span class="tlab">'+e[0]+'</span><span>'+e[1]+'</span></div>';}).join('');
  var wkRows=wk.map(function(b){return'<div class="rsrow"><div class="rsleft"><div class="rsname">'+b.nome+' '+b.cognome+'</div><div class="rsdetail">'+b.camera+' · CI:'+fmtDate(b.checkin)+' CO:'+fmtDate(b.checkout)+'</div></div><div class="rsright"><div class="rsval">€'+b.importo+'</div><div class="rssub">'+b.adulti+' adulti</div></div></div>';}).join('');
  return'<div class="rdate">'+label+'</div><div class="rsec"><div class="rstitle">Arrivi questa settimana ('+wk.length+')</div>'+(wkRows||'<div class="nodata">Nessun arrivo</div>')+'</div><div class="rsec"><div class="rstitle">Finanziario</div><div class="totbox"><div class="totrow"><span class="tlab">Incasso lordo</span><span>€'+Math.round(rev).toLocaleString('it-IT')+'</span></div><div class="totrow"><span class="tlab">Commissioni</span><span style="color:var(--coral-d)">− €'+Math.round(comm).toLocaleString('it-IT')+'</span></div><div class="totrow"><span class="tlab">Tasse soggiorno</span><span style="color:var(--coral-d)">− €'+Math.round(tasse).toLocaleString('it-IT')+'</span></div><div class="totrow main"><span>Incasso netto</span><span style="color:var(--accent-d)">€'+Math.round(netto).toLocaleString('it-IT')+'</span></div></div></div><div class="rsec"><div class="rstitle">Per canale</div><div class="totbox">'+cRows+'</div></div>';
}
