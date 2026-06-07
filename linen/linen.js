// ─── CONFIGURAZIONE ───────────────────────────────────────────────────────
var SB_URL = localStorage.getItem('ln_sb_url') || '';
var SB_KEY = localStorage.getItem('ln_sb_key') || '';

var SB = {
  h: function(){
    return {
      'Content-Type':'application/json',
      'apikey': SB_KEY,
      'Authorization':'Bearer ' + SB_KEY,
      'Prefer':'return=representation'
    };
  },
  url: function(t, q){ return SB_URL + '/rest/v1/' + t + (q ? '?' + q : ''); },
  get: async function(t, q){
    var r = await fetch(SB.url(t,q), {headers: SB.h()});
    return r.ok ? r.json() : [];
  },
  upsert: async function(t, data, conflict){
    var url = SB.url(t) + (conflict ? '?on_conflict=' + conflict : '');
    var r = await fetch(url, {
      method:'POST',
      headers: Object.assign({}, SB.h(), {'Prefer':'resolution=merge-duplicates,return=minimal'}),
      body: JSON.stringify(Array.isArray(data) ? data : [data])
    });
    return r;
  },
  insert: async function(t, data){
    var r = await fetch(SB.url(t), {
      method:'POST',
      headers: Object.assign({}, SB.h(), {'Prefer':'return=representation'}),
      body: JSON.stringify(data)
    });
    return r.ok ? r.json() : null;
  },
  del: async function(t, q){
    var r = await fetch(SB.url(t,q), {method:'DELETE', headers: SB.h()});
    return r.ok;
  },
  ok: function(){ return SB_URL && SB_KEY && SB_URL.startsWith('https://'); }
};

// ─── STATO APP ────────────────────────────────────────────────────────────
var currentUser = null;   // {id, nome, ruolo}
var pinBuffer   = '';
var pendingUser = null;
var fotos       = {};     // {camera_id: [base64...]}
var bookings    = [];     // prenotazioni dal DB principale
var rooms       = [];     // camere configurate
var stock       = {};     // {item: qty}
var checklist   = {};     // {camera_id: {confermato, pezzi:{...}, anomalie:[...]}}

// Dotazione standard per camera matrimoniale
var DOTAZIONE = [
  {id:'lenzuola',   label:'Lenzuola matrimoniali', qty:2, set:'letto'},
  {id:'federe',     label:'Federe',                qty:2, set:'letto'},
  {id:'teli_doccia',label:'Teli doccia',           qty:2, set:'bagno'},
  {id:'teli_viso',  label:'Teli viso',             qty:2, set:'bagno'},
  {id:'teli_bidet', label:'Teli bidet',            qty:2, set:'bagno'},
  {id:'tappetino',  label:'Tappetino doccia',      qty:1, set:'bagno'},
  {id:'spazzolini', label:'Spazzolini monouso',    qty:2, set:'arrivo'},
  {id:'cuffie',     label:'Cuffie doccia',         qty:2, set:'arrivo'},
];
var SOGLIA_ALERT = 2; // moltiplicatore max prima di blocco

// Motivi anomalia
var MOTIVI = [
  'Pezzo difettoso/macchiato all\'arrivo',
  'Richiesta extra dell\'ospite',
  'Pezzo lasciato per terra / sporcato in modo evidente',
  'Altro',
];

// ─── INIT ─────────────────────────────────────────────────────────────────
document.addEventListener('DOMContentLoaded', function(){
  bindUI();
  loadUsers();
});

function bindUI(){
  // PIN grid
  var grid = document.getElementById('pin-grid');
  [1,2,3,4,5,6,7,8,9,'',0,'⌫'].forEach(function(k){
    var btn = document.createElement('button');
    btn.className = 'pin-btn';
    btn.textContent = k === '' ? '' : k;
    if(k !== '') btn.addEventListener('click', function(){ onPin(k); });
    grid.appendChild(btn);
  });
  document.getElementById('pin-cancel').addEventListener('click', showLogin);
  document.getElementById('logout-btn').addEventListener('click', logout);

  // Tabs
  document.querySelectorAll('.tab').forEach(function(t){
    t.addEventListener('click', function(){ showTab(t.dataset.tab); });
  });

  // Admin: aggiungi utente/camera
  document.getElementById('btn-add-user').addEventListener('click', addUser);
  document.getElementById('btn-add-room').addEventListener('click', addRoom);
  document.getElementById('btn-save-sb').addEventListener('click', saveSB);
  document.getElementById('btn-export-report').addEventListener('click', exportReport);

  // Foto
  document.getElementById('foto-input').addEventListener('change', onFotoSelected);
}

// ─── SUPABASE CONFIG ──────────────────────────────────────────────────────
function saveSB(){
  var url = document.getElementById('sb-url-input').value.trim();
  var key = document.getElementById('sb-key-input').value.trim();
  if(!url || !key){ toast('Inserisci URL e Key'); return; }
  localStorage.setItem('ln_sb_url', url);
  localStorage.setItem('ln_sb_key', key);
  SB_URL = url; SB_KEY = key;
  document.getElementById('sb-status').textContent = '✓ Salvato';
  toast('Database configurato!');
  loadAll();
}

// ─── LOGIN ────────────────────────────────────────────────────────────────
async function loadUsers(){
  // Utenti locali (localStorage) + eventualmente Supabase
  var localUsers = JSON.parse(localStorage.getItem('ln_users') || '[]');
  // Se non ci sono utenti, crea admin di default
  if(localUsers.length === 0){
    localUsers = [{id:1, nome:'Admin', pin:'0000', ruolo:'admin'}];
    localStorage.setItem('ln_users', JSON.stringify(localUsers));
  }
  renderUserList(localUsers);
}

function renderUserList(users){
  var list = document.getElementById('user-list');
  list.innerHTML = '';
  users.forEach(function(u){
    var btn = document.createElement('button');
    btn.className = 'user-btn';
    var emoji = u.ruolo === 'admin' ? '👑' : '🧹';
    btn.innerHTML = '<div class="user-icon">' + emoji + '</div>' +
      '<div class="user-info"><div class="user-name">' + u.nome + '</div>' +
      '<div class="user-role">' + (u.ruolo === 'admin' ? 'Amministratore' : 'Staff pulizie') + '</div></div>';
    btn.addEventListener('click', function(){ showPIN(u); });
    list.appendChild(btn);
  });
}

function showLogin(){
  document.getElementById('login-screen').style.display = 'flex';
  document.getElementById('pin-screen').style.display = 'none';
  document.getElementById('app-screen').style.display = 'none';
  pinBuffer = '';
  pendingUser = null;
}

function showPIN(user){
  pendingUser = user;
  pinBuffer = '';
  document.getElementById('pin-user-name').textContent = user.nome;
  updatePinDots();
  document.getElementById('login-screen').style.display = 'none';
  document.getElementById('pin-screen').style.display = 'flex';
}

function onPin(k){
  if(k === '⌫'){
    pinBuffer = pinBuffer.slice(0,-1);
  } else if(pinBuffer.length < 4){
    pinBuffer += String(k);
  }
  updatePinDots();
  if(pinBuffer.length === 4) setTimeout(checkPIN, 80);
}

function updatePinDots(){
  var dots = document.querySelectorAll('.pin-dot');
  dots.forEach(function(d, i){
    d.classList.toggle('filled', i < pinBuffer.length);
    d.classList.remove('error');
  });
}

function checkPIN(){
  if(pinBuffer === pendingUser.pin){
    currentUser = pendingUser;
    enterApp();
  } else {
    // Errore: trema i dots
    document.querySelectorAll('.pin-dot').forEach(function(d){ d.classList.add('error'); });
    setTimeout(function(){
      pinBuffer = '';
      updatePinDots();
    }, 700);
    toast('PIN errato');
  }
}

async function enterApp(){
  document.getElementById('pin-screen').style.display = 'none';
  var app = document.getElementById('app-screen');
  app.style.display = 'flex';
  document.getElementById('header-title').textContent = 'Benvenuto, ' + currentUser.nome;
  document.getElementById('header-role').textContent = currentUser.ruolo === 'admin' ? '👑 Amministratore' : '🧹 Staff pulizie';

  // Mostra/nascondi tab Admin
  document.getElementById('tab-admin').style.display = currentUser.ruolo === 'admin' ? 'block' : 'none';
  // Mostra il tasto aggiunta lavanderia solo ad admin
  document.getElementById('add-lavanderia-card').style.display = currentUser.ruolo === 'admin' ? 'block' : 'none';
  document.getElementById('report-riordino-card').style.display = currentUser.ruolo === 'admin' ? 'block' : 'none';

  await loadAll();
  showTab('oggi');

  if(currentUser.ruolo === 'admin' && SB.ok()){
    renderAdminUsers();
    renderAdminRooms();
  }
  if(SB.ok()){
    document.getElementById('sb-url-input').value = SB_URL;
  }
}

function logout(){
  currentUser = null;
  checklist = {};
  fotos = {};
  showLogin();
}

// ─── LOAD DATA ────────────────────────────────────────────────────────────
async function loadAll(){
  await loadRooms();
  await loadStock();
  await loadBookingsFromMainDB();
  buildChecklist();
  renderOggi();
  renderStock();
  renderReport();
}

async function loadRooms(){
  var local = JSON.parse(localStorage.getItem('ln_rooms') || '[]');
  rooms = local;
}

async function loadStock(){
  var local = JSON.parse(localStorage.getItem('ln_stock') || '{}');
  if(Object.keys(local).length === 0){
    // Inizializza stock vuoto
    DOTAZIONE.forEach(function(d){ local[d.id] = 0; });
    localStorage.setItem('ln_stock', JSON.stringify(local));
  }
  stock = local;

  // Se Supabase configurato, prendi da lì
  if(SB.ok()){
    var rows = await SB.get('linen_stock', 'select=item,qty');
    if(rows && rows.length > 0){
      rows.forEach(function(r){ stock[r.item] = r.qty; });
    }
  }
}

async function loadBookingsFromMainDB(){
  if(!SB.ok()) return;
  var today = dsLocal(new Date());
  var rows = await SB.get('bookings',
    'checkout=gte.' + today + '&stato=in.(Attiva,Modificata)&select=codice,camera,checkin,checkout,adulti,bambini,cognome,nome');
  bookings = Array.isArray(rows) ? rows : [];
}

// ─── DATE HELPERS ─────────────────────────────────────────────────────────
function dsLocal(d){
  if(!d) return '';
  var dt = typeof d === 'string' ? new Date(d.includes('T') ? d : d + 'T00:00:00') : d;
  var y = dt.getFullYear(), m = String(dt.getMonth()+1).padStart(2,'0'), day = String(dt.getDate()).padStart(2,'0');
  return y + '-' + m + '-' + day;
}
function today0(){ var d=new Date(); d.setHours(0,0,0,0); return d; }
function parseDate(s){ if(!s) return null; var p=s.split('-'); return new Date(+p[0],+p[1]-1,+p[2]); }
function fmtShort(s){ if(!s) return '—'; var d=parseDate(s); return d.toLocaleDateString('it-IT',{day:'2-digit',month:'2-digit'}); }

// ─── LOGICA BIANCHERIA ────────────────────────────────────────────────────
function calcolaTipoIntervento(booking){
  var today = today0();
  var ci = parseDate(booking.checkin);
  var co = parseDate(booking.checkout);
  if(!ci || !co) return null;

  var isCheckout = dsLocal(co) === dsLocal(today);
  var isCheckin  = dsLocal(ci) === dsLocal(today);

  if(isCheckout || isCheckin) return {tipo:'totale', setLetto:true, setBagno:true, setArrivo:isCheckin};

  // Fermata
  var giorno = Math.round((today - ci) / 86400000); // giorni trascorsi dall'arrivo
  var domani = new Date(today); domani.setDate(domani.getDate()+1);
  var isUltimaNotte = dsLocal(domani) === dsLocal(co);

  var cambioLetto = giorno > 0 && giorno % 4 === 0;
  var cambioBagno = giorno > 0 && giorno % 2 === 0;

  if(!cambioLetto && !cambioBagno) return {tipo:'riassetto', setLetto:false, setBagno:false, setArrivo:false};

  return {
    tipo: cambioLetto && cambioBagno ? 'letto+bagno' : cambioLetto ? 'letto' : 'bagno',
    setLetto: cambioLetto,
    setBagno: cambioBagno,
    setArrivo: false,
    ultimaNotte: isUltimaNotte,
    giorno: giorno
  };
}

function dotazioneTeoria(intervento){
  if(!intervento) return {};
  var pezzi = {};
  DOTAZIONE.forEach(function(d){
    var includi = false;
    if(intervento.tipo === 'totale') includi = true;
    else if(d.set === 'letto' && intervento.setLetto) includi = true;
    else if(d.set === 'bagno' && intervento.setBagno) includi = true;
    else if(d.set === 'arrivo' && intervento.setArrivo) includi = true;
    pezzi[d.id] = includi ? d.qty : 0;
  });
  return pezzi;
}

// ─── CHECKLIST BUILD ──────────────────────────────────────────────────────
function buildChecklist(){
  var today = today0();
  // Per ogni camera, trova la prenotazione attiva oggi
  rooms.forEach(function(room){
    if(checklist[room.id] && checklist[room.id].confermato) return; // già confermata

    var booking = bookings.find(function(b){
      var ci = parseDate(b.checkin), co = parseDate(b.checkout);
      return b.camera === room.nome && ci && co && ci <= today && co >= today;
    });

    if(!booking){ checklist[room.id] = null; return; }

    var intervento = calcolaTipoIntervento(booking);
    var teorici = dotazioneTeoria(intervento);
    var effettivi = {};
    Object.keys(teorici).forEach(function(k){ effettivi[k] = teorici[k]; });

    checklist[room.id] = {
      booking: booking,
      intervento: intervento,
      teorici: teorici,
      effettivi: effettivi,
      anomalie: [],
      note: '',
      confermato: false,
      fotoIds: []
    };
  });
}

// ─── RENDER OGGI ──────────────────────────────────────────────────────────
function renderOggi(){
  var container = document.getElementById('checklist-container');
  if(rooms.length === 0){
    container.innerHTML = '<div style="padding:24px;text-align:center;color:var(--text2);font-size:14px">' +
      (currentUser && currentUser.ruolo === 'admin'
        ? 'Nessuna camera configurata.<br>Aggiungile nel tab <strong>Admin</strong>.'
        : 'Nessuna camera configurata. Contatta l\'admin.') +
      '</div>';
    return;
  }

  var html = '';
  var tutteConfermate = 0, totale = 0;

  rooms.forEach(function(room){
    var entry = checklist[room.id];
    if(!entry){
      // Camera libera oggi
      html += '<div style="margin:12px 14px 0;padding:12px 14px;background:var(--bg2);border-radius:var(--r);border:0.5px solid var(--border)">' +
        '<div style="font-size:15px;font-weight:600">' + room.nome + '</div>' +
        '<div style="font-size:12px;color:var(--text2);margin-top:2px">🟢 Camera libera — nessun intervento</div></div>';
      return;
    }
    totale++;
    if(entry.confermato) tutteConfermate++;

    var iv = entry.intervento;
    var tipoBadge = '', tipoClass = '';
    if(!iv || iv.tipo === 'riassetto'){
      tipoBadge = '✨ Solo riassetto'; tipoClass = 'tipo-riassetto';
    } else if(iv.tipo === 'totale'){
      tipoBadge = '🔄 Cambio totale'; tipoClass = 'tipo-totale';
    } else if(iv.tipo === 'bagno'){
      tipoBadge = '🚿 Solo Set Bagno'; tipoClass = 'tipo-bagno';
    } else if(iv.tipo === 'letto'){
      tipoBadge = '🛏 Solo Set Letto'; tipoClass = 'tipo-letto';
    } else {
      tipoBadge = '🛏🚿 Set Letto + Bagno'; tipoClass = 'tipo-totale';
    }

    var b = entry.booking;
    var confBtn = entry.confermato
      ? '<button class="confirm-btn confirmed" disabled>✓ Confermato</button>'
      : '<button class="confirm-btn" data-room="' + room.id + '">Conferma intervento</button>';

    // Ultima notte warning
    var ultNotte = (iv && iv.ultimaNotte && (iv.setLetto || iv.setBagno))
      ? '<div style="margin:8px 0;padding:8px;background:var(--amber-l);border-radius:var(--r-sm);font-size:12px;font-weight:600;color:var(--amber-d)">⚠ ULTIMA NOTTE — Cambiare solo se in pessime condizioni</div>'
      : '';

    // Righe dotazione (solo pezzi > 0)
    var dotRows = '';
    if(iv && iv.tipo !== 'riassetto'){
      DOTAZIONE.forEach(function(d){
        if(entry.teorici[d.id] === 0) return;
        var eff = entry.effettivi[d.id] || 0;
        var teorico = entry.teorici[d.id];
        var diffColor = eff !== teorico ? 'var(--coral)' : 'var(--text)';
        dotRows += '<div class="dotaz-row">' +
          '<span class="dotaz-item">' + d.label + '</span>' +
          '<div class="dotaz-qty">' +
            '<button class="qty-btn" data-room="' + room.id + '" data-item="' + d.id + '" data-op="minus">−</button>' +
            '<span class="qty-val" style="color:' + diffColor + '" id="qty-' + room.id + '-' + d.id + '">' + eff + '</span>' +
            '<button class="qty-btn" data-room="' + room.id + '" data-item="' + d.id + '" data-op="plus">+</button>' +
            '<span class="qty-teorico">/' + teorico + '</span>' +
          '</div></div>';
      });
    }

    // Sezione anomalia
    var anomSection = '<div class="anomalia-section" id="anom-' + room.id + '">' +
      '<div class="anomalia-label">⚠ Motivo della variazione</div>' +
      '<select class="anomalia-select" id="anom-sel-' + room.id + '">' +
        '<option value="">— Seleziona motivo —</option>' +
        MOTIVI.map(function(m){ return '<option value="' + m + '">' + m + '</option>'; }).join('') +
      '</select>' +
      '<button class="foto-btn" data-room="' + room.id + '"><span>📷</span> Scatta/allega foto (obbligatoria per difetti)</button>' +
      '<div class="foto-preview" id="foto-prev-' + room.id + '"></div>' +
    '</div>';

    html += '<div class="cam-card">' +
      '<div class="cam-header">' +
        '<div>' +
          '<div class="cam-name">' + room.nome + '</div>' +
          '<div class="cam-ospite">' + b.cognome + ' ' + b.nome +
            ' · CI ' + fmtShort(b.checkin) + ' CO ' + fmtShort(b.checkout) + '</div>' +
        '</div>' +
        '<span class="tipo-badge ' + tipoClass + '">' + tipoBadge + '</span>' +
      '</div>' +
      '<div class="cam-body">' +
        ultNotte +
        (dotRows
          ? '<div class="dotaz-section"><div class="dotaz-label">Dotazione da portare</div>' + dotRows + '</div>' + anomSection
          : '<div style="font-size:13px;color:var(--text2)">Solo riassetto, nessuna biancheria da portare.</div>') +
        '<textarea class="note-input" id="note-' + room.id + '" placeholder="Note aggiuntive (opzionale)…">' + (entry.note||'') + '</textarea>' +
        confBtn +
      '</div></div>';
  });

  // Riepilogo in cima
  var riepilogo = totale > 0
    ? '<div style="margin:12px 14px 0;padding:10px 14px;background:var(--accent-l);border-radius:var(--r-sm);font-size:13px;font-weight:500;color:var(--accent-d)">' +
      '✓ ' + tutteConfermate + ' / ' + totale + ' camere confermate oggi</div>'
    : '';

  container.innerHTML = riepilogo + html;
  bindChecklistEvents();
}

function bindChecklistEvents(){
  // Qty buttons
  document.querySelectorAll('.qty-btn').forEach(function(btn){
    btn.addEventListener('click', function(){
      var room = btn.dataset.room, item = btn.dataset.item, op = btn.dataset.op;
      var entry = checklist[room];
      if(!entry || entry.confermato) return;
      var cur = entry.effettivi[item] || 0;
      if(op === 'plus') cur++;
      else if(op === 'minus' && cur > 0) cur--;

      // Blocco soglia
      var teorico = entry.teorici[item] || 0;
      if(cur > teorico * SOGLIA_ALERT && cur > 0){
        toast('⚠ Quantità troppo alta — richiede approvazione admin');
        return;
      }
      entry.effettivi[item] = cur;

      // Aggiorna display
      var el = document.getElementById('qty-' + room + '-' + item);
      if(el){
        el.textContent = cur;
        el.style.color = cur !== teorico ? 'var(--coral)' : 'var(--text)';
      }

      // Mostra/nascondi sezione anomalia
      checkAnomalia(room);
    });
  });

  // Foto button
  document.querySelectorAll('.foto-btn').forEach(function(btn){
    btn.addEventListener('click', function(){
      var room = btn.dataset.room;
      var fotoInput = document.getElementById('foto-input');
      fotoInput.dataset.room = room;
      fotoInput.click();
    });
  });

  // Conferma
  document.querySelectorAll('.confirm-btn:not([disabled])').forEach(function(btn){
    btn.addEventListener('click', function(){
      var room = btn.dataset.room;
      confermaCamera(room);
    });
  });
}

function checkAnomalia(roomId){
  var entry = checklist[roomId];
  if(!entry) return;
  var hasDiff = Object.keys(entry.effettivi).some(function(k){
    return entry.effettivi[k] !== entry.teorici[k];
  });
  var anomSec = document.getElementById('anom-' + roomId);
  if(anomSec) anomSec.style.display = hasDiff ? 'block' : 'none';
}

async function confermaCamera(roomId){
  var entry = checklist[roomId];
  if(!entry) return;

  // Verifica anomalie
  var hasDiff = Object.keys(entry.effettivi).some(function(k){
    return entry.effettivi[k] !== entry.teorici[k];
  });

  if(hasDiff){
    var motivo = document.getElementById('anom-sel-' + roomId)?.value;
    if(!motivo){ toast('Seleziona il motivo della variazione'); return; }

    // Foto obbligatoria per difetti
    if(motivo === 'Pezzo difettoso/macchiato all\'arrivo'){
      if(!fotos[roomId] || fotos[roomId].length === 0){
        toast('📷 Obbligatoria foto per pezzi difettosi');
        return;
      }
    }

    entry.anomalie.push({
      motivo: motivo,
      effettivi: Object.assign({}, entry.effettivi),
      teorici: Object.assign({}, entry.teorici),
      utente: currentUser.nome,
      ts: new Date().toISOString(),
      foto: fotos[roomId] || []
    });
  }

  // Salva nota
  entry.note = document.getElementById('note-' + roomId)?.value || '';
  entry.confermato = true;

  // Scarica magazzino
  aggiornaStock(entry.effettivi);

  // Salva log su Supabase
  if(SB.ok()) await salvaLog(roomId, entry);

  toast('✓ Camera ' + (rooms.find(function(r){return r.id==roomId;})||{}).nome + ' confermata!');
  renderOggi();
  renderStock();
  renderReport();
}

function aggiornaStock(effettivi){
  Object.keys(effettivi).forEach(function(k){
    stock[k] = Math.max(0, (stock[k]||0) - effettivi[k]);
  });
  localStorage.setItem('ln_stock', JSON.stringify(stock));
  if(SB.ok()){
    Object.keys(stock).forEach(function(k){
      SB.upsert('linen_stock', {item:k, qty:stock[k]}, 'item').catch(function(){});
    });
  }
}

async function salvaLog(roomId, entry){
  var room = rooms.find(function(r){return r.id==roomId;});
  var log = {
    data: dsLocal(new Date()),
    camera: room ? room.nome : roomId,
    utente: currentUser.nome,
    tipo_intervento: entry.intervento ? entry.intervento.tipo : 'riassetto',
    pezzi_teorici: JSON.stringify(entry.teorici),
    pezzi_effettivi: JSON.stringify(entry.effettivi),
    anomalie: JSON.stringify(entry.anomalie),
    note: entry.note
  };
  await SB.insert('linen_log', log);
}

function onFotoSelected(e){
  var roomId = e.target.dataset.room;
  if(!fotos[roomId]) fotos[roomId] = [];
  Array.from(e.target.files).forEach(function(file){
    var reader = new FileReader();
    reader.onload = function(ev){
      fotos[roomId].push(ev.target.result);
      renderFotoPreview(roomId);
    };
    reader.readAsDataURL(file);
  });
  e.target.value = '';
}

function renderFotoPreview(roomId){
  var prev = document.getElementById('foto-prev-' + roomId);
  if(!prev) return;
  prev.innerHTML = (fotos[roomId]||[]).map(function(src){
    return '<img class="foto-thumb" src="' + src + '">';
  }).join('');
}

// ─── MAGAZZINO ────────────────────────────────────────────────────────────
function renderStock(){
  var list = document.getElementById('stock-list');
  var html = '';
  DOTAZIONE.forEach(function(d){
    var qty = stock[d.id] || 0;
    var color = qty < 5 ? 'var(--coral)' : qty < 15 ? 'var(--amber-d)' : 'var(--accent-d)';
    var isAdmin = currentUser && currentUser.ruolo === 'admin';
    html += '<div class="stock-row">' +
      '<span class="stock-item">' + d.label + '</span>' +
      '<div class="stock-actions">' +
        (isAdmin ? '<button class="qty-btn" data-stock-item="' + d.id + '" data-stock-op="minus">−</button>' : '') +
        '<span class="stock-qty" style="color:' + color + '">' + qty + '</span>' +
        (isAdmin ? '<button class="qty-btn" data-stock-item="' + d.id + '" data-stock-op="plus">+</button>' : '') +
      '</div></div>';
  });

  // Form arrivo lavanderia per admin
  var lavForm = '';
  if(currentUser && currentUser.ruolo === 'admin'){
    lavForm = '<div style="margin-top:14px">' +
      '<div class="form-row"><label class="form-label">Seleziona articolo</label>' +
        '<select class="form-input" id="lav-item">' +
          DOTAZIONE.map(function(d){ return '<option value="' + d.id + '">' + d.label + '</option>'; }).join('') +
        '</select></div>' +
      '<div class="form-row"><label class="form-label">Quantità arrivata</label>' +
        '<input class="form-input" id="lav-qty" type="number" min="1" placeholder="es. 20"></div>' +
      '<button class="btn-primary" id="btn-add-lav">+ Registra arrivo</button></div>';
  }
  document.getElementById('lavanderia-form').innerHTML = lavForm;
  list.innerHTML = html;

  // Bind stock buttons
  document.querySelectorAll('[data-stock-item]').forEach(function(btn){
    btn.addEventListener('click', function(){
      var item = btn.dataset.stockItem, op = btn.dataset.stockOp;
      if(op === 'plus') stock[item] = (stock[item]||0) + 1;
      else stock[item] = Math.max(0,(stock[item]||0) - 1);
      localStorage.setItem('ln_stock', JSON.stringify(stock));
      if(SB.ok()) SB.upsert('linen_stock', {item:item, qty:stock[item]}, 'item').catch(function(){});
      renderStock();
    });
  });

  var btnLav = document.getElementById('btn-add-lav');
  if(btnLav) btnLav.addEventListener('click', registraLavanderia);
}

function registraLavanderia(){
  var item = document.getElementById('lav-item')?.value;
  var qty = parseInt(document.getElementById('lav-qty')?.value) || 0;
  if(!item || qty <= 0){ toast('Inserisci articolo e quantità'); return; }
  stock[item] = (stock[item]||0) + qty;
  localStorage.setItem('ln_stock', JSON.stringify(stock));
  if(SB.ok()) SB.upsert('linen_stock', {item:item, qty:stock[item]}, 'item').catch(function(){});
  var nome = DOTAZIONE.find(function(d){return d.id===item;})?.label || item;
  toast('+' + qty + ' ' + nome + ' registrati');
  renderStock();
}

// ─── REPORT ───────────────────────────────────────────────────────────────
function renderReport(){
  // Report giornaliero (da checklist confermata)
  var daily = document.getElementById('report-daily');
  var confermati = Object.keys(checklist).filter(function(k){ return checklist[k] && checklist[k].confermato; });
  if(confermati.length === 0){
    daily.innerHTML = '<div style="color:var(--text2);font-size:13px">Nessuna camera confermata oggi.</div>';
  } else {
    var totPezzi = {};
    DOTAZIONE.forEach(function(d){ totPezzi[d.id] = 0; });
    confermati.forEach(function(k){
      var eff = checklist[k].effettivi;
      Object.keys(eff).forEach(function(p){ totPezzi[p] = (totPezzi[p]||0) + eff[p]; });
    });
    daily.innerHTML = '<table style="width:100%;border-collapse:collapse">' +
      DOTAZIONE.map(function(d){
        return '<tr><td style="padding:5px 0;font-size:13px">' + d.label + '</td>' +
          '<td style="text-align:right;font-size:13px;font-weight:600">' + totPezzi[d.id] + '</td></tr>';
      }).join('') + '</table>';
  }

  // Anomalie
  var anomDiv = document.getElementById('report-anomalie');
  var tutteAnom = [];
  Object.keys(checklist).forEach(function(k){
    if(checklist[k] && checklist[k].anomalie.length > 0){
      var room = rooms.find(function(r){return r.id==k;});
      checklist[k].anomalie.forEach(function(a){
        tutteAnom.push({camera: room ? room.nome : k, anom: a});
      });
    }
  });
  if(tutteAnom.length === 0){
    anomDiv.innerHTML = '<div style="color:var(--text2);font-size:13px">Nessuna anomalia oggi.</div>';
  } else {
    anomDiv.innerHTML = tutteAnom.map(function(x){
      return '<div class="log-entry">' +
        '<div class="log-anomalia">📍 ' + x.camera + '</div>' +
        '<div style="font-size:12px;margin-top:2px">' + x.anom.motivo + '</div>' +
        '<div class="log-date">' + x.anom.utente + ' · ' + new Date(x.anom.ts).toLocaleTimeString('it-IT',{hour:'2-digit',minute:'2-digit'}) + '</div>' +
        (x.anom.foto.length ? '<div style="font-size:11px;color:var(--accent-d)">📷 ' + x.anom.foto.length + ' foto allegate</div>' : '') +
      '</div>';
    }).join('');
  }

  // Riordino (solo admin)
  if(currentUser && currentUser.ruolo === 'admin'){
    var rior = document.getElementById('report-riordino');
    // Stima: giacenze attuali + consumo previsto prossimi 3 giorni
    rior.innerHTML = '<div style="font-size:12px;color:var(--text2);margin-bottom:8px">Giacenze attuali vs scorta minima consigliata (15 pz)</div>' +
      DOTAZIONE.map(function(d){
        var qty = stock[d.id]||0;
        var alert = qty < 15;
        return '<div style="display:flex;justify-content:space-between;padding:4px 0;font-size:13px;color:' + (alert?'var(--coral)':'var(--text)') + '">' +
          '<span>' + d.label + (alert?' ⚠':'') + '</span>' +
          '<span>' + qty + ' pz</span></div>';
      }).join('');
  }
}

function exportReport(){
  var html = '<h2>Report Biancheria — ' + new Date().toLocaleDateString('it-IT') + '</h2>' +
    document.getElementById('report-daily').outerHTML +
    document.getElementById('report-anomalie').outerHTML +
    (document.getElementById('report-riordino')||{}).outerHTML;
  var blob = new Blob(['<!DOCTYPE html><html><body style="font-family:sans-serif;padding:20px">' + html + '</body></html>'], {type:'text/html'});
  var file = new File([blob], 'biancheria-' + dsLocal(new Date()) + '.html', {type:'text/html'});
  if(navigator.share && navigator.canShare && navigator.canShare({files:[file]})){
    navigator.share({files:[file], title:'Report Biancheria'}).catch(function(){downloadBlob(blob);});
  } else { downloadBlob(blob); }
}

function downloadBlob(blob){
  var a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = 'biancheria-' + dsLocal(new Date()) + '.html';
  a.click();
}

// ─── ADMIN ────────────────────────────────────────────────────────────────
function renderAdminUsers(){
  var users = JSON.parse(localStorage.getItem('ln_users')||'[]');
  var list = document.getElementById('users-list');
  list.innerHTML = users.map(function(u){
    return '<div style="display:flex;align-items:center;justify-content:space-between;padding:8px 0;border-bottom:0.5px solid var(--border)">' +
      '<div><div style="font-size:14px;font-weight:500">' + u.nome + '</div>' +
        '<div style="font-size:11px;color:var(--text2)">' + (u.ruolo==='admin'?'Admin':'Staff') + ' · PIN: ' + u.pin + '</div></div>' +
      (users.length > 1 ? '<button class="btn-danger" style="padding:6px 12px;font-size:12px" data-del-user="' + u.id + '">✕</button>' : '') +
    '</div>';
  }).join('');
  document.querySelectorAll('[data-del-user]').forEach(function(btn){
    btn.addEventListener('click', function(){
      delUser(parseInt(btn.dataset.delUser));
    });
  });
}

function addUser(){
  var nome = document.getElementById('new-user-name').value.trim();
  var pin  = document.getElementById('new-user-pin').value.trim();
  var ruolo= document.getElementById('new-user-role').value;
  if(!nome || pin.length !== 4 || isNaN(pin)){ toast('Nome e PIN 4 cifre obbligatori'); return; }
  var users = JSON.parse(localStorage.getItem('ln_users')||'[]');
  var maxId = users.reduce(function(m,u){return Math.max(m,u.id);},0);
  users.push({id:maxId+1, nome:nome, pin:pin, ruolo:ruolo});
  localStorage.setItem('ln_users', JSON.stringify(users));
  document.getElementById('new-user-name').value = '';
  document.getElementById('new-user-pin').value = '';
  renderAdminUsers();
  renderUserList(users);
  toast('Utente ' + nome + ' aggiunto');
}

function delUser(id){
  var users = JSON.parse(localStorage.getItem('ln_users')||'[]').filter(function(u){return u.id!==id;});
  localStorage.setItem('ln_users', JSON.stringify(users));
  renderAdminUsers();
  renderUserList(users);
}

function renderAdminRooms(){
  var list = document.getElementById('rooms-list');
  list.innerHTML = rooms.map(function(r){
    return '<div style="display:flex;align-items:center;justify-content:space-between;padding:8px 0;border-bottom:0.5px solid var(--border)">' +
      '<span style="font-size:14px">' + r.nome + '</span>' +
      '<button class="btn-danger" style="padding:6px 12px;font-size:12px" data-del-room="' + r.id + '">✕</button>' +
    '</div>';
  }).join('');
  document.querySelectorAll('[data-del-room]').forEach(function(btn){
    btn.addEventListener('click', function(){ delRoom(parseInt(btn.dataset.delRoom)); });
  });
}

function addRoom(){
  var nome = document.getElementById('new-room-name').value.trim();
  if(!nome){ toast('Inserisci il nome della camera'); return; }
  var maxId = rooms.reduce(function(m,r){return Math.max(m,r.id);},0);
  rooms.push({id:maxId+1, nome:nome});
  localStorage.setItem('ln_rooms', JSON.stringify(rooms));
  document.getElementById('new-room-name').value = '';
  renderAdminRooms();
  buildChecklist();
  renderOggi();
  toast('Camera ' + nome + ' aggiunta');
}

function delRoom(id){
  rooms = rooms.filter(function(r){return r.id!==id;});
  localStorage.setItem('ln_rooms', JSON.stringify(rooms));
  renderAdminRooms();
  renderOggi();
}

// ─── UI HELPERS ───────────────────────────────────────────────────────────
function showTab(name){
  document.querySelectorAll('.tab').forEach(function(t){
    t.classList.toggle('active', t.dataset.tab === name);
  });
  document.querySelectorAll('.tab-content').forEach(function(c){
    c.classList.remove('active');
  });
  var tc = document.getElementById('tab-' + name);
  if(tc) tc.classList.add('active');
}

function toast(msg){
  var t = document.getElementById('toast');
  t.textContent = msg;
  t.classList.add('show');
  setTimeout(function(){ t.classList.remove('show'); }, 2500);
}
