// ============================================================
// CAPA DE COMUNICACIÓN CON LA API (reemplaza a google.script.run)
// ============================================================
// El content-type "text/plain" es a propósito: evita que el navegador
// mande una petición de "preflight" (OPTIONS) que Apps Script no sabe
// responder. Apps Script igual recibe el texto y lo interpretamos como
// JSON del lado del servidor.
function apiCall(action, params) {
  return fetch(API_URL, {
    method: 'POST',
    headers: { 'Content-Type': 'text/plain;charset=utf-8' },
    body: JSON.stringify({ key: API_KEY, action: action, params: params || {} })
  }).then(function (res) {
    if (!res.ok) throw new Error('Respuesta del servidor: ' + res.status);
    return res.json();
  });
}

// ============================================================
// COLA SIN CONEXIÓN (para "Agregar gasto" cuando no hay señal)
// ============================================================
function leerCola(){
  try { return JSON.parse(localStorage.getItem('colaPendiente') || '[]'); }
  catch(e){ return []; }
}
function guardarCola(cola){ localStorage.setItem('colaPendiente', JSON.stringify(cola)); }
function encolarPendiente(action, params){
  var cola = leerCola();
  cola.push({ action: action, params: params, ts: Date.now() });
  guardarCola(cola);
  actualizarBadgeCola();
}
function actualizarBadgeCola(){
  var n = leerCola().length;
  var el = document.getElementById('badgeCola');
  if(!el) return;
  if(n>0){ el.style.display='block'; el.textContent = n + ' pendiente' + (n>1?'s':'') + ' por sincronizar'; }
  else { el.style.display='none'; }
}
function sincronizarPendientes(){
  var cola = leerCola();
  if(cola.length===0) return;
  var restantes = [];
  var pendientesOriginal = cola.length;
  var procesados = 0;
  cola.forEach(function(item){
    apiCall(item.action, item.params).then(function(resp){
      procesados++;
      if(!resp.ok) restantes.push(item); // si el servidor lo rechazó, lo dejamos para revisión manual
      if(procesados===pendientesOriginal){
        guardarCola(restantes);
        actualizarBadgeCola();
        if(restantes.length < pendientesOriginal){ cargarInicio(); }
      }
    }).catch(function(){
      procesados++;
      restantes.push(item); // sigue sin señal, lo dejamos en la cola
      if(procesados===pendientesOriginal){
        guardarCola(restantes);
        actualizarBadgeCola();
      }
    });
  });
}
window.addEventListener('online', sincronizarPendientes);

// Guarda un gasto; si falla por red, lo encola para enviarlo después.
function registrarGastoConCola(datos){
  return apiCall('registrarGasto', { datos: datos }).catch(function(){
    encolarPendiente('registrarGasto', { datos: datos });
    return { ok: true, mensaje: 'Sin conexión: se guardó en tu celular y se enviará solo cuando vuelva la señal 📶', encolado: true };
  });
}

// ============================================================
// ESTADO GENERAL
// ============================================================
var COLORES = ['#1F3864','#2E7D32','#B7791F','#8E44AD','#1B9C85','#C62828','#6B7280'];
var mesSeleccionadoInicio = null;
var mesSeleccionadoFijos = null;
var mesSeleccionadoAnalisis = null;
var mesesInfo = null;
var fijaActual = null; // {fila, montoPlaneado} en el modal
var ultimoResumen = null;
var mostrarIngresos = localStorage.getItem('mostrarIngresos') === 'true';
var tipoMovimiento = 'gasto';
var editandoFila = null;
var editandoCategoria = null;
var eliminandoFila = null;

// Modo oscuro por defecto (a menos que el usuario ya haya elegido claro antes)
var modoOscuroGuardado = localStorage.getItem('modoOscuro');
var oscuro = modoOscuroGuardado === null ? true : modoOscuroGuardado === 'true';
aplicarModoOscuro();

function aplicarModoOscuro(){
  document.body.classList.toggle('dark', oscuro);
  var chk = document.getElementById('chkOscuro');
  if(chk) chk.checked = oscuro;
}
function toggleModoOscuro(){
  oscuro = document.getElementById('chkOscuro').checked;
  localStorage.setItem('modoOscuro', oscuro);
  aplicarModoOscuro();
}
function abrirConfiguracion(){
  document.getElementById('chkOscuro').checked = oscuro;
  document.getElementById('chkFaceId').checked = localStorage.getItem('faceIdActivo') === 'true';
  document.getElementById('overlayConfig').classList.add('open');
}
function cerrarConfiguracion(){ document.getElementById('overlayConfig').classList.remove('open'); }
function abrirHojaCalculo(){
  apiCall('getSpreadsheetUrl', {}).then(function(url){
    window.open(url, '_blank');
  }).catch(function(err){ alert('Error: '+err.message); });
}

// ============================================================
// FACE ID / TOUCH ID (WebAuthn) — candado local del dispositivo
// ============================================================
function ab2b64(buf){
  var bytes = new Uint8Array(buf);
  var bin = '';
  for(var i=0;i<bytes.byteLength;i++) bin += String.fromCharCode(bytes[i]);
  return btoa(bin);
}
function b642ab(b64){
  var bin = atob(b64);
  var bytes = new Uint8Array(bin.length);
  for(var i=0;i<bin.length;i++) bytes[i] = bin.charCodeAt(i);
  return bytes.buffer;
}

function toggleFaceId(){
  var activar = document.getElementById('chkFaceId').checked;
  if(activar){ activarFaceId(); }
  else { desactivarFaceId(); }
}

function activarFaceId(){
  if(!window.PublicKeyCredential){
    alert('Tu navegador no soporta Face ID / Touch ID en la web.');
    document.getElementById('chkFaceId').checked = false;
    return;
  }
  var challenge = crypto.getRandomValues(new Uint8Array(32));
  var userId = crypto.getRandomValues(new Uint8Array(16));
  navigator.credentials.create({
    publicKey: {
      challenge: challenge,
      rp: { name: 'Control de Gastos' },
      user: { id: userId, name: 'control-gastos', displayName: 'Control de Gastos' },
      pubKeyCredParams: [{ type:'public-key', alg:-7 }, { type:'public-key', alg:-257 }],
      authenticatorSelection: { authenticatorAttachment:'platform', userVerification:'required' },
      timeout: 60000,
      attestation: 'none'
    }
  }).then(function(cred){
    localStorage.setItem('faceIdCredId', ab2b64(cred.rawId));
    localStorage.setItem('faceIdActivo', 'true');
    alert('Face ID / Touch ID activado ✅. La próxima vez que abras la app, te lo va a pedir.');
  }).catch(function(err){
    alert('No se pudo activar: ' + err.message);
    document.getElementById('chkFaceId').checked = false;
  });
}

function desactivarFaceId(){
  localStorage.removeItem('faceIdCredId');
  localStorage.removeItem('faceIdActivo');
}

function verificarFaceId(){
  var credId = localStorage.getItem('faceIdCredId');
  if(!credId) return Promise.resolve(true);
  var challenge = crypto.getRandomValues(new Uint8Array(32));
  return navigator.credentials.get({
    publicKey: {
      challenge: challenge,
      allowCredentials: [{ id: b642ab(credId), type:'public-key' }],
      userVerification: 'required',
      timeout: 60000
    }
  }).then(function(){ return true; }).catch(function(){ return false; });
}

function intentarDesbloqueo(){
  var btn = document.getElementById('btnDesbloquear');
  var err = document.getElementById('lockError');
  err.textContent = '';
  btn.disabled = true; btn.innerHTML = '<span class="spinner"></span>Verificando…';
  verificarFaceId().then(function(ok){
    btn.disabled = false; btn.textContent = '🔓 Desbloquear';
    if(ok){
      document.getElementById('lockScreen').classList.remove('show');
      document.getElementById('appRoot').style.display = 'block';
      iniciarApp();
    } else {
      err.textContent = 'No se pudo verificar. Intenta de nuevo.';
    }
  });
}

function hoyISO(){
  var d = new Date();
  return d.getFullYear() + '-' + String(d.getMonth()+1).padStart(2,'0') + '-' + String(d.getDate()).padStart(2,'0');
}
function fmt(n){
  n = Number(n)||0;
  return '$' + Math.round(n).toLocaleString('es-CO');
}
function skeletonLineas(n, alturas){
  var out = '';
  for(var i=0;i<n;i++){
    var h = (alturas && alturas[i]) || 14;
    var w = 100 - (i*12 % 40);
    out += '<div class="skel skel-line" style="height:'+h+'px; width:'+w+'%;"></div>';
  }
  return out;
}
function animarNumero(el, valorFinal, prefijoClase){
  var reduce = window.matchMedia && window.matchMedia('(prefers-reduced-motion: reduce)').matches;
  if(reduce){ el.textContent = fmt(valorFinal); return; }
  var inicio = 0;
  var duracion = 650;
  var t0 = null;
  function paso(ts){
    if(!t0) t0 = ts;
    var p = Math.min((ts - t0) / duracion, 1);
    var actual = inicio + (valorFinal - inicio) * (1 - Math.pow(1-p, 3)); // ease-out cúbico
    el.textContent = fmt(actual);
    if(p < 1) requestAnimationFrame(paso);
    else el.textContent = fmt(valorFinal);
  }
  requestAnimationFrame(paso);
}

// ---------------- NAV ----------------
var ORDEN_TABS = ['inicio','fijos','agregar','analisis','ahorro'];
function irATab(nombre){
  document.querySelectorAll('.tab').forEach(function(t){ t.classList.remove('active'); });
  document.getElementById('tab-'+nombre).classList.add('active');
  document.querySelectorAll('.navbtn').forEach(function(b){ b.classList.remove('active'); });
  document.querySelector('.navbtn[data-tab="'+nombre+'"]').classList.add('active');
  var ind = document.getElementById('navIndicator');
  if(ind){
    var idx = ORDEN_TABS.indexOf(nombre);
    ind.style.transform = 'translateX(' + (idx*100) + '%)';
  }
  if(nombre === 'analisis') cargarAnalisis();
}

// ---------------- INIT ----------------
window.addEventListener('DOMContentLoaded', function(){
  var faceIdActivo = localStorage.getItem('faceIdActivo') === 'true';
  if(faceIdActivo){
    document.getElementById('lockScreen').classList.add('show');
    document.getElementById('appRoot').style.display = 'none';
  } else {
    iniciarApp();
  }

  if('serviceWorker' in navigator){
    navigator.serviceWorker.register('./service-worker.js').catch(function(){});
  }
});

function iniciarApp(){
  document.getElementById('fecha').value = hoyISO();
  cargarCategorias();
  actualizarBadgeCola();
  sincronizarPendientes();

  apiCall('getMesesDisponibles', {}).then(function(info){
    mesesInfo = info;
    mesSeleccionadoInicio = info.meses[info.actual];
    mesSeleccionadoFijos = info.meses[info.actual];
    mesSeleccionadoAnalisis = info.meses[info.actual];
    cargarInicio();
    cargarFijos();
  }).catch(function(err){
    document.getElementById('cardDisponible').innerHTML = cardDisponibleHead() + '<div class="empty">Sin conexión con la hoja. Revisa tu internet.</div>';
  });
  cargarAhorro();
}

function renderPills(contId, meses, seleccionado, onClick){
  var cont = document.getElementById(contId);
  cont.innerHTML = '';
  meses.forEach(function(m){
    var p = document.createElement('div');
    p.className = 'pill' + (m===seleccionado ? ' active':'');
    p.textContent = m;
    p.onclick = function(){ onClick(m); };
    cont.appendChild(p);
  });
}

// ---------------- INICIO ----------------
function cargarInicio(){
  if(!mesesInfo) return;
  renderPills('pillsInicio', mesesInfo.meses, mesSeleccionadoInicio, function(m){
    mesSeleccionadoInicio = m; cargarInicio();
  });
  document.getElementById('cardDisponible').innerHTML = cardDisponibleHead() + skeletonLineas(3, [34,14,14]);
  document.getElementById('listaUltimos').innerHTML = skeletonLineas(4, [30,30,30,30]);

  apiCall('getResumenMes', { mesAbr: mesSeleccionadoInicio }).then(function(r){
    ultimoResumen = r;
    renderDisponible();
    if(!r.error && r.categorias && r.categorias.length>0){
      document.getElementById('cardChart').style.display='block';
      renderDonut(r.categorias);
    } else {
      document.getElementById('cardChart').style.display='none';
    }
  }).catch(function(err){
    ultimoResumen = { error: 'Sin conexión (' + err.message + ')' };
    renderDisponible();
  });

  apiCall('getUltimosGastos', { n: 6 }).then(function(lista){
    if(!lista || lista.length===0){
      document.getElementById('listaUltimos').innerHTML = '<div class="empty">Aún no has registrado gastos.</div>';
      return;
    }
    document.getElementById('listaUltimos').innerHTML = lista.map(function(tx){
      var esReembolso = Number(tx.monto) < 0;
      return '<div class="tx-item"><div class="tx-left"><div class="cat">'+tx.categoria+'</div>' +
        '<div class="desc">'+(tx.descripcion||'')+'</div></div>' +
        '<div class="tx-right"><div class="monto'+(esReembolso?' refund':'')+'">'+(esReembolso?'+':'')+fmt(Math.abs(tx.monto))+'</div><div class="fecha">'+tx.fecha+'</div></div></div>';
    }).join('');
  }).catch(function(err){
    document.getElementById('listaUltimos').innerHTML = '<div class="empty">Sin conexión.</div>';
  });
}

function cardDisponibleHead(){
  return '<div class="card-head"><h3 style="margin:0;">Disponible este mes</h3>' +
    '<button class="eye-btn" id="btnOjo" onclick="toggleIngresos()">'+(mostrarIngresos?'🙈':'👁️')+'</button></div>';
}

function toggleIngresos(){
  mostrarIngresos = !mostrarIngresos;
  localStorage.setItem('mostrarIngresos', mostrarIngresos);
  renderDisponible();
}

function renderDisponible(){
  var r = ultimoResumen;
  var cont = document.getElementById('cardDisponible');
  if(!r || r.error){
    cont.innerHTML = cardDisponibleHead() + '<div class="empty">'+(r?r.error:'Cargando…')+'</div>';
    return;
  }
  var claseColor = r.restante >= 0 ? 'pos' : 'neg';
  var oculto = '••••••';
  cont.innerHTML = cardDisponibleHead() +
    '<div class="big-number '+claseColor+'" id="numDisponible">$0</div>' +
    '<div class="sub-row"><span>Ingreso</span><b>'+(mostrarIngresos?fmt(r.ingreso):oculto)+'</b></div>' +
    '<div class="sub-row"><span>Ahorro aportado</span><b>'+(mostrarIngresos?fmt(r.ahorro):oculto)+'</b></div>' +
    (r.retiro>0 ? '<div class="sub-row"><span>Retiro del ahorro</span><b>'+fmt(r.retiro)+'</b></div>' : '') +
    '<div class="sub-row"><span>Gasto real</span><b>'+fmt(r.gastoReal)+'</b></div>';
  animarNumero(document.getElementById('numDisponible'), r.restante);
}

function renderDonut(categorias){
  var total = categorias.reduce(function(s,c){ return s+c.valor; }, 0);
  var r = 42, circ = 2*Math.PI*r;
  var acumulado = 0;
  var svg = '<svg width="100" height="100" viewBox="0 0 100 100">';
  svg += '<circle cx="50" cy="50" r="'+r+'" fill="none" style="stroke:var(--border)" stroke-width="14"></circle>';
  categorias.forEach(function(c,i){
    var frac = c.valor/total;
    var largo = frac*circ;
    var offset = circ - acumulado;
    svg += '<circle cx="50" cy="50" r="'+r+'" fill="none" stroke="'+COLORES[i%COLORES.length]+'" stroke-width="14" ' +
      'stroke-dasharray="'+largo+' '+circ+'" stroke-dashoffset="'+offset+'" transform="rotate(-90 50 50)"></circle>';
    acumulado += largo;
  });
  svg += '</svg>';
  document.getElementById('donut').innerHTML = svg;
  document.getElementById('legend').innerHTML = categorias.map(function(c,i){
    return '<div class="legend-item"><span><span class="dot" style="background:'+COLORES[i%COLORES.length]+'"></span>' +
      '<span class="nombre">'+c.nombre+'</span></span><span class="valor">'+fmt(c.valor)+'</span></div>';
  }).join('');
}

// ---------------- FIJOS ----------------
function cargarFijos(){
  if(!mesesInfo) return;
  renderPills('pillsFijos', mesesInfo.meses, mesSeleccionadoFijos, function(m){
    mesSeleccionadoFijos = m; cargarFijos();
  });
  document.getElementById('progresoTexto').textContent = 'Cargando…';
  document.getElementById('listaFijos').innerHTML = skeletonLineas(5, [40,40,40,40,40]);

  apiCall('getFijosEstado', { mesAbr: mesSeleccionadoFijos }).then(function(r){
    if(r.error){
      document.getElementById('listaFijos').innerHTML = '<div class="empty">'+r.error+'</div>';
      return;
    }
    document.getElementById('progresoTexto').textContent = r.pagados + ' de ' + r.total + ' pagados en ' + mesSeleccionadoFijos;
    var pct = r.total ? Math.round((r.pagados/r.total)*100) : 0;
    document.getElementById('progresoFill').style.width = pct + '%';

    if(!r.items || r.items.length===0){
      document.getElementById('listaFijos').innerHTML = '<div class="empty">No hay fijos configurados para este mes.</div>';
      return;
    }
    document.getElementById('listaFijos').innerHTML = r.items.map(function(it){
      var multiParte = it.parte && it.parte.indexOf('/1') === -1;
      var diaTag = '<span class="dia-pago" onclick="abrirModalDia(\''+it.categoria+'\','+(it.diaPago||'null')+')">' +
        (it.diaPago ? 'vence día '+it.diaPago : 'poner día') + '</span>';
      return '<div class="fijo-item">' +
        '<div class="fijo-left"><div class="nombre">'+it.categoria + (multiParte? '<span class="parte">'+it.parte+'</span>':'')+'</div>' +
        '<div class="monto">'+fmt(it.montoPlaneado)+(it.pagado? ' · pagado '+fmt(it.montoPagado)+' ('+it.fechaPago+')':'')+diaTag+'</div></div>' +
        '<div class="fijo-right">' +
        (it.pagado? '' : '<button class="edit-btn" onclick="abrirModalMonto('+it.fila+','+it.montoPlaneado+')">✎</button>') +
        (it.pagado? '' : '<button class="edit-btn" onclick="abrirModalEliminar('+it.fila+')">🗑</button>') +
        '<div class="check-btn'+(it.pagado?' done':'')+'" onclick="clickFijo('+it.fila+','+it.montoPlaneado+','+it.pagado+')">✓</div>' +
        '</div></div>';
    }).join('');
  }).catch(function(err){
    document.getElementById('listaFijos').innerHTML = '<div class="empty">Sin conexión.</div>';
  });
}

function clickFijo(fila, montoPlaneado, pagado){
  if(pagado){
    if(confirm('¿Deshacer este pago? Se borrará también del Registro.')){
      apiCall('desmarcarFijoPagado', { fila: fila }).then(function(){ cargarFijos(); cargarInicio(); })
        .catch(function(err){ alert('Error: '+err.message); });
    }
    return;
  }
  fijaActual = { fila: fila };
  document.getElementById('modalMonto').value = montoPlaneado;
  document.getElementById('modalFecha').value = hoyISO();
  document.getElementById('overlayPago').classList.add('open');
}
function cerrarModalPago(){ document.getElementById('overlayPago').classList.remove('open'); }
function confirmarPago(){
  var monto = document.getElementById('modalMonto').value;
  var fecha = document.getElementById('modalFecha').value;
  var btn = document.getElementById('btnConfirmarPago');
  btn.disabled = true; btn.innerHTML = '<span class="spinner"></span>Guardando…';
  apiCall('marcarFijoPagado', { fila: fijaActual.fila, monto: monto, fechaStr: fecha }).then(function(resp){
    btn.disabled = false; btn.textContent = 'Confirmar';
    cerrarModalPago();
    if(resp.ok){ cargarFijos(); cargarInicio(); }
    else { alert(resp.mensaje); }
  }).catch(function(err){
    btn.disabled = false; btn.textContent = 'Confirmar';
    alert('Error: '+err.message);
  });
}

// --- Editar monto planeado ---
function abrirModalMonto(fila, montoActual){
  editandoFila = fila;
  document.getElementById('montoNuevoInput').value = montoActual;
  document.getElementById('aplicarFuturoInput').checked = false;
  document.getElementById('overlayMonto').classList.add('open');
}
function cerrarModalMonto(){ document.getElementById('overlayMonto').classList.remove('open'); }
function confirmarMonto(){
  var nuevoMonto = document.getElementById('montoNuevoInput').value;
  var aplicarFuturo = document.getElementById('aplicarFuturoInput').checked;
  var btn = document.getElementById('btnConfirmarMonto');
  btn.disabled = true; btn.innerHTML = '<span class="spinner"></span>Guardando…';
  apiCall('actualizarMontoFijo', { fila: editandoFila, nuevoMonto: nuevoMonto, aplicarFuturo: aplicarFuturo }).then(function(resp){
    btn.disabled = false; btn.textContent = 'Guardar';
    cerrarModalMonto();
    if(resp.ok){ cargarFijos(); } else { alert(resp.mensaje); }
  }).catch(function(err){
    btn.disabled = false; btn.textContent = 'Guardar';
    alert('Error: '+err.message);
  });
}

// --- Día de pago ---
function abrirModalDia(categoria, diaActual){
  editandoCategoria = categoria;
  document.getElementById('tituloDia').textContent = 'Día de pago · ' + categoria;
  document.getElementById('diaInput').value = diaActual || '';
  document.getElementById('overlayDia').classList.add('open');
}
function cerrarModalDia(){ document.getElementById('overlayDia').classList.remove('open'); }
function confirmarDia(){
  var dia = document.getElementById('diaInput').value;
  var btn = document.getElementById('btnConfirmarDia');
  btn.disabled = true; btn.innerHTML = '<span class="spinner"></span>Guardando…';
  apiCall('actualizarDiaPago', { categoria: editandoCategoria, dia: dia }).then(function(resp){
    btn.disabled = false; btn.textContent = 'Guardar';
    cerrarModalDia();
    if(resp.ok){ cargarFijos(); } else { alert(resp.mensaje); }
  }).catch(function(err){
    btn.disabled = false; btn.textContent = 'Guardar';
    alert('Error: '+err.message);
  });
}

// --- Agregar pago fijo nuevo ---
function abrirModalAgregarFijo(){
  apiCall('getCategorias', {}).then(function(categorias){
    var sel = document.getElementById('agregarFijoCategoria');
    sel.innerHTML = categorias.map(function(c){ return '<option value="'+c.nombre+'">'+c.nombre+'</option>'; }).join('');
  }).catch(function(){});
  document.getElementById('agregarFijoMonto').value = '';
  document.getElementById('agregarFijoTodos').checked = false;
  document.getElementById('overlayAgregarFijo').classList.add('open');
}
function cerrarModalAgregarFijo(){ document.getElementById('overlayAgregarFijo').classList.remove('open'); }
function confirmarAgregarFijo(){
  var categoria = document.getElementById('agregarFijoCategoria').value;
  var monto = document.getElementById('agregarFijoMonto').value;
  var todos = document.getElementById('agregarFijoTodos').checked;
  if(!monto || Number(monto)<=0){ alert('Escribe un monto válido.'); return; }
  var btn = document.getElementById('btnConfirmarAgregarFijo');
  btn.disabled = true; btn.innerHTML = '<span class="spinner"></span>Guardando…';
  apiCall('agregarFijoEstado', { categoria: categoria, monto: monto, aplicarATodos: todos, mesUnico: mesSeleccionadoFijos }).then(function(resp){
    btn.disabled = false; btn.textContent = 'Agregar';
    cerrarModalAgregarFijo();
    if(resp.ok){ cargarFijos(); } else { alert(resp.mensaje); }
  }).catch(function(err){
    btn.disabled = false; btn.textContent = 'Agregar';
    alert('Error: '+err.message);
  });
}

// --- Eliminar pago fijo ---
function abrirModalEliminar(fila){
  eliminandoFila = fila;
  document.getElementById('overlayEliminar').classList.add('open');
}
function cerrarModalEliminar(){ document.getElementById('overlayEliminar').classList.remove('open'); }
function confirmarEliminarFijo(todos){
  apiCall('eliminarFijoEstado', { fila: eliminandoFila, aplicarATodos: todos }).then(function(resp){
    cerrarModalEliminar();
    if(resp.ok){ cargarFijos(); } else { alert(resp.mensaje); }
  }).catch(function(err){
    cerrarModalEliminar();
    alert('Error: '+err.message);
  });
}

// ---------------- AGREGAR ----------------
function cargarCategorias(seleccionar){
  apiCall('getCategorias', {}).then(function(categorias){
    var sel = document.getElementById('categoria');
    sel.innerHTML = '';
    if(categorias && categorias.length>0){
      categorias.forEach(function(c){
        var opt = document.createElement('option');
        opt.value = c.nombre;
        opt.textContent = c.nombre + (c.tipo? '  ·  '+c.tipo : '');
        sel.appendChild(opt);
      });
    }
    var optNueva = document.createElement('option');
    optNueva.value = '__nueva__';
    optNueva.textContent = '➕ Nueva categoría...';
    sel.appendChild(optNueva);
    if(seleccionar) sel.value = seleccionar;
  }).catch(function(){
    document.getElementById('categoria').innerHTML = '<option value="">Sin conexión</option>';
  });
}
function onCategoriaChange(){
  var val = document.getElementById('categoria').value;
  document.getElementById('nuevaCategoriaPanel').style.display = (val==='__nueva__') ? 'block' : 'none';
}
function crearCategoriaNueva(){
  var nombre = document.getElementById('nuevaCategoriaNombre').value.trim();
  var tipo = document.getElementById('nuevaCategoriaTipo').value;
  if(!nombre){ alert('Escribe un nombre para la categoría.'); return; }
  apiCall('crearCategoria', { nombre: nombre, tipo: tipo, montoInicial: 0 }).then(function(resp){
    if(resp.ok){
      document.getElementById('nuevaCategoriaPanel').style.display = 'none';
      document.getElementById('nuevaCategoriaNombre').value = '';
      cargarCategorias(resp.nombre);
    } else { alert(resp.mensaje); }
  }).catch(function(err){ alert('Error: '+err.message); });
}
function elegirTipo(tipo){
  tipoMovimiento = tipo;
  document.getElementById('btnTipoGasto').classList.toggle('active', tipo==='gasto');
  document.getElementById('btnTipoReembolso').classList.toggle('active', tipo==='reembolso');
  document.getElementById('btnGuardar').textContent = tipo==='reembolso' ? 'Guardar reembolso' : 'Guardar gasto';
  document.getElementById('descripcion').placeholder = tipo==='reembolso' ? '¿Quién te reembolsó / por qué?' : '¿En qué fue?';
}

function mostrarToast(id, msg, ok){
  var t = document.getElementById(id);
  t.textContent = msg; t.className = 'toast ' + (ok?'ok':'err');
}
function guardarGasto(){
  var btn = document.getElementById('btnGuardar');
  var datos = {
    fecha: document.getElementById('fecha').value,
    categoria: document.getElementById('categoria').value,
    descripcion: document.getElementById('descripcion').value,
    monto: document.getElementById('monto').value,
    tipo: tipoMovimiento
  };
  if(!datos.categoria || datos.categoria==='__nueva__'){ mostrarToast('toastAgregar','Elige o crea una categoría primero.', false); return; }
  if(!datos.monto || Number(datos.monto)<=0){ mostrarToast('toastAgregar','Escribe un monto válido.', false); return; }

  var textoOriginal = tipoMovimiento==='reembolso' ? 'Guardar reembolso' : 'Guardar gasto';
  btn.disabled = true; btn.innerHTML = '<span class="spinner"></span>Guardando…';
  registrarGastoConCola(datos).then(function(resp){
    btn.disabled = false; btn.textContent = textoOriginal;
    if(resp.ok){
      mostrarToast('toastAgregar', resp.mensaje, true);
      document.getElementById('descripcion').value = '';
      document.getElementById('monto').value = '';
      document.getElementById('fecha').value = hoyISO();
      if(!resp.encolado) cargarInicio();
      actualizarBadgeCola();
    } else {
      mostrarToast('toastAgregar', resp.mensaje, false);
    }
  }).catch(function(err){
    btn.disabled = false; btn.textContent = textoOriginal;
    mostrarToast('toastAgregar','Error: '+err.message, false);
  });
}

// ---------------- AHORRO ----------------
function cargarAhorro(){
  apiCall('getAhorroResumen', {}).then(function(r){
    if(r.error){ document.getElementById('listaAhorro').innerHTML = '<div class="empty">'+r.error+'</div>'; return; }
    document.getElementById('saldoAhorro').textContent = fmt(r.saldoActual);
    document.getElementById('listaAhorro').innerHTML = r.meses.map(function(m, idx){
      return '<div class="mes-row">' +
        '<div class="mes-row-head" onclick="toggleMesAhorro('+idx+')">' +
          '<span class="nombre">'+m.mes+'</span><span class="saldo">Saldo: '+fmt(m.saldo)+'</span>' +
        '</div>' +
        '<div class="mes-edit" id="ahorroEdit'+idx+'">' +
          '<div class="mes-edit-grid">' +
            '<div class="field"><label>Aportado</label><input type="number" id="ahorroAportado'+idx+'" value="'+m.aportado+'"></div>' +
            '<div class="field"><label>Retiro</label><input type="number" id="ahorroRetiro'+idx+'" value="'+m.retiro+'"></div>' +
          '</div>' +
          '<div class="field"><label>Nota</label><input type="text" id="ahorroNota'+idx+'" value="'+(m.nota||'')+'"></div>' +
          '<button class="btn" onclick="guardarAhorro('+idx+','+m.fila+')">Guardar este mes</button>' +
        '</div>' +
      '</div>';
    }).join('');
  }).catch(function(err){
    document.getElementById('listaAhorro').innerHTML = '<div class="empty">Sin conexión.</div>';
  });
}
function toggleMesAhorro(idx){
  var el = document.getElementById('ahorroEdit'+idx);
  el.classList.toggle('open');
}
function guardarAhorro(idx, fila){
  var aportado = document.getElementById('ahorroAportado'+idx).value;
  var retiro = document.getElementById('ahorroRetiro'+idx).value;
  var nota = document.getElementById('ahorroNota'+idx).value;
  apiCall('actualizarAhorroMes', { fila: fila, aportado: aportado, retiro: retiro, nota: nota }).then(function(resp){
    if(resp.ok){ cargarAhorro(); cargarInicio(); }
    else alert(resp.mensaje);
  }).catch(function(err){ alert('Error: '+err.message); });
}

// ---------------- ANÁLISIS ----------------
function cargarAnalisis(){
  if(!mesesInfo) return;
  renderPills('pillsAnalisis', mesesInfo.meses, mesSeleccionadoAnalisis, function(m){
    mesSeleccionadoAnalisis = m; cargarAnalisis();
  });
  document.getElementById('cardResumenMes').innerHTML = skeletonLineas(3, [30,14,14]);
  document.getElementById('listaRanking').innerHTML = skeletonLineas(5, [30,30,30,30,30]);
  document.getElementById('cardFijoVariable').style.display = 'none';
  document.getElementById('cardTendencia').style.display = 'none';

  apiCall('getAnalisisMes', { mesAbr: mesSeleccionadoAnalisis }).then(function(r){
    if(r.error){
      document.getElementById('cardResumenMes').innerHTML = '<div class="empty">'+r.error+'</div>';
      document.getElementById('listaRanking').innerHTML = '';
      return;
    }
    renderResumenMesAnalisis(r);
    renderFijoVariable(r);
    renderRanking(r);
    renderTendencia(r);
  }).catch(function(err){
    document.getElementById('cardResumenMes').innerHTML = '<div class="empty">Sin conexión.</div>';
    document.getElementById('listaRanking').innerHTML = '';
  });
}

function renderResumenMesAnalisis(r){
  var cmpHtml = '';
  if(r.mesAnteriorTotal !== null && r.mesAnteriorTotal > 0){
    var diff = r.totalGeneral - r.mesAnteriorTotal;
    var pct = Math.round((diff / r.mesAnteriorTotal) * 100);
    if(pct !== 0){
      var subio = pct > 0;
      cmpHtml = '<span class="cmp-badge '+(subio?'up':'down')+'">'+(subio?'▲':'▼')+' '+Math.abs(pct)+'% vs. mes anterior</span>';
    } else {
      cmpHtml = '<span class="cmp-badge down">= que el mes anterior</span>';
    }
  }
  document.getElementById('cardResumenMes').innerHTML =
    '<h3>Gasto total · '+mesSeleccionadoAnalisis+'</h3>' +
    '<div class="big-number" id="numAnalisisTotal">$0</div>' +
    '<div class="sub-row"><span>'+(cmpHtml||'Sin datos del mes anterior')+'</span></div>' +
    '<div class="sub-row"><span>Promedio diario</span><b>'+fmt(r.promedioDiario)+'</b></div>';
  animarNumero(document.getElementById('numAnalisisTotal'), r.totalGeneral);
}

function renderFijoVariable(r){
  var total = (Number(r.totalFijo)||0) + (Number(r.totalVariable)||0);
  if(!total || isNaN(total) || total <= 0){ return; }
  document.getElementById('cardFijoVariable').style.display = 'block';
  var pctFijo = Math.round((r.totalFijo/total)*100);
  var pctVar = 100 - pctFijo;
  document.getElementById('barraFijoVariable').innerHTML =
    '<div class="fv-row"><div class="fv-row-head"><span>Fijos</span><b>'+fmt(r.totalFijo)+' · '+pctFijo+'%</b></div>' +
    '<div class="fv-bar"><div class="fv-fill" style="width:'+pctFijo+'%; background:var(--accent);"></div></div></div>' +
    '<div class="fv-row"><div class="fv-row-head"><span>Variables</span><b>'+fmt(r.totalVariable)+' · '+pctVar+'%</b></div>' +
    '<div class="fv-bar"><div class="fv-fill" style="width:'+pctVar+'%; background:var(--muted);"></div></div></div>';
}

function renderRanking(r){
  if(!r.categorias || r.categorias.length === 0){
    var debugTxt = r.debug ? '<pre style="white-space:pre-wrap; font-size:10px; color:var(--muted); text-align:left; margin-top:10px;">'+JSON.stringify(r.debug, null, 1)+'</pre>' : '';
    document.getElementById('listaRanking').innerHTML = '<div class="empty">Sin gastos registrados este mes.</div>' + debugTxt;
    return;
  }
  var maxGasto = r.categorias[0].gasto;
  document.getElementById('listaRanking').innerHTML = r.categorias.map(function(c){
    var pctBarra = maxGasto ? Math.round((c.gasto/maxGasto)*100) : 0;
    var pctTotal = r.totalGeneral ? Math.round((c.gasto/r.totalGeneral)*100) : 0;
    var excedido = c.presupuesto > 0 && c.gasto > c.presupuesto;
    return '<div class="rank-item">' +
      '<div class="rank-item-head"><span class="nombre">'+c.nombre+'</span>' +
      '<span><b>'+fmt(c.gasto)+'</b> <span class="pct">'+pctTotal+'%</span></span></div>' +
      '<div class="rank-bar"><div class="rank-fill'+(excedido?' excedido':'')+'" style="width:'+pctBarra+'%;"></div></div>' +
      (excedido ? '<div class="rank-excedido-tag">⚠ superó el presupuesto de '+fmt(c.presupuesto)+'</div>' : '') +
      '</div>';
  }).join('');
}

function renderTendencia(r){
  if(!r.tendencia || r.tendencia.length === 0) return;
  document.getElementById('cardTendencia').style.display = 'block';
  var max = Math.max.apply(null, r.tendencia.map(function(t){ return t.total; }).concat([1]));
  document.getElementById('barrasTendencia').innerHTML =
    '<div class="tendencia-bars">' +
    r.tendencia.map(function(t){
      var h = Math.max(4, Math.round((t.total/max)*100));
      var esActual = t.mes === mesSeleccionadoAnalisis;
      return '<div class="tendencia-col">' +
        '<div class="tendencia-bar'+(esActual?' actual':'')+'" style="height:'+h+'%;" title="'+fmt(t.total)+'"></div>' +
        '<div class="tendencia-label">'+t.mes+'</div>' +
        '</div>';
    }).join('') +
    '</div>';
}