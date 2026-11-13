/* Add-in MyGeotab: Mapa de asientos PID (Cesetti) */
geotab.addin.asientosPid = function () {
  'use strict';

  var ADDIN_ID = 'ad6jPCvZGTvCRcGDo4yXoiQ';
  var XLSX_URL = 'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js';
  var TIPOS = { camioneta: 'Camioneta', combi: 'Combi', colectivo: 'Colectivo' };
  var ASIENTOS_DEF = { camioneta: 4, combi: 15, colectivo: 45 };
  var CFG_DEF = { descensoMin: 30, desbloqueoSeg: 60, modo: 'pid', resetIgnicion: false, ventanaHoras: 12, refrescoSeg: 5 };

  var api = null;
  var root = null;
  var timer = null;
  var cargando = false;

  var S = {
    cfg: Object.assign({}, CFG_DEF),
    cfgRec: null,              // {id}
    pasajeros: {},             // hex -> {id, d}
    vehiculos: {},             // deviceId -> {id, d:{tipo:'vehiculo', deviceId, clase, asientos}}
    devices: {},               // deviceId -> {id, name, licensePlate}
    lecturas: {},              // deviceId -> [{t, tipo, hex?, texto?, id}]
    ignOff: {},                // deviceId -> [t]
    conductores: {},           // deviceId -> userId
    usuarios: {},              // userId -> nombre
    usuariosTarjeta: {},       // hex -> {userId, nombre, legajo} (usuarios de Geotab con llave)
    usuariosLeidos: 0,         // timestamp de la última lectura de usuarios
    condLeidos: 0,
    ignLeida: 0,
    feedVersion: null,
    vista: 'flota',
    sel: null,
    filtroFlota: '',
    filtroPas: '',
    filtroVeh: '',
    hist: { desde: '', hasta: '', dev: '', q: '', evento: '', filas: null, cargando: false, aviso: '' },
    actualizado: null,
    error: null
  };

  /* ---------- utilidades ---------- */
  function call(method, params) {
    return new Promise(function (res, rej) { api.call(method, params, res, rej); });
  }
  function multi(calls) {
    if (!calls.length) return Promise.resolve([]);
    var chunks = [];
    for (var i = 0; i < calls.length; i += 100) chunks.push(calls.slice(i, i + 100));
    return chunks.reduce(function (p, ch) {
      return p.then(function (acc) {
        return new Promise(function (res, rej) {
          api.multiCall(ch, function (r) { res(acc.concat(r)); }, rej);
        });
      });
    }, Promise.resolve([]));
  }
  function esc(s) {
    return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
    });
  }
  function hora(t) {
    if (!t) return '';
    var d = new Date(t);
    var hoy = new Date();
    var hh = d.toLocaleTimeString('es-AR', { hour: '2-digit', minute: '2-digit', hourCycle: 'h23' });
    return d.toDateString() === hoy.toDateString() ? hh : d.toLocaleDateString('es-AR', { day: '2-digit', month: '2-digit' }) + ' ' + hh;
  }
  function iniciales(nombre) {
    var p = String(nombre || '').trim().split(/\s+/).filter(Boolean);
    if (!p.length) return '?';
    return ((p[0][0] || '') + (p.length > 1 ? p[p.length - 1][0] : (p[0][1] || ''))).toUpperCase();
  }
  function errMsg(e) {
    if (!e) return 'Error desconocido';
    if (typeof e === 'string') return e;
    return e.message || (e.errors && e.errors[0] && e.errors[0].message) || JSON.stringify(e);
  }
  function detalles(rec) {
    if (rec.details && typeof rec.details === 'object') return rec.details;
    if (typeof rec.data === 'string') { try { return JSON.parse(rec.data); } catch (e) { return {}; } }
    return {};
  }
  function $(sel) { return root.querySelector(sel); }
  function toast(msg, tipo) {
    var el = document.createElement('div');
    el.className = 'ap-toast ' + (tipo || '');
    el.textContent = msg;
    root.appendChild(el);
    setTimeout(function () { el.remove(); }, 3500);
  }

  /* ---------- AddInData ---------- */
  function guardarRec(rec, d) {
    var entity = { addInId: ADDIN_ID, groups: [{ id: 'GroupCompanyId' }], details: d };
    if (rec && rec.id) {
      entity.id = rec.id;
      return call('Set', { typeName: 'AddInData', entity: entity }).then(function () { return { id: rec.id, d: d }; });
    }
    return call('Add', { typeName: 'AddInData', entity: entity }).then(function (id) { return { id: id, d: d }; });
  }
  function borrarRec(rec) {
    return call('Remove', { typeName: 'AddInData', entity: { id: rec.id, addInId: ADDIN_ID } });
  }

  function cargarAddInData() {
    return call('Get', { typeName: 'AddInData', search: { addInId: ADDIN_ID } }).then(function (recs) {
      S.pasajeros = {}; S.vehiculos = {}; S.cfgRec = null; S.cfg = Object.assign({}, CFG_DEF);
      (recs || []).forEach(function (r) {
        var d = detalles(r);
        if (d.tipo === 'pasajero' && d.hex) S.pasajeros[PidLogic.normHex(d.hex)] = { id: r.id, d: d };
        else if (d.tipo === 'vehiculo' && d.deviceId) S.vehiculos[d.deviceId] = { id: r.id, d: d };
        else if (d.tipo === 'config') { S.cfgRec = { id: r.id }; S.cfg = Object.assign({}, CFG_DEF, d.valores || {}); }
      });
    });
  }

  function cargarDevices() {
    return call('Get', {
      typeName: 'Device',
      search: { fromDate: new Date().toISOString() },
      propertySelector: { fields: ['id', 'name', 'licensePlate', 'serialNumber'], isIncluded: true }
    }).then(function (ds) {
      S.devices = {};
      (ds || []).forEach(function (d) { S.devices[d.id] = d; });
    });
  }

  /* Usuarios de Geotab con llaves: el serialNumber de la llave es el hex de la tarjeta RFID. */
  function cargarUsuariosTarjeta() {
    return call('Get', {
      typeName: 'User',
      search: { isDriver: true, fromDate: new Date().toISOString() },
      propertySelector: { fields: ['id', 'name', 'firstName', 'lastName', 'employeeNo', 'keys'], isIncluded: true }
    }).then(function (us) {
      var mapa = {};
      (us || []).forEach(function (u) {
        (u.keys || []).forEach(function (k) {
          var hex = PidLogic.normHex(k && k.serialNumber);
          if (hex.length < 8) return;
          mapa[hex] = {
            userId: u.id,
            nombre: ((u.firstName || '') + ' ' + (u.lastName || '')).trim() || u.name,
            usuario: u.name,
            legajo: u.employeeNo || '',
            tipoLlave: k.driverKeyType || ''
          };
        });
      });
      S.usuariosTarjeta = mapa;
      S.usuariosLeidos = Date.now();
    });
  }

  /* Quién es una tarjeta: primero el usuario de Geotab con esa llave, si no el padrón del add-in. */
  function persona(hex) {
    var u = S.usuariosTarjeta[hex];
    if (u) return { nombre: u.nombre, legajo: u.legajo, origen: 'geotab', userId: u.userId, usuario: u.usuario };
    var p = S.pasajeros[hex];
    if (p) return { nombre: p.d.nombre, legajo: p.d.legajo, empresa: p.d.empresa, turno: p.d.turno, origen: 'padron' };
    return null;
  }

  /* ---------- datos en vivo ---------- */
  function desdeVentana() { return Date.now() - (Number(S.cfg.ventanaHoras) || 12) * 3600000; }

  function leerFeed() {
    var params = { typeName: 'CustomData', resultsLimit: 5000 };
    if (S.feedVersion) params.fromVersion = S.feedVersion;
    else params.search = { fromDate: new Date(desdeVentana()).toISOString() };
    return call('GetFeed', params).then(function (r) {
      var data = (r && r.data) || [];
      data.forEach(function (cd) {
        var devId = cd.device && cd.device.id;
        if (!devId) return;
        var dec = PidLogic.decodeCustomData(cd.data);
        var arr = S.lecturas[devId] || (S.lecturas[devId] = []);
        if (arr.some(function (x) { return x.id === cd.id; })) return;
        arr.push({ id: cd.id, t: new Date(cd.dateTime).getTime(), tipo: dec.tipo === 'tarjeta' ? 'tarjeta' : 'texto', hex: dec.hex, texto: dec.texto || (dec.tipo === 'otro' ? 'Trama ' + dec.hex : '') });
      });
      S.feedVersion = r && r.toVersion;
      if (data.length >= 5000) return leerFeed();
    });
  }

  function podarVentana() {
    var desde = desdeVentana();
    Object.keys(S.lecturas).forEach(function (k) {
      S.lecturas[k] = S.lecturas[k].filter(function (x) { return x.t >= desde; });
    });
  }

  function leerConductores() {
    return call('Get', { typeName: 'DeviceStatusInfo' }).then(function (rs) {
      S.conductores = {};
      (rs || []).forEach(function (r) {
        var dv = r.device && r.device.id;
        var dr = r.driver && r.driver.id;
        if (dv && dr && dr !== 'UnknownDriverId' && typeof dr === 'string') S.conductores[dv] = dr;
      });
      var faltan = Object.keys(S.conductores).map(function (k) { return S.conductores[k]; })
        .filter(function (u, i, a) { return !(u in S.usuarios) && a.indexOf(u) === i; });
      if (!faltan.length) return;
      return multi(faltan.map(function (id) {
        return ['Get', { typeName: 'User', search: { id: id }, propertySelector: { fields: ['id', 'firstName', 'lastName', 'name'], isIncluded: true } }];
      })).then(function (res) {
        res.forEach(function (r, i) {
          var u = r && r[0];
          S.usuarios[faltan[i]] = u ? ((u.firstName || '') + ' ' + (u.lastName || '')).trim() || u.name : faltan[i];
        });
      });
    });
  }

  function leerIgnicion() {
    if (!S.cfg.resetIgnicion) { S.ignOff = {}; return Promise.resolve(); }
    var ids = Object.keys(S.vehiculos);
    var desde = new Date(desdeVentana()).toISOString();
    return multi(ids.map(function (id) {
      return ['Get', { typeName: 'StatusData', search: { deviceSearch: { id: id }, diagnosticSearch: { id: 'DiagnosticIgnitionId' }, fromDate: desde } }];
    })).then(function (res) {
      S.ignOff = {};
      res.forEach(function (r, i) {
        S.ignOff[ids[i]] = (r || []).filter(function (s) { return Number(s.data) === 0; })
          .map(function (s) { return new Date(s.dateTime).getTime(); });
      });
    });
  }

  function refrescar() {
    if (cargando || !api || (document.hidden && S.actualizado)) return Promise.resolve();
    cargando = true;
    return leerFeed()
      .then(function () {
        podarVentana();
        var tareas = [];
        if (Date.now() - S.condLeidos > 30000) tareas.push(leerConductores().then(function () { S.condLeidos = Date.now(); }));
        if (Date.now() - S.ignLeida > 30000) tareas.push(leerIgnicion().then(function () { S.ignLeida = Date.now(); }));
        if (Date.now() - S.usuariosLeidos > 300000) tareas.push(cargarUsuariosTarjeta());
        return Promise.all(tareas);
      })
      .then(function () { S.actualizado = Date.now(); S.error = null; })
      .catch(function (e) { S.error = errMsg(e); })
      .then(function () { cargando = false; renderVivo(); });
  }

  function estadoVehiculo(devId) {
    var v = S.vehiculos[devId];
    var asientos = v ? Number(v.d.asientos) || 0 : 0;
    var evs = (S.lecturas[devId] || []).slice();
    (S.ignOff[devId] || []).forEach(function (t) { evs.push({ t: t, tipo: 'ign_off' }); });
    var r = PidLogic.procesar(evs, S.cfg, asientos);
    r.asientos = asientos;
    r.desconocidas = r.abordo.filter(function (a) { return !persona(a.hex); }).length;
    r.ultima = evs.length ? Math.max.apply(null, evs.map(function (e) { return e.t; })) : null;
    return r;
  }

  function nombrePas(hex) {
    var p = persona(hex);
    return p ? p.nombre : null;
  }
  function nombreDev(id) {
    var d = S.devices[id];
    return d ? d.name : id;
  }
  function nombreConductor(devId) {
    var u = S.conductores[devId];
    return u ? (S.usuarios[u] || '…') : null;
  }

  /* ---------- render ---------- */
  function render() {
    var tabs = [['flota', 'Flota'], ['historial', 'Historial'], ['pasajeros', 'Pasajeros'], ['vehiculos', 'Vehículos'], ['config', 'Configuración']];
    var activa = S.vista === 'detalle' ? 'flota' : S.vista;
    root.innerHTML =
      '<div class="ap-head"><div class="ap-title">Mapa de asientos</div>' +
      '<nav class="ap-tabs">' + tabs.map(function (t) {
        return '<button data-tab="' + t[0] + '" class="' + (t[0] === activa ? 'on' : '') + '">' + t[1] + '</button>';
      }).join('') + '</nav>' +
      '<div class="ap-status" id="ap-status"></div></div>' +
      '<div id="ap-body"></div>';
    if (S.vista !== 'detalle') Mapa.cerrar();
    root.querySelectorAll('[data-tab]').forEach(function (b) {
      b.onclick = function () { S.vista = b.getAttribute('data-tab'); render(); };
    });
    if (S.vista === 'flota') renderFlota();
    else if (S.vista === 'detalle') renderDetalle();
    else if (S.vista === 'pasajeros') renderPasajeros();
    else if (S.vista === 'vehiculos') renderVehiculos();
    else if (S.vista === 'historial') renderHistorial();
    else renderConfig();
    renderStatus();
  }

  function renderStatus() {
    var el = $('#ap-status');
    if (!el) return;
    if (S.error) el.innerHTML = '<span class="err">Error: ' + esc(S.error) + '</span>';
    else el.textContent = S.actualizado ? 'Actualizado ' + new Date(S.actualizado).toLocaleTimeString('es-AR', { hourCycle: 'h23' }) : 'Cargando…';
  }

  /* Re-render liviano tras cada refresco: solo las vistas en vivo y sin pisar inputs con foco. */
  function renderVivo() {
    renderStatus();
    if (S.vista === 'flota') renderFlotaLista();
    else if (S.vista === 'detalle') renderDetalle();
  }

  /* --- Flota --- */
  function renderFlota() {
    $('#ap-body').innerHTML =
      '<div class="ap-bar"><input id="ap-qf" type="search" placeholder="Buscar patente o nombre" value="' + esc(S.filtroFlota) + '"></div>' +
      '<div id="ap-flota"></div>';
    var q = $('#ap-qf');
    q.oninput = function () { S.filtroFlota = q.value; renderFlotaLista(); };
    renderFlotaLista();
  }

  function renderFlotaLista() {
    var cont = $('#ap-flota');
    if (!cont) return;
    var q = S.filtroFlota.trim().toLowerCase().replace(/\s+/g, '');
    var ids = Object.keys(S.vehiculos).filter(function (id) {
      if (!q) return true;
      var d = S.devices[id] || {};
      return ((d.licensePlate || '') + ' ' + (d.name || '')).toLowerCase().replace(/\s+/g, '').indexOf(q) >= 0;
    }).sort(function (a, b) { return nombreDev(a).localeCompare(nombreDev(b)); });

    var sinConfig = Object.keys(S.lecturas).filter(function (id) {
      return !S.vehiculos[id] && S.lecturas[id].some(function (x) { return x.tipo === 'tarjeta'; });
    });

    var html = '';
    if (!Object.keys(S.vehiculos).length) {
      html += '<div class="ap-empty">Todavía no hay vehículos configurados. Definí el tipo y la cantidad de asientos en la pestaña <a href="#" data-go="vehiculos">Vehículos</a>.</div>';
    } else if (!ids.length) {
      html += '<div class="ap-empty">Ningún vehículo coincide con la búsqueda.</div>';
    }
    if (sinConfig.length) {
      html += '<div class="ap-warn">' + sinConfig.length + ' vehículo(s) con lecturas de tarjetas sin configurar: ' +
        sinConfig.map(function (id) { return esc(nombreDev(id)); }).join(', ') + '. <a href="#" data-go="vehiculos">Configurar</a></div>';
    }
    html += '<div class="ap-grid">' + ids.map(function (id) {
      var d = S.devices[id] || { name: id };
      var v = S.vehiculos[id].d;
      var e = estadoVehiculo(id);
      var n = e.abordo.length;
      var pct = e.asientos ? Math.min(100, Math.round(n / e.asientos * 100)) : 0;
      var cls = e.excedidos.length ? 'red' : (n ? 'green' : '');
      var cond = nombreConductor(id);
      return '<button class="ap-card" data-dev="' + esc(id) + '">' +
        '<div class="ap-card-top"><span class="ap-plate">' + esc(d.licensePlate || '—') + '</span><span class="ap-tag">' + esc(TIPOS[v.clase] || v.clase) + '</span></div>' +
        '<div class="ap-card-name">' + esc(d.name) + '</div>' +
        '<div class="ap-occ"><b>' + n + '</b> / ' + e.asientos + '</div>' +
        '<div class="ap-bar-bg"><div class="ap-bar-fg ' + cls + '" style="width:' + pct + '%"></div></div>' +
        '<div class="ap-card-foot">' + (cond ? 'Conductor: ' + esc(cond) : 'Sin conductor') +
        (e.desconocidas ? ' · <span class="amber">' + e.desconocidas + ' sin registrar</span>' : '') +
        (e.excedidos.length ? ' · <span class="red">' + e.excedidos.length + ' sin asiento</span>' : '') + '</div>' +
        '</button>';
    }).join('') + '</div>';
    cont.innerHTML = html;
    cont.querySelectorAll('[data-dev]').forEach(function (b) {
      b.onclick = function () { S.sel = b.getAttribute('data-dev'); S.vista = 'detalle'; render(); };
    });
    cont.querySelectorAll('[data-go]').forEach(function (a) {
      a.onclick = function (ev) { ev.preventDefault(); S.vista = a.getAttribute('data-go'); render(); };
    });
  }

  /* --- Detalle --- */
  function svgPlano(v, e) {
    var pl = PidLogic.plano(v.clase, e.asientos);
    var C = 36, G = 8, M = 16, TOP = 44;
    var W = M * 2 + pl.cols * C + (pl.cols - 1) * G;
    var H = TOP + pl.filas.length * (C + G) + M;
    var porAsiento = {};
    e.abordo.forEach(function (a) { if (a.asiento) porAsiento[a.asiento] = a; });
    var cond = nombreConductor(S.sel);
    var s = '<svg viewBox="0 0 ' + W + ' ' + H + '" width="' + Math.min(W * 1.5, 420) + '" class="ap-svg" role="img" aria-label="Plano de asientos">' +
      '<rect x="3" y="3" width="' + (W - 6) + '" height="' + (H - 6) + '" rx="26" class="ap-body"/>' +
      '<path d="M' + (M + 4) + ' 34 Q' + (W / 2) + ' 10 ' + (W - M - 4) + ' 34" class="ap-glass"/>';
    pl.filas.forEach(function (fila, fi) {
      fila.forEach(function (celda, ci) {
        if (celda == null) return;
        var x = M + ci * (C + G), y = TOP + fi * (C + G);
        var cls, txt, tip;
        if (celda === 'D') {
          cls = 'drv'; txt = cond ? iniciales(cond) : 'C'; tip = 'Conductor: ' + (cond || 'sin identificar');
        } else {
          var a = porAsiento[celda];
          if (!a) { cls = 'free'; txt = celda; tip = 'Asiento ' + celda + ' libre'; }
          else {
            var nom = nombrePas(a.hex);
            cls = nom ? 'occ' : 'unk'; txt = nom ? iniciales(nom) : '?';
            tip = 'Asiento ' + celda + ': ' + (nom || 'tarjeta sin registrar') + ' (' + a.hex + ') desde ' + hora(a.desde);
          }
        }
        s += '<g class="ap-seat ' + cls + '" data-seat="' + esc(celda) + '"><title>' + esc(tip) + '</title>' +
          '<rect x="' + x + '" y="' + y + '" width="' + C + '" height="' + C + '" rx="7"/>' +
          '<text x="' + (x + C / 2) + '" y="' + (y + C / 2 + 4) + '" text-anchor="middle">' + esc(txt) + '</text></g>';
      });
    });
    return s + '</svg>';
  }

  var ACCIONES = {
    sube: ['↗', 'subió', 'green'], baja: ['↙', 'bajó', ''], reascenso: ['↻', 'volvió a marcar (sigue a bordo)', ''],
    excedido: ['!', 'subió sin asiento libre', 'red'], reubicado: ['→', 'pasó a asiento', 'green'],
    bloqueada: ['·', 'lectura repetida ignorada', 'muted'], vaciado: ['↙', 'bajó (motor apagado)', '']
  };

  function renderDetalle() {
    var body = $('#ap-body');
    var id = S.sel;
    var v = S.vehiculos[id];
    if (!v) { S.vista = 'flota'; render(); return; }
    var d = S.devices[id] || { name: id };
    var e = estadoVehiculo(id);
    var n = e.abordo.length;
    var cond = nombreConductor(id);
    var pct = e.asientos ? Math.min(100, Math.round(n / e.asientos * 100)) : 0;

    // El esqueleto se arma una sola vez por vehículo para no destruir el mapa en cada refresco.
    if (body.getAttribute('data-detalle') !== id || !$('#ap-map')) {
      body.setAttribute('data-detalle', id);
      body.innerHTML =
        '<div class="ap-bar"><button id="ap-back">← Flota</button><div class="grow"><b>' + esc(d.licensePlate || '') + '</b> ' + esc(d.name) +
        ' <span class="ap-tag">' + esc(TIPOS[v.d.clase] || v.d.clase) + '</span></div></div>' +
        '<div class="ap-detail"><div class="ap-col">' +
        '<div class="ap-metrics" id="ap-d-metrics"></div>' +
        '<div class="ap-plano" id="ap-d-plano"></div>' +
        '<div class="ap-legend"><span><i class="sw occ"></i>Pasajero</span><span><i class="sw unk"></i>Sin registrar</span><span><i class="sw drv"></i>Conductor</span><span><i class="sw free"></i>Libre</span></div>' +
        '<div id="ap-d-warn"></div>' +
        '<h4>Recorrido</h4><div class="ap-map-wrap"><div id="ap-map"></div><div class="ap-map-msg" id="ap-map-msg">Cargando mapa…</div></div>' +
        '<div class="ap-legend"><span><i class="sw ruta"></i>Últimas ' + esc(S.cfg.ventanaHoras) + ' h</span><span><i class="dot up"></i>Subida</span><span><i class="dot down"></i>Bajada</span><span><i class="dot veh"></i>Posición actual</span></div>' +
        '</div><div class="ap-col">' +
        '<h4>A bordo</h4><div class="ap-list" id="ap-d-abordo"></div>' +
        '<h4>Últimos eventos</h4><div class="ap-list" id="ap-d-log"></div>' +
        '</div></div>';
      $('#ap-back').onclick = function () { S.vista = 'flota'; render(); };
      if (!body.getAttribute('data-lst')) body.setAttribute('data-lst', '1'), body.addEventListener('click', function (ev) {
        var b = ev.target.closest && ev.target.closest('[data-reg]');
        if (b && body.contains(b)) abrirFormPasajero(null, b.getAttribute('data-reg'));
        var f = ev.target.closest && ev.target.closest('[data-foco]');
        if (f && body.contains(f)) Mapa.enfocar(f.getAttribute('data-foco'));
      });
      Mapa.iniciar(id);
    }

    $('#ap-d-metrics').innerHTML =
      '<div class="ap-metric"><div class="lbl">A bordo</div><div class="val">' + n + ' / ' + e.asientos + '</div>' +
      '<div class="ap-bar-bg"><div class="ap-bar-fg ' + (e.excedidos.length ? 'red' : 'green') + '" style="width:' + pct + '%"></div></div></div>' +
      '<div class="ap-metric"><div class="lbl">Conductor (NFC)</div><div class="val sm">' + esc(cond || 'Sin identificar') + '</div></div>';
    $('#ap-d-plano').innerHTML = svgPlano(v.d, e);
    $('#ap-d-warn').innerHTML = e.excedidos.length ? '<div class="ap-warn red">' + e.excedidos.length + ' pasajero(s) a bordo sin asiento: se superó la capacidad.</div>' : '';

    $('#ap-d-abordo').innerHTML = e.abordo.map(function (a) {
      var nom = nombrePas(a.hex);
      return '<div class="ap-row"><span class="ap-num ' + (a.asiento ? (nom ? 'green' : 'amber') : 'red') + '">' + (a.asiento || '—') + '</span>' +
        '<span class="grow">' + (nom ? esc(nom) : '<span class="amber">Tarjeta sin registrar</span>') + ' <span class="muted mono">' + esc(a.hex) + '</span></span>' +
        '<span class="muted">' + hora(a.desde) + '</span>' +
        (nom ? '' : '<button class="ap-link" data-reg="' + esc(a.hex) + '">Registrar</button>') + '</div>';
    }).join('') || '<div class="ap-empty">Nadie a bordo.</div>';

    $('#ap-d-log').innerHTML = e.log.slice().reverse().filter(function (l) { return l.accion !== 'bloqueada'; }).slice(0, 40).map(function (l) {
      if (l.accion === 'texto' || l.accion === 'ignicion') {
        return '<div class="ap-row"><span class="ap-ico muted">i</span><span class="grow muted">' + esc(l.texto === 'READER CONNECTED' ? 'Lector conectado' : l.texto) + '</span><span class="muted">' + hora(l.t) + '</span></div>';
      }
      var a = ACCIONES[l.accion] || ['·', l.accion, ''];
      var nom = nombrePas(l.hex) || l.hex;
      var k = Mapa.clave(l);
      return '<div class="ap-row' + (Mapa.EV[l.accion] ? ' clic" data-foco="' + esc(k) + '" title="Ver en el mapa' : '') + '"><span class="ap-ico ' + a[2] + '">' + a[0] + '</span><span class="grow">' + esc(nom) + ' ' + a[1] +
        (l.asiento ? ' · asiento ' + l.asiento : '') + '</span><span class="muted">' + hora(l.t) + '</span></div>';
    }).join('') || '<div class="ap-empty">Sin lecturas en las últimas ' + esc(S.cfg.ventanaHoras) + ' h.</div>';

    Mapa.actualizar(id, e);
  }

  /* --- Minimapa (Leaflet + OpenStreetMap) --- */
  var Mapa = (function () {
    var LEAFLET_JS = 'https://cdn.jsdelivr.net/npm/leaflet@1.9.4/dist/leaflet.js';
    var LEAFLET_CSS = 'https://cdn.jsdelivr.net/npm/leaflet@1.9.4/dist/leaflet.css';
    var EV = { sube: 'up', excedido: 'up', baja: 'down', vaciado: 'down' };
    var M = null;         // estado del mapa del vehículo actual
    var carga = null;     // promesa de carga de Leaflet
    var posCache = {};    // clave de evento -> [lat, lon] | null

    function cargarLeaflet() {
      if (window.L) return Promise.resolve(window.L);
      if (carga) return carga;
      carga = new Promise(function (res, rej) {
        if (!document.querySelector('link[data-ap-leaflet]')) {
          var l = document.createElement('link');
          l.rel = 'stylesheet'; l.href = LEAFLET_CSS; l.setAttribute('data-ap-leaflet', '1');
          document.head.appendChild(l);
        }
        var s = document.createElement('script');
        s.src = LEAFLET_JS;
        s.onload = function () { window.L ? res(window.L) : rej(new Error('Leaflet no disponible')); };
        s.onerror = function () { carga = null; rej(new Error('No se pudo cargar el mapa.')); };
        document.head.appendChild(s);
      });
      return carga;
    }
    function clave(l) { return l.t + '|' + l.hex + '|' + l.accion; }
    function msg(t) { var el = $('#ap-map-msg'); if (el) { el.textContent = t || ''; el.style.display = t ? '' : 'none'; } }

    function iniciar(dev) {
      if (M && M.map) { try { M.map.remove(); } catch (e) { /* ya removido */ } }
      M = { dev: dev, map: null, ruta: [], ultimoT: null, linea: null, veh: null, capa: null, marcas: {}, encuadrado: false, ocupado: false };
      var yo = M;
      cargarLeaflet().then(function (L) {
        var el = $('#ap-map');
        if (M !== yo || !el) return;
        yo.map = L.map(el, { zoomControl: true, attributionControl: true }).setView([-38.95, -68.06], 11);
        L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png', {
          maxZoom: 19, attribution: '&copy; OpenStreetMap'
        }).addTo(yo.map);
        yo.capa = L.layerGroup().addTo(yo.map);
        msg('');
        if (S.vista === 'detalle' && S.sel === dev) actualizar(dev, estadoVehiculo(dev));
      }).catch(function (e) { msg(errMsg(e)); });
    }

    function inicioRuta(e) {
      var ini = Date.now() - (Number(S.cfg.ventanaHoras) || 12) * 3600000;
      return ini;
    }

    function puntoColor(tipo) { return tipo === 'up' ? '#4aa75e' : '#4e677e'; }

    function actualizar(dev, e) {
      if (!M || M.dev !== dev || !M.map || M.ocupado) return;
      var yo = M, L = window.L;
      yo.ocupado = true;
      var desde = yo.ultimoT ? new Date(yo.ultimoT + 1) : new Date(inicioRuta(e));
      var hasta = new Date();
      var eventos = e.log.filter(function (l) { return EV[l.accion]; });
      var faltan = eventos.filter(function (l) { return !(clave(l) in posCache); });

      var calls = [['Get', { typeName: 'LogRecord', search: { deviceSearch: { id: dev }, fromDate: desde.toISOString(), toDate: hasta.toISOString() } }]];
      faltan.forEach(function (l) {
        var t = new Date(l.t).toISOString();
        calls.push(['Get', { typeName: 'LogRecord', search: { deviceSearch: { id: dev }, fromDate: t, toDate: t } }]);
      });
      multi(calls).then(function (res) {
        if (M !== yo) return;
        (res[0] || []).forEach(function (lr) {
          var t = new Date(lr.dateTime).getTime();
          if ((!lr.latitude && !lr.longitude) || (yo.ultimoT && t <= yo.ultimoT)) return;
          yo.ruta.push([lr.latitude, lr.longitude]);
          yo.ultimoT = t;
        });
        faltan.forEach(function (l, i) {
          var lr = (res[i + 1] || []).filter(function (x) { return x.latitude || x.longitude; })[0];
          posCache[clave(l)] = lr ? [lr.latitude, lr.longitude] : null;
        });
        dibujar(yo, eventos, L);
      }).catch(function (err) { msg('No se pudo leer el recorrido: ' + errMsg(err)); })
        .then(function () { yo.ocupado = false; });
    }

    function dibujar(yo, eventos, L) {
      var ruta = yo.ruta;
      if (ruta.length > 4000) { // aligera el trazo sin perder el último punto
        var paso = Math.ceil(ruta.length / 4000);
        ruta = ruta.filter(function (p, i) { return i % paso === 0 || i === yo.ruta.length - 1; });
      }
      if (!yo.linea) yo.linea = L.polyline(ruta, { color: '#0078d3', weight: 4, opacity: 0.8 }).addTo(yo.map);
      else yo.linea.setLatLngs(ruta);

      var ult = yo.ruta[yo.ruta.length - 1];
      if (ult) {
        if (!yo.veh) yo.veh = L.circleMarker(ult, { radius: 8, color: '#fff', weight: 3, fillColor: '#25477b', fillOpacity: 1 }).addTo(yo.map).bindTooltip('Posición actual');
        else yo.veh.setLatLng(ult);
        yo.veh.bringToFront();
      }

      var vivos = {};
      eventos.forEach(function (l) {
        var k = clave(l), pos = posCache[k];
        if (!pos) return;
        vivos[k] = 1;
        if (yo.marcas[k]) return;
        var tipo = EV[l.accion];
        var nom = nombrePas(l.hex) || 'Tarjeta ' + l.hex;
        var html = '<b>' + esc(nom) + '</b><br>' + (tipo === 'up' ? 'Subió' : 'Bajó') + (l.asiento ? ' · asiento ' + l.asiento : '') +
          '<br>' + fechaHora(l.t) + '<br><span style="color:#4e677e">' + pos[0].toFixed(5) + ', ' + pos[1].toFixed(5) + '</span>';
        yo.marcas[k] = L.circleMarker(pos, { radius: 7, color: '#fff', weight: 2, fillColor: puntoColor(tipo), fillOpacity: 1 })
          .bindPopup(html).bindTooltip(esc(nom) + ' · ' + (tipo === 'up' ? 'subió' : 'bajó') + ' ' + hora(l.t))
          .addTo(yo.capa);
      });
      Object.keys(yo.marcas).forEach(function (k) { if (!vivos[k]) { yo.capa.removeLayer(yo.marcas[k]); delete yo.marcas[k]; } });

      if (!yo.encuadrado) {
        var pts = ruta.slice();
        Object.keys(yo.marcas).forEach(function (k) { pts.push(yo.marcas[k].getLatLng()); });
        if (pts.length > 1) { yo.map.fitBounds(L.latLngBounds(pts), { padding: [24, 24], maxZoom: 16 }); yo.encuadrado = true; }
        else if (pts.length === 1) { yo.map.setView(pts[0], 15); yo.encuadrado = true; }
        else msg('Sin posiciones GPS en las últimas ' + S.cfg.ventanaHoras + ' h.');
      }
      if (yo.ruta.length || Object.keys(yo.marcas).length) msg('');
    }

    function enfocar(k) {
      if (!M || !M.map || !M.marcas[k]) return;
      var el = $('#ap-map');
      if (el && el.scrollIntoView) el.scrollIntoView({ block: 'nearest', behavior: 'smooth' });
      M.map.setView(M.marcas[k].getLatLng(), Math.max(M.map.getZoom(), 16));
      M.marcas[k].openPopup();
    }

    function cerrar() {
      if (M && M.map) { try { M.map.remove(); } catch (e) { /* nada */ } }
      M = null;
    }

    return { EV: EV, iniciar: iniciar, actualizar: actualizar, enfocar: enfocar, clave: clave, cerrar: cerrar };
  })();

  /* --- Pasajeros --- */
  function tarjetasSinRegistrar() {
    var vistas = {};
    Object.keys(S.lecturas).forEach(function (dev) {
      S.lecturas[dev].forEach(function (l) {
        if (l.tipo !== 'tarjeta' || persona(l.hex)) return;
        if (!vistas[l.hex] || vistas[l.hex].t < l.t) vistas[l.hex] = { t: l.t, dev: dev };
      });
    });
    return Object.keys(vistas).map(function (h) { return { hex: h, t: vistas[h].t, dev: vistas[h].dev }; })
      .sort(function (a, b) { return b.t - a.t; });
  }

  function renderPasajeros() {
    var q = S.filtroPas.trim().toLowerCase();
    var filas = Object.keys(S.usuariosTarjeta).map(function (h) {
      var u = S.usuariosTarjeta[h];
      return { hex: h, nombre: u.nombre, legajo: u.legajo, empresa: '', turno: '', activo: true, origen: 'geotab', userId: u.userId, usuario: u.usuario };
    }).concat(Object.keys(S.pasajeros).map(function (h) {
      var d = S.pasajeros[h].d;
      return { hex: h, nombre: d.nombre, legajo: d.legajo, empresa: d.empresa, turno: d.turno, activo: d.activo !== false, origen: 'padron', pisado: !!S.usuariosTarjeta[h] };
    }));
    var nGeotab = Object.keys(S.usuariosTarjeta).length, nPadron = Object.keys(S.pasajeros).length;
    var lista = filas.filter(function (f) {
      if (!q) return true;
      return [f.hex, f.nombre, f.legajo, f.empresa, f.turno, f.usuario].join(' ').toLowerCase().indexOf(q) >= 0;
    }).sort(function (a, b) { return String(a.nombre).localeCompare(String(b.nombre)); });
    var sinReg = tarjetasSinRegistrar();

    $('#ap-body').innerHTML =
      (sinReg.length ? '<div class="ap-warn"><b>Tarjetas leídas sin registrar</b><div class="ap-chips">' + sinReg.map(function (s) {
        return '<button class="ap-chip" data-reg="' + esc(s.hex) + '">' + esc(s.hex) + ' <span class="muted">' + esc(nombreDev(s.dev)) + ' ' + hora(s.t) + '</span></button>';
      }).join('') + '</div></div>' : '') +
      '<div class="ap-bar"><input id="ap-qp" type="search" class="grow" placeholder="Buscar nombre, legajo, hex…" value="' + esc(S.filtroPas) + '">' +
      '<button id="ap-new" class="primary">+ Pasajero</button><button id="ap-imp">Importar Excel</button><input id="ap-file" type="file" accept=".xlsx,.xls,.csv" hidden></div>' +
      '<div class="ap-muted-line">' + (nGeotab + nPadron) + ' tarjetas identificadas: ' + nGeotab + ' por usuarios de Geotab con llave, ' + nPadron + ' en el padrón del add-in</div>' +
      '<table class="ap-table"><thead><tr><th>Hex tarjeta</th><th>Nombre</th><th>Legajo / DNI</th><th>Empresa</th><th>Turno</th><th>Origen</th><th></th></tr></thead><tbody>' +
      (lista.map(function (f) {
        var origen = f.origen === 'geotab'
          ? '<span class="ap-tag geo">Usuario Geotab</span>'
          : '<span class="ap-tag">Padrón</span>' + (f.pisado ? ' <span class="muted" title="Hay un usuario de Geotab con esta llave; se usa el usuario.">(reemplazado)</span>' : '') + (f.activo ? '' : ' <span class="muted">inactivo</span>');
        var acciones = f.origen === 'geotab'
          ? '<a class="ap-link" href="#user,id:' + encodeURIComponent(f.userId) + '" target="_top">Ver usuario</a>'
          : '<button data-edit="' + esc(f.hex) + '">Editar</button><button class="danger" data-del="' + esc(f.hex) + '">Borrar</button>';
        return '<tr><td class="mono">' + esc(f.hex) + '</td><td>' + esc(f.nombre) + (f.usuario && f.usuario !== f.nombre ? ' <span class="muted">' + esc(f.usuario) + '</span>' : '') + '</td><td>' + esc(f.legajo) + '</td><td>' + esc(f.empresa) + '</td><td>' + esc(f.turno) + '</td>' +
          '<td>' + origen + '</td><td class="ap-actions">' + acciones + '</td></tr>';
      }).join('') || '<tr><td colspan="7" class="ap-empty">Sin pasajeros.</td></tr>') +
      '</tbody></table>';

    var qp = $('#ap-qp');
    qp.oninput = function () {
      S.filtroPas = qp.value; renderPasajeros();
      var nq = $('#ap-qp'); nq.focus(); nq.setSelectionRange(nq.value.length, nq.value.length);
    };
    $('#ap-new').onclick = function () { abrirFormPasajero(null, ''); };
    $('#ap-imp').onclick = function () { $('#ap-file').click(); };
    $('#ap-file').onchange = function (ev) { if (ev.target.files[0]) importarExcel(ev.target.files[0]); };
    root.querySelectorAll('[data-reg]').forEach(function (b) { b.onclick = function () { abrirFormPasajero(null, b.getAttribute('data-reg')); }; });
    root.querySelectorAll('[data-edit]').forEach(function (b) { b.onclick = function () { abrirFormPasajero(S.pasajeros[b.getAttribute('data-edit')]); }; });
    root.querySelectorAll('[data-del]').forEach(function (b) {
      b.onclick = function () {
        var p = S.pasajeros[b.getAttribute('data-del')];
        if (!p) return;
        confirmar('¿Borrar a ' + p.d.nombre + ' (' + p.d.hex + ')?').then(function (ok) {
          if (!ok) return;
          return borrarRec(p).then(function () { delete S.pasajeros[p.d.hex]; toast('Pasajero borrado'); renderPasajeros(); });
        }).catch(function (e) { toast(errMsg(e), 'err'); });
      };
    });
  }

  function confirmar(msg) {
    return new Promise(function (res) {
      var m = document.createElement('div');
      m.className = 'ap-modal';
      m.innerHTML = '<div class="ap-dialog"><p>' + esc(msg) + '</p><div class="ap-dlg-btns"><button type="button" data-no>Cancelar</button><button type="button" class="primary" data-si>Aceptar</button></div></div>';
      root.appendChild(m);
      m.querySelector('[data-no]').onclick = function () { m.remove(); res(false); };
      m.querySelector('[data-si]').onclick = function () { m.remove(); res(true); };
    });
  }

  function modal(titulo, html, onOk) {
    var m = document.createElement('div');
    m.className = 'ap-modal';
    m.innerHTML = '<div class="ap-dialog"><h3>' + esc(titulo) + '</h3><form>' + html +
      '<div class="ap-err" id="ap-merr"></div><div class="ap-dlg-btns"><button type="button" data-x>Cancelar</button><button type="submit" class="primary">Guardar</button></div></form></div>';
    root.appendChild(m);
    var form = m.querySelector('form');
    m.querySelector('[data-x]').onclick = function () { m.remove(); };
    form.onsubmit = function (ev) {
      ev.preventDefault();
      var err = m.querySelector('#ap-merr');
      err.textContent = '';
      var btn = form.querySelector('[type=submit]');
      btn.disabled = true;
      Promise.resolve().then(function () { return onOk(form); })
        .then(function () { m.remove(); })
        .catch(function (e) { err.textContent = errMsg(e); btn.disabled = false; });
    };
    var first = form.querySelector('input');
    if (first) first.focus();
  }

  function abrirFormPasajero(rec, hexInicial) {
    var d = rec ? rec.d : { hex: hexInicial || '', nombre: '', legajo: '', empresa: '', turno: '', activo: true };
    modal(rec ? 'Editar pasajero' : 'Nuevo pasajero',
      '<label>Hex de la tarjeta<input name="hex" required value="' + esc(d.hex) + '" placeholder="8C9A8C13"></label>' +
      '<label>Nombre y apellido<input name="nombre" required value="' + esc(d.nombre) + '"></label>' +
      '<label>Legajo / DNI<input name="legajo" value="' + esc(d.legajo) + '"></label>' +
      '<label>Empresa / departamento<input name="empresa" value="' + esc(d.empresa) + '"></label>' +
      '<label>Turno<input name="turno" value="' + esc(d.turno) + '"></label>' +
      '<label class="chk"><input type="checkbox" name="activo"' + (d.activo === false ? '' : ' checked') + '> Activo</label>',
      function (f) {
        var hex = PidLogic.normHex(f.elements.hex.value);
        var nombre = f.elements.nombre.value.trim();
        if (hex.length < 8) throw new Error('El hex debe tener al menos 8 caracteres (4 bytes).');
        if (!nombre) throw new Error('Ingresá el nombre.');
        if (S.usuariosTarjeta[hex]) throw new Error('Esa tarjeta ya es la llave del usuario de Geotab ' + S.usuariosTarjeta[hex].usuario + '.');
        var existente = S.pasajeros[hex];
        if (existente && (!rec || existente.id !== rec.id)) throw new Error('Esa tarjeta ya está asignada a ' + existente.d.nombre + '.');
        var nd = { tipo: 'pasajero', hex: hex, nombre: nombre, legajo: f.elements.legajo.value.trim(), empresa: f.elements.empresa.value.trim(), turno: f.elements.turno.value.trim(), activo: f.elements.activo.checked };
        return guardarRec(rec, nd).then(function (r) {
          if (rec && rec.d.hex !== hex) delete S.pasajeros[rec.d.hex];
          S.pasajeros[hex] = r;
          toast('Pasajero guardado');
          render();
        });
      });
  }

  function cargarXlsx() {
    if (window.XLSX) return Promise.resolve(window.XLSX);
    return new Promise(function (res, rej) {
      var s = document.createElement('script');
      s.src = XLSX_URL;
      s.onload = function () { res(window.XLSX); };
      s.onerror = function () { rej(new Error('No se pudo cargar el lector de Excel.')); };
      document.head.appendChild(s);
    });
  }

  function mapearColumnas(headers) {
    var reglas = {
      hex: /hex|uid|tarjeta|rfid/,
      nombre: /nombre|apellido|pasajero/,
      legajo: /legajo|dni|documento|folio|id empleado|n[uú]mero/,
      empresa: /empresa|depart|[aá]rea|contratista/,
      turno: /turno|horario/,
      activo: /activo|estado/
    };
    var m = {};
    headers.forEach(function (h, i) {
      var k = String(h || '').toLowerCase().trim();
      Object.keys(reglas).forEach(function (campo) {
        if (m[campo] == null && reglas[campo].test(k)) {
          if (campo === 'legajo' && /hex/.test(k)) return;
          m[campo] = i;
        }
      });
    });
    return m;
  }

  function importarExcel(file) {
    cargarXlsx().then(function (XLSX) {
      return file.arrayBuffer().then(function (buf) {
        var wb = XLSX.read(buf, { type: 'array' });
        var filas = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, raw: false, defval: '' });
        if (filas.length < 2) throw new Error('El archivo no tiene filas.');
        var col = mapearColumnas(filas[0]);
        if (col.hex == null || col.nombre == null) throw new Error('No encontré las columnas de hex de tarjeta y nombre en la primera fila.');
        var nuevos = [], vistos = {}, invalidos = 0;
        filas.slice(1).forEach(function (f) {
          var hex = PidLogic.normHex(f[col.hex]);
          var nombre = String(f[col.nombre] || '').trim();
          if (hex.length < 8 || !nombre) { if (hex || nombre) invalidos++; return; }
          if (vistos[hex]) return;
          vistos[hex] = 1;
          var act = col.activo == null ? true : !/^(no|0|false|inactivo|baja)$/i.test(String(f[col.activo]).trim());
          nuevos.push({ tipo: 'pasajero', hex: hex, nombre: nombre, legajo: col.legajo == null ? '' : String(f[col.legajo]).trim(),
            empresa: col.empresa == null ? '' : String(f[col.empresa]).trim(), turno: col.turno == null ? '' : String(f[col.turno]).trim(), activo: act });
        });
        if (!nuevos.length) throw new Error('No hay filas válidas para importar.');
        return confirmar('¿Importar ' + nuevos.length + ' pasajeros' + (invalidos ? ' (' + invalidos + ' filas inválidas se omiten)' : '') + '? Las tarjetas que ya existen se actualizan.').then(function (ok) {
        if (!ok) return;
        var calls = nuevos.map(function (d) {
          var ex = S.pasajeros[d.hex];
          var entity = { addInId: ADDIN_ID, groups: [{ id: 'GroupCompanyId' }], details: d };
          if (ex) { entity.id = ex.id; return ['Set', { typeName: 'AddInData', entity: entity }]; }
          return ['Add', { typeName: 'AddInData', entity: entity }];
        });
        toast('Importando ' + nuevos.length + ' pasajeros…');
        return multi(calls).then(function () { return cargarAddInData(); }).then(function () {
          toast(nuevos.length + ' pasajeros importados');
          render();
        });
        });
      });
    }).catch(function (e) { toast(errMsg(e), 'err'); });
  }

  /* --- Vehículos --- */
  function renderVehiculos() {
    var q = S.filtroVeh.trim().toLowerCase().replace(/\s+/g, '');
    var conLect = {};
    Object.keys(S.lecturas).forEach(function (id) { if (S.lecturas[id].some(function (x) { return x.tipo === 'tarjeta'; })) conLect[id] = 1; });
    var ids = Object.keys(S.devices).filter(function (id) {
      var d = S.devices[id];
      if (!q) return S.vehiculos[id] || conLect[id];
      return ((d.licensePlate || '') + ' ' + (d.name || '') + ' ' + (d.serialNumber || '')).toLowerCase().replace(/\s+/g, '').indexOf(q) >= 0;
    }).sort(function (a, b) {
      return (S.vehiculos[b] ? 1 : 0) - (S.vehiculos[a] ? 1 : 0) || nombreDev(a).localeCompare(nombreDev(b));
    }).slice(0, 200);

    $('#ap-body').innerHTML =
      '<div class="ap-bar"><input id="ap-qv" type="search" class="grow" placeholder="Buscar patente, nombre o serie para agregar un vehículo" value="' + esc(S.filtroVeh) + '"></div>' +
      '<div class="ap-muted-line">' + (q ? 'Resultados de la búsqueda' : 'Vehículos configurados y con lecturas del PID') + '</div>' +
      '<table class="ap-table"><thead><tr><th>Vehículo</th><th>Patente</th><th>Tipo</th><th>Asientos de pasajeros</th><th></th></tr></thead><tbody>' +
      (ids.map(function (id) {
        var d = S.devices[id];
        var v = S.vehiculos[id];
        var clase = v ? v.d.clase : 'combi';
        var asientos = v ? v.d.asientos : ASIENTOS_DEF[clase];
        return '<tr data-row="' + esc(id) + '"><td>' + esc(d.name) + (conLect[id] ? ' <span class="ap-tag">PID</span>' : '') + '</td><td>' + esc(d.licensePlate || '') + '</td>' +
          '<td><select data-k="clase">' + Object.keys(TIPOS).map(function (t) { return '<option value="' + t + '"' + (t === clase ? ' selected' : '') + '>' + TIPOS[t] + '</option>'; }).join('') + '</select></td>' +
          '<td><input data-k="asientos" type="number" min="1" max="80" step="1" value="' + esc(asientos) + '" style="width:80px"></td>' +
          '<td class="ap-actions"><button data-save class="' + (v ? '' : 'primary') + '">' + (v ? 'Guardar' : 'Agregar') + '</button>' + (v ? '<button class="danger" data-rm>Quitar</button>' : '') + '</td></tr>';
      }).join('') || '<tr><td colspan="5" class="ap-empty">' + (q ? 'Sin resultados.' : 'Buscá un vehículo por patente o nombre para configurarlo.') + '</td></tr>') +
      '</tbody></table>';

    var qv = $('#ap-qv');
    qv.oninput = function () {
      S.filtroVeh = qv.value; renderVehiculos();
      var nq = $('#ap-qv'); nq.focus(); nq.setSelectionRange(nq.value.length, nq.value.length);
    };
    root.querySelectorAll('tr[data-row]').forEach(function (tr) {
      var id = tr.getAttribute('data-row');
      var sel = tr.querySelector('[data-k=clase]');
      var inp = tr.querySelector('[data-k=asientos]');
      sel.onchange = function () { if (!S.vehiculos[id]) inp.value = ASIENTOS_DEF[sel.value]; };
      tr.querySelector('[data-save]').onclick = function () {
        var n = Math.round(Number(inp.value));
        if (!(n >= 1 && n <= 80)) { toast('La cantidad de asientos debe estar entre 1 y 80', 'err'); return; }
        var d = { tipo: 'vehiculo', deviceId: id, clase: sel.value, asientos: n };
        guardarRec(S.vehiculos[id], d).then(function (r) { S.vehiculos[id] = r; toast('Vehículo guardado'); renderVehiculos(); })
          .catch(function (e) { toast(errMsg(e), 'err'); });
      };
      var rm = tr.querySelector('[data-rm]');
      if (rm) rm.onclick = function () {
        confirmar('¿Quitar ' + nombreDev(id) + ' del mapa de asientos?').then(function (ok) {
          if (!ok) return;
          return borrarRec(S.vehiculos[id]).then(function () { delete S.vehiculos[id]; toast('Vehículo quitado'); renderVehiculos(); });
        }).catch(function (e) { toast(errMsg(e), 'err'); });
      };
    });
  }

  /* --- Historial --- */
  var EVENTOS_HIST = { sube: 'Subió', baja: 'Bajó', vaciado: 'Bajó (motor apagado)', excedido: 'Subió sin asiento', reubicado: 'Pasó a asiento', reascenso: 'Volvió a marcar' };
  var MAX_FILAS_HIST = 3000;

  function fechaInput(d) {
    var z = function (n) { return (n < 10 ? '0' : '') + n; };
    return d.getFullYear() + '-' + z(d.getMonth() + 1) + '-' + z(d.getDate());
  }
  function fechaHora(t) {
    var d = new Date(t);
    return d.toLocaleDateString('es-AR', { day: '2-digit', month: '2-digit', year: 'numeric' }) + ' ' +
      d.toLocaleTimeString('es-AR', { hour: '2-digit', minute: '2-digit', second: '2-digit', hourCycle: 'h23' });
  }

  function renderHistorial() {
    var h = S.hist;
    if (!h.desde) { h.desde = fechaInput(new Date()); h.hasta = h.desde; }
    var vehs = Object.keys(S.vehiculos).sort(function (a, b) { return nombreDev(a).localeCompare(nombreDev(b)); });
    $('#ap-body').innerHTML =
      '<form class="ap-bar ap-hist-f" id="ap-hf">' +
      '<label>Desde<input type="date" name="desde" value="' + esc(h.desde) + '" required></label>' +
      '<label>Hasta<input type="date" name="hasta" value="' + esc(h.hasta) + '" required></label>' +
      '<label>Vehículo<select name="dev"><option value="">Todos</option>' + vehs.map(function (id) {
        var d = S.devices[id] || {};
        return '<option value="' + esc(id) + '"' + (h.dev === id ? ' selected' : '') + '>' + esc((d.licensePlate ? d.licensePlate + ' · ' : '') + nombreDev(id)) + '</option>';
      }).join('') + '</select></label>' +
      '<label>Evento<select name="evento"><option value="">Subidas y bajadas</option><option value="sube"' + (h.evento === 'sube' ? ' selected' : '') + '>Solo subidas</option><option value="baja"' + (h.evento === 'baja' ? ' selected' : '') + '>Solo bajadas</option></select></label>' +
      '<label class="grow">Pasajero<input type="search" name="q" placeholder="Nombre, legajo o hex" value="' + esc(h.q) + '"></label>' +
      '<button type="submit" class="primary"' + (h.cargando ? ' disabled' : '') + '>' + (h.cargando ? 'Buscando…' : 'Buscar') + '</button>' +
      '<button type="button" id="ap-csv"' + (h.filas && h.filas.length ? '' : ' disabled') + '>Exportar a Excel</button>' +
      '</form>' +
      (h.aviso ? '<div class="ap-warn">' + esc(h.aviso) + '</div>' : '') +
      '<div id="ap-hres"></div>';
    var f = $('#ap-hf');
    f.onsubmit = function (ev) {
      ev.preventDefault();
      h.desde = f.elements.desde.value; h.hasta = f.elements.hasta.value; h.dev = f.elements.dev.value;
      h.evento = f.elements.evento.value; h.q = f.elements.q.value;
      buscarHistorial();
    };
    f.elements.q.oninput = function () { h.q = f.elements.q.value; renderHistTabla(); };
    f.elements.evento.onchange = function () { h.evento = f.elements.evento.value; renderHistTabla(); };
    $('#ap-csv').onclick = exportarHistorial;
    renderHistTabla();
  }

  function filasHistFiltradas() {
    var h = S.hist;
    var q = h.q.trim().toLowerCase();
    return (h.filas || []).filter(function (r) {
      if (h.evento === 'sube' && ['sube', 'excedido'].indexOf(r.accion) < 0) return false;
      if (h.evento === 'baja' && ['baja', 'vaciado'].indexOf(r.accion) < 0) return false;
      if (!q) return true;
      var p = persona(r.hex) || {};
      return [r.hex, p.nombre, p.legajo, p.usuario].join(' ').toLowerCase().indexOf(q) >= 0;
    });
  }

  function renderHistTabla() {
    var cont = $('#ap-hres');
    if (!cont) return;
    var h = S.hist;
    if (h.filas == null) { cont.innerHTML = '<div class="ap-empty">Elegí el rango de fechas y tocá Buscar.</div>'; return; }
    var filas = filasHistFiltradas();
    var mostrar = filas.slice(0, 500);
    cont.innerHTML = '<div class="ap-muted-line">' + filas.length + ' eventos' + (filas.length > mostrar.length ? ' (se muestran los primeros 500; exportá a Excel para ver todos)' : '') + '</div>' +
      '<table class="ap-table"><thead><tr><th>Fecha y hora</th><th>Evento</th><th>Pasajero</th><th>Hex tarjeta</th><th>Vehículo</th><th>Asiento</th><th>Ubicación</th></tr></thead><tbody>' +
      (mostrar.map(function (r) {
        var p = persona(r.hex);
        var d = S.devices[r.dev] || {};
        var sube = r.accion === 'sube' || r.accion === 'excedido';
        var ubic = r.lat != null
          ? '<a href="https://www.google.com/maps?q=' + r.lat + ',' + r.lon + '" target="_blank" rel="noopener">' + esc(r.direccion || (r.lat.toFixed(5) + ', ' + r.lon.toFixed(5))) + '</a>'
          : '<span class="muted">Sin posición</span>';
        return '<tr><td>' + fechaHora(r.t) + '</td>' +
          '<td><span class="ap-ev ' + (sube ? 'up' : 'down') + '">' + (sube ? '↗ ' : '↙ ') + esc(EVENTOS_HIST[r.accion] || r.accion) + '</span></td>' +
          '<td>' + (p ? esc(p.nombre) : '<span class="amber">Sin registrar</span>') + '</td>' +
          '<td class="mono">' + esc(r.hex) + '</td>' +
          '<td>' + esc(d.licensePlate || '') + ' <span class="muted">' + esc(nombreDev(r.dev)) + '</span></td>' +
          '<td>' + (r.asiento || '—') + '</td><td>' + ubic + '</td></tr>';
      }).join('') || '<tr><td colspan="7" class="ap-empty">Sin subidas ni bajadas en el período.</td></tr>') +
      '</tbody></table>';
  }

  function buscarHistorial() {
    var h = S.hist;
    var d0 = new Date(h.desde + 'T00:00:00'), d1 = new Date(h.hasta + 'T23:59:59.999');
    h.aviso = '';
    if (!(d1 >= d0)) { h.aviso = 'La fecha "hasta" tiene que ser igual o posterior a "desde".'; renderHistorial(); return; }
    if (d1 - d0 > 31 * 86400000) { h.aviso = 'El rango máximo es de 31 días.'; renderHistorial(); return; }
    var ids = h.dev ? [h.dev] : Object.keys(S.vehiculos);
    if (!ids.length) { h.aviso = 'No hay vehículos configurados.'; renderHistorial(); return; }
    // Se leen 12 h antes del rango para reconstruir quién ya iba a bordo al empezar.
    var look = new Date(d0.getTime() - 12 * 3600000).toISOString(), fin = d1.toISOString();
    h.cargando = true; renderHistorial();

    var calls = ids.map(function (id) {
      return ['Get', { typeName: 'CustomData', search: { deviceSearch: { id: id }, fromDate: look, toDate: fin } }];
    });
    if (S.cfg.resetIgnicion) ids.forEach(function (id) {
      calls.push(['Get', { typeName: 'StatusData', search: { deviceSearch: { id: id }, diagnosticSearch: { id: 'DiagnosticIgnitionId' }, fromDate: look, toDate: fin } }]);
    });
    multi(calls).then(function (res) {
      var filas = [];
      ids.forEach(function (id, i) {
        var evs = (res[i] || []).map(function (cd) {
          var dec = PidLogic.decodeCustomData(cd.data);
          return { t: new Date(cd.dateTime).getTime(), tipo: dec.tipo === 'tarjeta' ? 'tarjeta' : 'texto', hex: dec.hex, texto: dec.texto };
        });
        if (S.cfg.resetIgnicion) {
          (res[ids.length + i] || []).forEach(function (sd) { if (Number(sd.data) === 0) evs.push({ t: new Date(sd.dateTime).getTime(), tipo: 'ign_off' }); });
        }
        var asientos = S.vehiculos[id] ? Number(S.vehiculos[id].d.asientos) || 0 : 99;
        PidLogic.procesar(evs, S.cfg, asientos).log.forEach(function (l) {
          if (!EVENTOS_HIST[l.accion] || l.t < d0.getTime() || l.t > d1.getTime()) return;
          filas.push({ t: l.t, accion: l.accion, hex: l.hex, asiento: l.asiento, dev: id });
        });
      });
      filas.sort(function (a, b) { return b.t - a.t; });
      if (filas.length > MAX_FILAS_HIST) {
        h.aviso = 'Hay ' + filas.length + ' eventos; se ubican los ' + MAX_FILAS_HIST + ' más recientes. Acotá el rango para ver todos.';
        filas = filas.slice(0, MAX_FILAS_HIST);
      }
      return ubicarFilas(filas).then(function () { return filas; });
    }).then(function (filas) {
      h.filas = filas;
    }).catch(function (e) {
      h.aviso = 'No se pudo leer el historial: ' + errMsg(e);
    }).then(function () {
      h.cargando = false;
      if (S.vista === 'historial') renderHistorial();
    });
  }

  /* Posición GPS de cada evento: LogRecord con fromDate = toDate devuelve la posición
     interpolada por Geotab en ese instante. Luego la dirección con GetAddresses. */
  function ubicarFilas(filas) {
    if (!filas.length) return Promise.resolve();
    return multi(filas.map(function (r) {
      var t = new Date(r.t).toISOString();
      return ['Get', { typeName: 'LogRecord', search: { deviceSearch: { id: r.dev }, fromDate: t, toDate: t } }];
    })).then(function (res) {
      res.forEach(function (lrs, i) {
        var lr = (lrs || []).filter(function (x) { return x.latitude || x.longitude; })[0];
        if (lr) { filas[i].lat = lr.latitude; filas[i].lon = lr.longitude; }
      });
      var claves = {}, coords = [];
      filas.forEach(function (r) {
        if (r.lat == null) return;
        var k = r.lat.toFixed(5) + ',' + r.lon.toFixed(5);
        if (!(k in claves)) { claves[k] = coords.length; coords.push({ x: r.lon, y: r.lat }); }
        r.kc = claves[k];
      });
      if (!coords.length) return;
      var lotes = [];
      for (var i = 0; i < coords.length; i += 100) lotes.push(['GetAddresses', { coordinates: coords.slice(i, i + 100), movingAddresses: false }]);
      return multi(lotes).then(function (rs) {
        var dirs = [].concat.apply([], rs);
        filas.forEach(function (r) {
          var a = r.kc != null && dirs[r.kc];
          if (!a) return;
          var zonas = (a.zones || []).map(function (z) { return z.name; }).filter(Boolean);
          r.direccion = zonas.length ? zonas.join(', ') + ' · ' + (a.formattedAddress || '') : a.formattedAddress;
        });
      }).catch(function () { /* sin direcciones: quedan las coordenadas */ });
    });
  }

  function exportarHistorial() {
    var filas = filasHistFiltradas().slice().reverse();
    var cel = function (v) { v = v == null ? '' : String(v); return /[;"\n]/.test(v) ? '"' + v.replace(/"/g, '""') + '"' : v; };
    var lineas = [['Fecha', 'Hora', 'Evento', 'Pasajero', 'Legajo', 'Hex tarjeta', 'Origen', 'Patente', 'Vehículo', 'Asiento', 'Latitud', 'Longitud', 'Ubicación'].join(';')];
    filas.forEach(function (r) {
      var p = persona(r.hex) || {};
      var d = new Date(r.t), dev = S.devices[r.dev] || {};
      lineas.push([
        d.toLocaleDateString('es-AR', { day: '2-digit', month: '2-digit', year: 'numeric' }),
        d.toLocaleTimeString('es-AR', { hour: '2-digit', minute: '2-digit', second: '2-digit', hourCycle: 'h23' }),
        EVENTOS_HIST[r.accion] || r.accion, p.nombre || 'Sin registrar', p.legajo || '', r.hex,
        p.origen === 'geotab' ? 'Usuario Geotab' : (p.origen ? 'Padrón' : ''),
        dev.licensePlate || '', nombreDev(r.dev), r.asiento || '',
        r.lat != null ? String(r.lat).replace('.', ',') : '', r.lon != null ? String(r.lon).replace('.', ',') : '', r.direccion || ''
      ].map(cel).join(';'));
    });
    var blob = new Blob(['﻿' + lineas.join('\r\n')], { type: 'text/csv;charset=utf-8' });
    var a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = 'historial_pasajeros_' + S.hist.desde + '_' + S.hist.hasta + '.csv';
    root.appendChild(a); a.click(); a.remove();
    setTimeout(function () { URL.revokeObjectURL(a.href); }, 2000);
  }

  /* --- Configuración --- */
  function renderConfig() {
    var c = S.cfg;
    $('#ap-body').innerHTML =
      '<form class="ap-form" id="ap-cfg">' +
      '<label>Descenso (minutos)<input name="descensoMin" type="number" min="0" step="1" value="' + esc(c.descensoMin) + '">' +
      '<small>Si la misma tarjeta vuelve a leerse antes de este tiempo desde que subió, cuenta como bajada. Usá el mismo valor que en el PID.</small></label>' +
      '<label>Desbloqueo (segundos)<input name="desbloqueoSeg" type="number" min="0" step="1" value="' + esc(c.desbloqueoSeg) + '">' +
      '<small>Las relecturas de la misma tarjeta dentro de este lapso se ignoran.</small></label>' +
      '<label>Regla de bajada<select name="modo">' +
      '<option value="pid"' + (c.modo === 'pid' ? ' selected' : '') + '>Igual al PID</option>' +
      '<option value="alternar"' + (c.modo === 'alternar' ? ' selected' : '') + '>Alternar</option></select>' +
      '<small>Igual al PID: pasado el tiempo de descenso, una relectura cuenta como nuevo ascenso. Alternar: toda relectura de una tarjeta a bordo es bajada.</small></label>' +
      '<label class="chk"><input type="checkbox" name="resetIgnicion"' + (c.resetIgnicion ? ' checked' : '') + '> Vaciar el vehículo al apagar el motor</label>' +
      '<label>Ventana de lecturas (horas)<input name="ventanaHoras" type="number" min="1" max="48" step="1" value="' + esc(c.ventanaHoras) + '">' +
      '<small>Cuántas horas hacia atrás se reconstruye quién está a bordo.</small></label>' +
      '<label>Refresco (segundos)<input name="refrescoSeg" type="number" min="5" max="300" step="1" value="' + esc(c.refrescoSeg) + '">' +
      '<small>Cada cuánto se consultan las lecturas nuevas. Mínimo 5 s.</small></label>' +
      '<div class="ap-err" id="ap-cerr"></div><button type="submit" class="primary">Guardar configuración</button></form>';
    $('#ap-cfg').onsubmit = function (ev) {
      ev.preventDefault();
      var f = ev.target;
      var nv = {
        descensoMin: Math.max(0, Math.round(Number(f.elements.descensoMin.value) || 0)),
        desbloqueoSeg: Math.max(0, Math.round(Number(f.elements.desbloqueoSeg.value) || 0)),
        modo: f.elements.modo.value,
        resetIgnicion: f.elements.resetIgnicion.checked,
        ventanaHoras: Math.min(48, Math.max(1, Math.round(Number(f.elements.ventanaHoras.value) || 12))),
        refrescoSeg: Math.min(300, Math.max(5, Math.round(Number(f.elements.refrescoSeg.value) || 5)))
      };
      var cambiaVentana = nv.ventanaHoras !== S.cfg.ventanaHoras;
      guardarRec(S.cfgRec, { tipo: 'config', valores: nv }).then(function (r) {
        S.cfgRec = { id: r.id }; S.cfg = nv;
        if (cambiaVentana) { S.feedVersion = null; S.lecturas = {}; }
        iniciarTimer();
        toast('Configuración guardada');
        return refrescar();
      }).catch(function (e) { $('#ap-cerr').textContent = errMsg(e); });
    };
  }

  function iniciarTimer() {
    if (timer) clearInterval(timer);
    timer = setInterval(refrescar, Math.max(5, Number(S.cfg.refrescoSeg) || 5) * 1000);
  }

  return {
    initialize: function (freshApi, state, callback) {
      api = freshApi;
      root = document.getElementById('asientosPid');
      callback();
    },
    focus: function (freshApi) {
      api = freshApi;
      root = document.getElementById('asientosPid');
      S.feedVersion = null; S.lecturas = {}; S.error = null;
      render();
      Promise.all([cargarAddInData(), cargarDevices(), cargarUsuariosTarjeta()])
        .then(function () { render(); return refrescar(); })
        .catch(function (e) { S.error = errMsg(e); render(); })
        .then(iniciarTimer);
    },
    blur: function () {
      if (timer) clearInterval(timer);
      timer = null;
    }
  };
};
