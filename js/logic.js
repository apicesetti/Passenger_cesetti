/* Lógica pura del add-in (sin DOM ni API). Se reutiliza en los tests de node. */
var PidLogic = (function () {
  'use strict';

  var HEX = '0123456789ABCDEF';

  function bytesToHex(bytes) {
    var s = '';
    for (var i = 0; i < bytes.length; i++) s += HEX[(bytes[i] >> 4) & 15] + HEX[bytes[i] & 15];
    return s;
  }

  function b64ToBytes(b64) {
    var bin = typeof atob === 'function' ? atob(b64) : Buffer.from(b64, 'base64').toString('binary');
    var out = new Array(bin.length);
    for (var i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i) & 255;
    return out;
  }

  /* Normaliza un hex de tarjeta: mayúsculas, solo hex, sin bytes 00 de relleno
     a la izquierda (conserva al menos 4 bytes / 8 caracteres). */
  function normHex(v) {
    var s = String(v == null ? '' : v).toUpperCase().replace(/[^0-9A-F]/g, '');
    if (s.length % 2) s = '0' + s;
    while (s.length > 8 && s.slice(0, 2) === '00') s = s.slice(2);
    return s;
  }

  /* Decodifica el campo data (base64) de un CustomData del PID.
     'RUID' + bytes -> {tipo:'tarjeta', hex} | texto ASCII -> {tipo:'texto', texto} | otro -> {tipo:'otro', hex} */
  function decodeCustomData(b64) {
    var b;
    try { b = b64ToBytes(b64 || ''); } catch (e) { return { tipo: 'otro', hex: '' }; }
    if (b.length > 4 && b[0] === 82 && b[1] === 85 && b[2] === 73 && b[3] === 68) {
      return { tipo: 'tarjeta', hex: normHex(bytesToHex(b.slice(4))) };
    }
    var printable = b.length > 0 && b.every(function (c) { return c >= 32 && c < 127; });
    if (printable) return { tipo: 'texto', texto: String.fromCharCode.apply(null, b) };
    return { tipo: 'otro', hex: bytesToHex(b) };
  }

  /* Procesa los eventos de UN vehículo y devuelve quién está a bordo.
     eventos: [{t: ms, tipo:'tarjeta'|'texto'|'ign_off', hex?, texto?}]
     cfg: {descensoMin, desbloqueoSeg, modo:'pid'|'alternar', resetIgnicion}
     asientos: cantidad de asientos de pasajeros.
     Reglas:
       - Relectura de la misma tarjeta antes de "desbloqueo" -> se ignora.
       - Tarjeta a bordo releída antes de "descenso" desde su ascenso -> baja.
       - modo 'alternar': toda relectura válida de una tarjeta a bordo -> baja.
       - modo 'pid': relectura después de "descenso" -> nuevo ascenso (conserva asiento).
       - Al subir toma el primer asiento libre desde el 1; sin lugar -> excedido.
       - Al liberarse un asiento, lo toma el primer excedido en espera.
       - resetIgnicion: al apagar el motor se bajan todos. */
  function procesar(eventos, cfg, asientos) {
    var descMs = (Number(cfg.descensoMin) || 0) * 60000;
    var desbMs = (Number(cfg.desbloqueoSeg) || 0) * 1000;
    var alternar = cfg.modo === 'alternar';
    var evs = eventos.slice().sort(function (a, b) { return a.t - b.t; });
    var tarjetas = {};
    var ocupados = {};
    var excedidos = [];
    var log = [];

    function primerLibre() {
      for (var i = 1; i <= asientos; i++) if (!ocupados[i]) return i;
      return null;
    }
    function liberar(asiento, t) {
      delete ocupados[asiento];
      if (excedidos.length) {
        var hex = excedidos.shift();
        tarjetas[hex].asiento = asiento;
        ocupados[asiento] = hex;
        log.push({ t: t, hex: hex, accion: 'reubicado', asiento: asiento });
      }
    }
    function bajar(hex, t, accion) {
      var st = tarjetas[hex];
      var asiento = st.asiento;
      var ix = excedidos.indexOf(hex);
      if (ix >= 0) excedidos.splice(ix, 1);
      st.abordo = false; st.asiento = null;
      log.push({ t: t, hex: hex, accion: accion, asiento: asiento });
      if (asiento) liberar(asiento, t);
    }

    evs.forEach(function (ev) {
      if (ev.tipo === 'ign_off') {
        if (!cfg.resetIgnicion) return;
        var abordo = Object.keys(tarjetas).filter(function (h) { return tarjetas[h].abordo; });
        excedidos = [];
        abordo.forEach(function (h) {
          var st = tarjetas[h];
          log.push({ t: ev.t, hex: h, accion: 'vaciado', asiento: st.asiento });
          if (st.asiento) delete ocupados[st.asiento];
          st.abordo = false; st.asiento = null;
        });
        if (abordo.length) log.push({ t: ev.t, accion: 'ignicion', texto: 'Motor apagado: vehículo vaciado' });
        return;
      }
      if (ev.tipo !== 'tarjeta') {
        log.push({ t: ev.t, accion: 'texto', texto: ev.texto || '' });
        return;
      }
      var hex = ev.hex;
      var st = tarjetas[hex] || (tarjetas[hex] = { abordo: false, ultimo: null, ascenso: null, asiento: null });
      if (st.ultimo != null && ev.t - st.ultimo < desbMs) {
        log.push({ t: ev.t, hex: hex, accion: 'bloqueada', asiento: st.asiento });
        return;
      }
      st.ultimo = ev.t;
      if (st.abordo && (alternar || ev.t - st.ascenso < descMs)) {
        bajar(hex, ev.t, 'baja');
        return;
      }
      if (st.abordo) {
        st.ascenso = ev.t;
        log.push({ t: ev.t, hex: hex, accion: 'reascenso', asiento: st.asiento });
        return;
      }
      st.abordo = true; st.ascenso = ev.t;
      var libre = primerLibre();
      if (libre) { st.asiento = libre; ocupados[libre] = hex; }
      else excedidos.push(hex);
      log.push({ t: ev.t, hex: hex, accion: libre ? 'sube' : 'excedido', asiento: libre });
    });

    var abordo = Object.keys(tarjetas)
      .filter(function (h) { return tarjetas[h].abordo; })
      .map(function (h) { return { hex: h, asiento: tarjetas[h].asiento, desde: tarjetas[h].ascenso }; })
      .sort(function (a, b) { return (a.asiento || 9999) - (b.asiento || 9999) || a.desde - b.desde; });

    return { abordo: abordo, ocupados: ocupados, excedidos: excedidos.slice(), log: log };
  }

  /* Plano de asientos. Devuelve {cols, filas:[[celda]]}; celda = 'D' (conductor),
     número de asiento, o null (pasillo/vacío). Numeración de adelante hacia atrás,
     de izquierda a derecha. Conducción a la izquierda (Argentina). */
  var PATRONES = {
    camioneta: { ancho: 3, slots: [0, 1, 2], frente: [2] },
    combi: { ancho: 4, slots: [0, 1, 3], frente: [2, 3] },
    colectivo: { ancho: 5, slots: [0, 1, 3, 4], frente: [] }
  };

  function plano(tipo, n) {
    var p = PATRONES[tipo] || PATRONES.combi;
    var filas = [];
    var num = 1;
    var frente = new Array(p.ancho).fill(null);
    frente[0] = 'D';
    p.frente.forEach(function (c) { if (num <= n) frente[c] = num++; });
    filas.push(frente);
    while (num <= n) {
      var resto = n - num + 1;
      var fila = new Array(p.ancho).fill(null);
      if (resto > p.slots.length && resto <= p.ancho) {
        for (var c = 0; c < resto; c++) fila[c] = num++;
      } else {
        p.slots.forEach(function (c) { if (num <= n) fila[c] = num++; });
      }
      filas.push(fila);
    }
    return { cols: p.ancho, filas: filas };
  }

  return {
    normHex: normHex,
    bytesToHex: bytesToHex,
    decodeCustomData: decodeCustomData,
    procesar: procesar,
    plano: plano
  };
})();
if (typeof module !== 'undefined') module.exports = PidLogic;
