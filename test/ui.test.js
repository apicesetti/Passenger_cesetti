const { JSDOM } = require('jsdom');
const fs = require('fs');
const assert = require('assert');
const path = require('path');
const base = path.dirname(path.resolve(process.argv[2]));
// Arma el HTML con los archivos referenciados en línea (igual que los carga MyGeotab)
const html = fs.readFileSync(process.argv[2], 'utf8')
  .replace(/<link rel="stylesheet" href="([^"]+)">/g, (m, f) => '<style>' + fs.readFileSync(path.join(base, f), 'utf8') + '</style>')
  .replace(/<script src="([^"]+)"><\/script>/g, (m, f) => '<script>' + fs.readFileSync(path.join(base, f), 'utf8') + '</script>');

const now = Date.now();
const iso = (minAgo) => new Date(now - minAgo * 60000).toISOString();
const store = [];
let nextId = 1;
const calls = [];
const feed = [
  { id: 'b5', device: { id: 'bBA' }, dateTime: iso(60), data: 'UkVBREVSIENPTk5FQ1RFRA==' },
  { id: 'b6', device: { id: 'bBA' }, dateTime: iso(10), data: 'UlVJRAAAAIyajBM=' },
  { id: 'b7', device: { id: 'bBA' }, dateTime: iso(9.9), data: 'UlVJRAAAADEBLA8=' },
  { id: 'b8', device: { id: 'bBB' }, dateTime: iso(5), data: 'UlVJRAAAAAECAwQ=' },
  { id: 'b9', device: { id: 'bBA' }, dateTime: iso(4), data: Buffer.from([82, 85, 73, 68, 4, 0xA1, 0xB2, 0xC3, 0xD4, 0xE5, 0xF6]).toString('base64') }
];
function handle(method, p) {
  calls.push([method, p.typeName]);
  if (method === 'Get' && p.typeName === 'AddInData') return JSON.parse(JSON.stringify(store));
  if (method === 'Add') { const id = 'a' + nextId++; store.push(Object.assign({ id }, p.entity)); return id; }
  if (method === 'Set') { const i = store.findIndex((s) => s.id === p.entity.id); store[i] = Object.assign({}, p.entity); return null; }
  if (method === 'Remove') { const i = store.findIndex((s) => s.id === p.entity.id); store.splice(i, 1); return null; }
  if (method === 'Get' && p.typeName === 'Device') return [{ id: 'bBA', name: 'CESETTI_TEST', licensePlate: 'AF123CD', serialNumber: 'G9Z8ZBFDMYZ6' }, { id: 'bBB', name: 'COLECTIVO 7', licensePlate: 'AE456GH' }];
  if (method === 'GetFeed') return p.fromVersion ? { data: [], toVersion: 'v2' } : { data: feed, toVersion: 'v1' };
  if (method === 'Get' && p.typeName === 'DeviceStatusInfo') return [{ device: { id: 'bBA' }, driver: { id: 'u1' } }, { device: { id: 'bBB' }, driver: 'UnknownDriverId' }];
  if (method === 'Get' && p.typeName === 'User' && p.search.isDriver) return [
    { id: 'b1C8', name: 'PASAJERO_1', firstName: 'PASAJERO', lastName: '1', keys: [{ driverKeyType: 'Custom', serialNumber: '8C9A8C13' }] },
    { id: 'u1', name: 'jperez', firstName: 'Juan', lastName: 'Pérez', keys: [{ driverKeyType: 'Nfc', serialNumber: '04A1B2C3D4E5F6' }] }];
  if (method === 'Get' && p.typeName === 'User') return [{ id: 'u1', firstName: 'Juan', lastName: 'Pérez' }];
  if (method === 'Get' && p.typeName === 'CustomData') return feed.filter((f) => f.device.id === p.search.deviceSearch.id && f.dateTime >= p.search.fromDate && f.dateTime <= p.search.toDate);
  if (method === 'Get' && p.typeName === 'LogRecord') return [{ latitude: -38.9504623, longitude: -68.0575638, dateTime: p.search.fromDate }];
  if (method === 'GetAddresses') return p.coordinates.map(() => ({ formattedAddress: 'Buenos Aires 444, Neuquén', zones: [{ name: 'Base Cesetti' }] }));
  if (method === 'Get' && p.typeName === 'StatusData') return [];
  if (method === 'Get' && p.typeName === 'IoxAddOn') return [{ channel: 1, type: 4247, device: { id: 'bBA' }, dateTime: '2026-11-12T16:55:40.918Z' }, { channel: 6, type: 4291, device: { id: 'bBB' }, dateTime: '2026-11-12T16:55:37.664Z' }];
  throw new Error('no mock ' + method + ' ' + p.typeName);
}
const api = {
  call(m, p, ok, err) { setTimeout(() => { try { ok(handle(m, p)); } catch (e) { err(e); } }, 1); },
  multiCall(cs, ok, err) { setTimeout(() => { try { ok(cs.map(([m, p]) => handle(m, p))); } catch (e) { err(e); } }, 1); }
};

const dom = new JSDOM(html, {
  runScripts: 'dangerously',
  beforeParse(w) { w.geotab = { addin: {} }; w.onerror = (m) => console.error('window error', m); }
});
const w = dom.window;
const d = w.document;
const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const $ = (s) => d.querySelector(s);
const $$ = (s) => [...d.querySelectorAll(s)];
const txt = () => $('#asientosPid').textContent;

(async () => {
  const addin = w.geotab.addin.asientosPid();
  await new Promise((r) => addin.initialize(api, {}, r));
  addin.focus(api, {});
  await sleep(150);
  assert(txt().includes('Todavía no hay vehículos configurados'), 'flota vacía');
  assert(txt().includes('2 vehículo(s) con lector PID o lecturas'), 'aviso sin configurar');

  // Configurar vehículo: pestaña Vehículos muestra los que tienen lecturas
  $('[data-tab=vehiculos]').click(); await sleep(20);
  const row = $('tr[data-row=bBA]');
  assert(row, 'fila bBA');
  row.querySelector('[data-k=clase]').value = 'combi';
  row.querySelector('[data-k=asientos]').value = '15';
  row.querySelector('[data-save]').click(); await sleep(50);
  assert(store.some((s) => s.details.tipo === 'vehiculo' && s.details.deviceId === 'bBA'), 'vehículo guardado');
  // buscador de vehículos
  const qv = $('#ap-qv'); qv.value = 'ae456'; qv.dispatchEvent(new w.Event('input')); await sleep(10);
  assert($('tr[data-row=bBB]') && !$('tr[data-row=bBA]'), 'búsqueda por patente en vehículos');
  $('tr[data-row=bBB] [data-k=clase]').value = 'colectivo';
  $('tr[data-row=bBB] [data-k=asientos]').value = '45';
  $('tr[data-row=bBB] [data-save]').click(); await sleep(50);

  // Flota con buscador
  $('[data-tab=flota]').click(); await sleep(20);
  assert.strictEqual($$('.ap-card').length, 2);
  assert(txt().includes('Conductor: Juan Pérez'));
  const qf = $('#ap-qf'); qf.value = 'AF 123'; qf.dispatchEvent(new w.Event('input')); await sleep(10);
  assert.strictEqual($$('.ap-card').length, 1, 'filtro por patente con espacios');
  assert($('.ap-card').textContent.includes('2 / 15'));

  // Detalle
  $('.ap-card').click(); await sleep(20);
  assert.strictEqual($$('.ap-seat').length, 16, '15 asientos + conductor');
  assert.strictEqual($$('.ap-seat.unk').length, 1, '1 tarjeta sin registrar (la otra es PASAJERO_1)');
  assert.strictEqual($('.ap-seat[data-seat="1"] text').textContent, 'P1', 'PASAJERO 1 por llave de Geotab');
  assert.strictEqual($('.ap-seat[data-seat="1"] title').textContent.includes('8C9A8C13'), true);
  assert(txt().includes('Lector conectado'));

  // Registrar desde el detalle
  $('[data-reg="31012C0F"]').click(); await sleep(10);
  const f = $('.ap-dialog form');
  assert.strictEqual(f.elements.hex.value, '31012C0F');
  f.elements.nombre.value = 'Ana Gómez'; f.elements.legajo.value = '1234';
  f.dispatchEvent(new w.Event('submit', { cancelable: true })); await sleep(50);
  assert(!$('.ap-modal'), 'modal cerrado');
  assert.strictEqual($$('.ap-seat.occ').length, 2);
  assert.strictEqual($('.ap-seat[data-seat="2"] text').textContent, 'AG');

  // Duplicado rechazado
  $('[data-tab=pasajeros]').click(); await sleep(10);
  $('#ap-new').click(); await sleep(5);
  const f2 = $('.ap-dialog form'); f2.elements.hex.value = '31 01 2c 0f'; f2.elements.nombre.value = 'Otro';
  f2.dispatchEvent(new w.Event('submit', { cancelable: true })); await sleep(30);
  assert($('#ap-merr').textContent.includes('ya está asignada'), 'duplicado padrón');
  f2.elements.hex.value = '8C9A8C13';
  f2.dispatchEvent(new w.Event('submit', { cancelable: true })); await sleep(30);
  assert($('#ap-merr').textContent.includes('PASAJERO_1'), 'duplicado usuario Geotab');
  $('[data-x]').click();
  assert(txt().includes('01020304'), 'chip sin registrar');
  assert(txt().includes('2 usuarios de MyGeotab y 1 en el padrón'), 'conteo personas');
  const rolP1 = $('select[data-rol="b1C8"]'), rolJP = $('select[data-rol="u1"]');
  assert.strictEqual(rolP1.value, 'pasajero', 'llave Custom => pasajero');
  assert.strictEqual(rolJP.value, 'conductor', 'llave NFC => conductor');

  // Inhabilitar a PASAJERO_1 => asiento rojo y alerta
  $('[data-hab="u:b1C8"]').click(); await sleep(5);
  const fh = $('.ap-dialog form'); fh.elements.hab.checked = false; fh.elements.motivo.value = 'Inducción vencida';
  fh.dispatchEvent(new w.Event('submit', { cancelable: true })); await sleep(40);
  assert(store.some((x) => x.details.tipo === 'persona' && x.details.userId === 'b1C8' && x.details.habilitado === false), 'persona guardada');
  assert($('#ap-alertas').textContent.includes('no está habilitado'), 'alerta no habilitado');
  $('[data-tab=flota]').click(); await sleep(10);
  $('.ap-card[data-dev=bBA]').click(); await sleep(20);
  assert.strictEqual($$('.ap-seat.nohab').length, 1, 'asiento rojo');
  assert(txt().includes('Lector PID conectado desde'), 'lector detectado por IoxAddOn');
  assert(txt().includes('llave de conductor'), 'llave NFC del conductor ignorada');
  assert.strictEqual(d.querySelectorAll('.ap-seat.occ, .ap-seat.unk, .ap-seat.nohab').length, 2, 'la llave del conductor no ocupa asiento');

  // Capacidad: bBA con 1 asiento => el segundo pasajero queda sin asiento y alerta
  $('[data-tab=vehiculos]').click(); await sleep(10);
  { const qv2 = $('#ap-qv'); qv2.value = ''; qv2.dispatchEvent(new w.Event('input')); await sleep(10); }
  $('tr[data-row=bBA] [data-k=asientos]').value = '1';
  $('tr[data-row=bBA] [data-save]').click(); await sleep(40);
  assert($('#ap-alertas').textContent.includes('alerta'), 'barra de alertas');
  if (!$('#ap-alertas').textContent.includes('Capacidad superada')) { $('#ap-alert-tg').click(); await sleep(5); }
  assert($('#ap-alertas').textContent.includes('Capacidad superada'), 'alerta capacidad');
  $('tr[data-row=bBA] [data-k=asientos]').value = '15';
  $('tr[data-row=bBA] [data-save]').click(); await sleep(40);
  $('[data-tab=pasajeros]').click(); await sleep(10);

  // Historial
  $('[data-tab=historial]').click(); await sleep(10);
  $('#ap-hf').dispatchEvent(new w.Event('submit', { cancelable: true })); await sleep(120);
  const filasH = $$('#ap-hres tbody tr');
  assert.strictEqual(filasH.length, 3, 'historial: 3 subidas (2 en bBA, 1 en bBB)');
  assert(txt().includes('Base Cesetti · Buenos Aires 444, Neuquén'), 'dirección con zona');
  assert(txt().includes('No habilitado: Inhabilitado: Inducción vencida'), 'alerta en historial');
  assert(txt().includes('PASAJERO 1') && txt().includes('Ana Gómez'));
  const hq = $('#ap-hf').elements.q; hq.value = 'pasajero'; hq.dispatchEvent(new w.Event('input')); await sleep(5);
  assert.strictEqual($$('#ap-hres tbody tr').length, 1, 'filtro por pasajero');
  $('[data-tab=pasajeros]').click(); await sleep(10);

  // Borrar con confirmación propia
  $('[data-del="31012C0F"]').click(); await sleep(5);
  $('[data-si]').click(); await sleep(30);
  assert(txt().includes('0 en el padrón'));

  // Configuración
  $('[data-tab=config]').click(); await sleep(5);
  const cf = $('#ap-cfg'); cf.elements.modo.value = 'alternar'; cf.elements.descensoMin.value = '20';
  cf.dispatchEvent(new w.Event('submit', { cancelable: true })); await sleep(60);
  const c = store.find((s) => s.details.tipo === 'config');
  assert.strictEqual(c.details.valores.modo, 'alternar');
  assert.strictEqual(c.details.valores.descensoMin, 20);

  // Re-focus: recarga todo desde AddInData
  addin.blur(); addin.focus(api, {}); await sleep(150); $('[data-tab=flota]').click(); await sleep(10);
  assert.strictEqual($$('.ap-card').length, 1, 'mantiene el filtro');
  addin.blur();
  console.log('ui: OK', calls.length, 'llamadas a la API');
  process.exit(0);
})().catch((e) => { console.error('FALLO', e); process.exit(1); });
