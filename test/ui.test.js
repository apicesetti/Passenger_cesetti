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
  { id: 'b8', device: { id: 'bBB' }, dateTime: iso(5), data: 'UlVJRAAAAAECAwQ=' }
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
  if (method === 'Get' && p.typeName === 'User') return [{ id: 'u1', firstName: 'Juan', lastName: 'Pérez' }];
  if (method === 'Get' && p.typeName === 'StatusData') return [];
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
  assert(txt().includes('2 vehículo(s) con lecturas'), 'aviso sin configurar');

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
  assert.strictEqual($$('.ap-seat.unk').length, 2, '2 tarjetas sin registrar');
  assert.strictEqual($('.ap-seat[data-seat="1"] title').textContent.includes('8C9A8C13'), true);
  assert(txt().includes('Lector conectado'));

  // Registrar desde el detalle
  $('[data-reg="8C9A8C13"]').click(); await sleep(10);
  const f = $('.ap-dialog form');
  assert.strictEqual(f.elements.hex.value, '8C9A8C13');
  f.elements.nombre.value = 'Ana Gómez'; f.elements.legajo.value = '1234';
  f.dispatchEvent(new w.Event('submit', { cancelable: true })); await sleep(50);
  assert(!$('.ap-modal'), 'modal cerrado');
  assert.strictEqual($$('.ap-seat.occ').length, 1);
  assert.strictEqual($('.ap-seat[data-seat="1"] text').textContent, 'AG');

  // Duplicado rechazado
  $('[data-tab=pasajeros]').click(); await sleep(10);
  $('#ap-new').click(); await sleep(5);
  const f2 = $('.ap-dialog form'); f2.elements.hex.value = '8c 9a 8c 13'; f2.elements.nombre.value = 'Otro';
  f2.dispatchEvent(new w.Event('submit', { cancelable: true })); await sleep(30);
  assert($('#ap-merr').textContent.includes('ya está asignada'), 'duplicado');
  $('[data-x]').click();
  assert(txt().includes('31012C0F') && txt().includes('01020304'), 'chips sin registrar');
  assert(txt().includes('1 pasajeros registrados'));

  // Borrar con confirmación propia
  $('[data-del="8C9A8C13"]').click(); await sleep(5);
  $('[data-si]').click(); await sleep(30);
  assert(txt().includes('0 pasajeros registrados'));

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
