/* node test/logic.test.js */
const assert = require('assert');
const L = require('../js/logic.js');

const T = (s) => new Date('2026-11-12T' + s + 'Z').getTime();
const cfg = { descensoMin: 30, desbloqueoSeg: 60, modo: 'pid', resetIgnicion: false };

// Tramas reales de CESETTI_TEST (12/11/2026)
const crudo = [
  ['UkVBREVSIENPTk5FQ1RFRA==', '16:55:41'],
  ['UlVJRAAAAIyajBM=', '18:07:04'],
  ['UlVJRAAAADEBLA8=', '18:07:07'],
  ['UlVJRAAAADEBLA8=', '18:14:04'],
  ['UlVJRAAAAIyajBM=', '18:14:07']
];
const dec = crudo.map(([b64, h]) => Object.assign({ t: T(h) }, L.decodeCustomData(b64)));
assert.deepStrictEqual(dec[0].tipo, 'texto');
assert.strictEqual(dec[0].texto, 'READER CONNECTED');
assert.strictEqual(dec[1].hex, '8C9A8C13');
assert.strictEqual(dec[2].hex, '31012C0F');

// Hasta 18:07:07: los dos a bordo en asientos 1 y 2
let r = L.procesar(dec.slice(0, 3), cfg, 15);
assert.deepStrictEqual(r.abordo.map((a) => [a.hex, a.asiento]), [['8C9A8C13', 1], ['31012C0F', 2]]);
// Con las relecturas de 18:14 (7 min < 30): bajan los dos
r = L.procesar(dec, cfg, 15);
assert.strictEqual(r.abordo.length, 0);
assert.deepStrictEqual(r.log.filter((l) => l.accion === 'baja').map((l) => l.hex), ['31012C0F', '8C9A8C13']);

// Desbloqueo: relectura a los 20 s se ignora
r = L.procesar([{ t: T('10:00:00'), tipo: 'tarjeta', hex: 'AA' + 'BBCCDD' }, { t: T('10:00:20'), tipo: 'tarjeta', hex: 'AABBCCDD' }], cfg, 4);
assert.strictEqual(r.abordo.length, 1);
assert.strictEqual(r.log[1].accion, 'bloqueada');

// Modo pid: relectura después de 30 min = reascenso (sigue a bordo)
r = L.procesar([{ t: T('10:00:00'), tipo: 'tarjeta', hex: 'AABBCCDD' }, { t: T('10:45:00'), tipo: 'tarjeta', hex: 'AABBCCDD' }], cfg, 4);
assert.strictEqual(r.abordo.length, 1);
assert.strictEqual(r.log[1].accion, 'reascenso');
// Modo alternar: la misma secuencia es bajada
r = L.procesar([{ t: T('10:00:00'), tipo: 'tarjeta', hex: 'AABBCCDD' }, { t: T('10:45:00'), tipo: 'tarjeta', hex: 'AABBCCDD' }], Object.assign({}, cfg, { modo: 'alternar' }), 4);
assert.strictEqual(r.abordo.length, 0);

// Primer asiento libre + excedido + reubicación
const ev = (h, hex) => ({ t: T(h), tipo: 'tarjeta', hex });
r = L.procesar([ev('08:00:00', 'A1A1A1A1'), ev('08:01:00', 'B2B2B2B2'), ev('08:02:00', 'C3C3C3C3'), ev('08:03:00', 'A1A1A1A1'), ev('08:04:00', 'D4D4D4D4')], cfg, 2);
assert.deepStrictEqual(r.abordo.map((a) => [a.hex, a.asiento]), [['C3C3C3C3', 1], ['B2B2B2B2', 2], ['D4D4D4D4', null]]);
assert.deepStrictEqual(r.excedidos, ['D4D4D4D4']);

// Vaciado por ignición
r = L.procesar([ev('08:00:00', 'A1A1A1A1'), { t: T('09:00:00'), tipo: 'ign_off' }, ev('09:10:00', 'B2B2B2B2')], Object.assign({}, cfg, { resetIgnicion: true }), 4);
assert.deepStrictEqual(r.abordo.map((a) => [a.hex, a.asiento]), [['B2B2B2B2', 1]]);

// Planos: la cantidad de asientos numerados coincide y empiezan en 1
for (const [tipo, n] of [['camioneta', 4], ['camioneta', 3], ['combi', 15], ['combi', 19], ['colectivo', 45], ['colectivo', 40], ['colectivo', 12]]) {
  const p = L.plano(tipo, n);
  const nums = p.filas.flat().filter((c) => typeof c === 'number');
  assert.deepStrictEqual(nums, Array.from({ length: n }, (_, i) => i + 1), tipo + ' ' + n);
  assert.strictEqual(p.filas[0][0], 'D');
}

assert.strictEqual(L.normHex('00 00 00 8c 9a 8c 13'), '8C9A8C13');
assert.strictEqual(L.normHex('04A1B2C3D4E5F6'), '04A1B2C3D4E5F6');
console.log('logic: OK');
