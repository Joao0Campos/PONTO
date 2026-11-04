'use strict';
// f5: dia aberto só com a ENTRADA mostra o saldo "se sair agora" (almoço automático). Hora LOCAL.
const test = require('node:test');
const assert = require('node:assert/strict');
const C = require('../js/core');

const L = (k, hhmm) => new Date(k + 'T' + hhmm + ':00').toISOString();
const S = dias => {
  const meses = {};
  for (const k of Object.keys(dias)) (meses[k.slice(0, 7)] = meses[k.slice(0, 7)] || { dias: {} }).dias[k] = dias[k];
  return C.normalizar({ meses, perfil: {} });
};
const K = '2026-11-04'; // quarta-feira
const agora = hhmm => new Date(K + 'T' + hhmm + ':00');

test('dia aberto só com a entrada: desconta o almoço depois de 6h, como se saísse agora', () => {
  const s = S({ [K]: { p: [L(K, '08:00')] } });
  assert.equal(C.trabalhado(s, K, agora('12:00')), 240);           // 4h: ainda sem almoço
  assert.equal(C.almocoAuto(s, K, agora('12:00')), 0);
  assert.equal(C.trabalhado(s, K, agora('17:18')), 558 - 90);      // 9h18 − 1h30 = 7h48
  assert.equal(C.almocoAuto(s, K, agora('17:18')), 90);
  assert.equal(C.saldoDia(s, K, agora('18:00')), 0);                // 08:00–18:00 = 0:00, igual ao dia fechado
  // fechando às 18:00 o resultado é o mesmo do "se sair agora"
  s.meses['2026-11'].dias[K].p.push(L(K, '18:00'));
  s.meses['2026-11'].dias[K].g.push(null);
  assert.equal(C.saldoDia(s, K, agora('18:30')), 0);
});

test('dia aberto com banco desconta 1h; com almoço registrado (3 marcações) não desconta', () => {
  const b = S({ [K]: { p: [L(K, '08:00')], banco: true } });
  assert.equal(C.almocoAuto(b, K, agora('18:00')), 60);
  assert.equal(C.saldoDia(b, K, agora('18:00')), 600 - 60 - 540);   // 0:00
  const r = S({ [K]: { p: [L(K, '08:00'), L(K, '12:00'), L(K, '13:00')] } });
  assert.equal(C.almocoAuto(r, K, agora('18:00')), 0);
  assert.equal(C.trabalhado(r, K, agora('18:00')), 240 + 300);
});

test('fim de semana e dia com tipo continuam sem desconto automático', () => {
  const sab = '2026-11-07';
  const s = S({ [sab]: { p: [L(sab, '08:00')] } });
  assert.equal(C.almocoAuto(s, sab, new Date(sab + 'T16:00:00')), 0);
  assert.equal(C.trabalhado(s, sab, new Date(sab + 'T16:00:00')), 480);
});
