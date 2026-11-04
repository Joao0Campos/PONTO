'use strict';
// Pontos da confirmação r3 (f4). Hora LOCAL; rode também com TZ=America/Sao_Paulo.
const test = require('node:test');
const assert = require('node:assert/strict');
const C = require('../js/core');

const L = (k, hhmm) => new Date(k + 'T' + hhmm + ':00').toISOString();
const S = dias => {
  const meses = {};
  for (const k of Object.keys(dias)) (meses[k.slice(0, 7)] = meses[k.slice(0, 7)] || { dias: {} }).dias[k] = dias[k];
  return C.normalizar({ meses, perfil: {} });
};

test('r3 #1: GPS por marcação no merge (cenário do Crivo)', () => {
  const g = { lat: -23.55, lon: -46.63, acc: 7 };
  const loc = { '2026-11': { dias: { '2026-11-04': { p: [L('2026-11-04', '08:00')], g: [g], obs: 'manual', banco: true, ajuste: 20 } } } };
  const rem = { '2026-11': { dias: { '2026-11-04': { p: [L('2026-11-04', '08:00'), L('2026-11-04', '17:00')], g: [null, null] } } } };
  const d = C.mesclarMeses(loc, rem).meses['2026-11'].dias['2026-11-04'];
  assert.deepEqual(d.p, [L('2026-11-04', '08:00'), L('2026-11-04', '17:00')]);
  assert.deepEqual(d.g, [g, null]);
  assert.deepEqual([d.obs, d.banco, d.ajuste], ['manual', true, 20]);
  // GPS só do remoto para uma marcação que o local também tem sem GPS
  const g2 = { lat: 1, lon: 2, acc: 3 };
  const d2 = C.mesclarDia(
    C.normalizarDia({ p: [L('2026-11-05', '08:00'), L('2026-11-05', '18:00')], g: [null, null] }),
    C.normalizarDia({ p: [L('2026-11-05', '08:00')], g: [g2] }));
  assert.deepEqual(d2.g, [g2, null]);
  // local e remoto com GPS na mesma marcação: local primeiro; GPS inválido do local não vale
  const d3 = C.mesclarDia(
    { p: [L('2026-11-06', '08:00')], g: [{ lat: 999, lon: 0, acc: 1 }], obs: '', tipo: '', banco: false, ajuste: 0 },
    C.normalizarDia({ p: [L('2026-11-06', '08:00'), L('2026-11-06', '12:00')], g: [g2, g] }));
  assert.deepEqual(d3.g, [g2, g]);
  const d4 = C.mesclarDia(
    C.normalizarDia({ p: [L('2026-11-06', '08:00'), L('2026-11-06', '12:00')], g: [g, null] }),
    C.normalizarDia({ p: [L('2026-11-06', '08:00'), L('2026-11-06', '12:05')], g: [g2, g2] }));
  assert.deepEqual(d4.p, [L('2026-11-06', '08:00'), L('2026-11-06', '12:00')]);   // empate: p local
  assert.deepEqual(d4.g, [g, null]);   // 12:00 só existe no local (sem GPS); o 12:05 do remoto não casa
});

test('r3 #2: pendentes() devolve todos os dias incompletos, do mais recente ao mais antigo', () => {
  const s = S({
    '2026-11-01': { p: [L('2026-11-01', '08:00')] },
    '2026-11-03': { p: [L('2026-11-03', '07:50')] },
    '2026-11-04': { p: [L('2026-11-04', '08:00'), L('2026-11-04', '18:00')] }
  });
  const agora = new Date('2026-11-05T10:00');
  assert.deepEqual(C.pendentes(s, agora), [
    { k: '2026-11-03', desde: L('2026-11-03', '07:50') },
    { k: '2026-11-01', desde: L('2026-11-01', '08:00') }
  ]);
  assert.equal(C.contaNoSaldo(s, '2026-11-01', agora), false);
  assert.equal(C.contaNoSaldo(s, '2026-11-03', agora), false);
  assert.equal(C.destinoBatida(s, agora).pendente.k, '2026-11-03');
  // corrigindo o 03, sobra o 01
  s.meses['2026-11'].dias['2026-11-03'].p.push(L('2026-11-03', '17:00'));
  s.meses['2026-11'].dias['2026-11-03'].g.push(null);
  assert.deepEqual(C.pendentes(s, agora).map(x => x.k), ['2026-11-01']);
  assert.equal(C.contaNoSaldo(s, '2026-11-03', agora), true);
  assert.equal(C.destinoBatida(s, agora).pendente.k, '2026-11-01');
  // corrigindo o 01, nenhum
  s.meses['2026-11'].dias['2026-11-01'].p.push(L('2026-11-01', '12:00'));
  s.meses['2026-11'].dias['2026-11-01'].g.push(null);
  assert.deepEqual(C.pendentes(s, agora), []);
});

test('r3 #2: o dia aberto atual (hoje ou virada < 16 h) não é pendente; meses diferentes entram', () => {
  const s = S({
    '2026-10-30': { p: [L('2026-10-30', '08:00')] },
    '2026-11-04': { p: [L('2026-11-04', '23:55')] }
  });
  const agora = new Date('2026-11-05T00:05');
  assert.equal(C.diaAberto(s, agora), '2026-11-04');
  assert.deepEqual(C.pendentes(s, agora).map(x => x.k), ['2026-10-30']);
  // passou de 16 h: o 04 também vira pendente
  assert.deepEqual(C.pendentes(s, new Date('2026-11-05T16:00')).map(x => x.k), ['2026-11-04', '2026-10-30']);
  // hoje aberto não aparece
  const s2 = S({ '2026-11-05': { p: [L('2026-11-05', '08:00')] } });
  assert.deepEqual(C.pendentes(s2, new Date('2026-11-05T09:00')), []);
});
