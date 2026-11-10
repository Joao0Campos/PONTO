'use strict';
// Botões SAÍDA ALMOÇO / VOLTA ALMOÇO e previsão da volta (10/11/2026): 1h30 no dia normal, 1h quando paga banco.
const test = require('node:test');
const assert = require('node:assert/strict');
const C = require('../js/core');

const L = (k, hhmm) => new Date(k + 'T' + hhmm + ':00').toISOString();
const S = dias => {
  const meses = {};
  for (const k of Object.keys(dias)) (meses[k.slice(0, 7)] = meses[k.slice(0, 7)] || { dias: {} }).dias[k] = dias[k];
  return C.normalizar({ meses, perfil: {} });
};
const K = '2026-11-10'; // terça-feira
const agora = hhmm => new Date(K + 'T' + hhmm + ':00');
const hl = iso => C.horaLocal(iso);
const noAlmoco = (extra) => S({ [K]: Object.assign({ p: [L(K, '08:00'), L(K, '12:00')], almoco: true }, extra) });

test('dia normal: saiu para o almoço às 12:00, volta prevista às 13:30', () => {
  const v = C.voltaDoAlmoco(noAlmoco(), K);
  assert.deepEqual([v.almoco, hl(v.saiu), hl(v.volta)], [90, '12:00', '13:30']);
});

test('pagando banco: almoço de 1h, volta às 13:00', () => {
  const v = C.voltaDoAlmoco(noAlmoco({ banco: true }), K);
  assert.deepEqual([v.almoco, hl(v.volta)], [60, '13:00']);
});

test('a saída comum (sem o botão do almoço) não vira almoço', () => {
  assert.equal(C.voltaDoAlmoco(S({ [K]: { p: [L(K, '08:00'), L(K, '12:00')] } }), K), null);
  assert.equal(C.nomeMarcacao(1, 2), 'Saída');
  assert.equal(C.nomeMarcacao(1, 2, true), 'Saída almoço');
});

test('depois da VOLTA ALMOÇO a previsão some e vale o almoço real, sem desconto automático', () => {
  const s = S({ [K]: { p: ['08:00', '12:00', '13:10'].map(h => L(K, h)), almoco: true } });
  assert.equal(C.voltaDoAlmoco(s, K), null);
  s.meses['2026-11'].dias[K].p.push(L(K, '18:00'));
  assert.equal(C.almocoAuto(s, K, agora('18:30')), 0);
  assert.equal(C.trabalhado(s, K, agora('18:30')), 240 + 290);
  assert.equal(C.nomeMarcacao(2, 4, true), 'Volta almoço');
});

test('no almoço não há desconto automático, mesmo com mais de 6h antes da saída', () => {
  const s = S({ [K]: { p: [L(K, '06:00'), L(K, '12:30')], almoco: true } });
  assert.equal(C.almocoAuto(s, K, agora('12:40')), 0);
  assert.equal(C.trabalhado(s, K, agora('12:40')), 390);
});

test('desfazer a SAÍDA ALMOÇO tira a marca de almoço; normalizar não deixa almoço sem as 2 marcações', () => {
  const s = noAlmoco(), d = s.meses['2026-11'].dias[K];
  assert.equal(C.removerUltima(s, K, d.p[1]), true);
  assert.equal(d.almoco, false);
  const n = S({ [K]: { p: [L(K, '08:00')], almoco: true } });
  assert.equal(n.meses['2026-11'].dias[K].almoco, false);
  assert.equal(S({ [K]: { p: [L(K, '08:00'), L(K, '12:00')], almoco: 'sim' } }).meses['2026-11'].dias[K].almoco, false);
});

test('mesclar com a nuvem: o almoço acompanha as marcações que vencem', () => {
  const local = { p: [L(K, '08:00'), L(K, '12:00')], g: [null, null], almoco: true };
  const remoto = { p: [L(K, '08:00')], g: [null] };
  assert.equal(C.mesclarDia(C.normalizar({ meses: { '2026-11': { dias: { [K]: local } } } }).meses['2026-11'].dias[K],
    C.normalizar({ meses: { '2026-11': { dias: { [K]: remoto } } } }).meses['2026-11'].dias[K]).almoco, true);
});
