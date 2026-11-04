'use strict';
// Achados da re-revisão r2 (f3). Hora LOCAL; rode também com TZ=America/Sao_Paulo.
const test = require('node:test');
const assert = require('node:assert/strict');
const ExcelJS = require('../vendor/exceljs.min');
const C = require('../js/core');
const K = require('../js/cartao');

const L = (k, hhmm) => new Date(k + 'T' + hhmm + ':00').toISOString();
const S = dias => {
  const meses = {};
  for (const k of Object.keys(dias)) (meses[k.slice(0, 7)] = meses[k.slice(0, 7)] || { dias: {} }).dias[k] = dias[k];
  return C.normalizar({ meses, perfil: {} });
};

test('r2 #1: dia anterior aberto + hoje com marcações pares → batida vai para hoje, anterior pendente e fora do saldo', () => {
  const s = S({
    '2026-11-04': { p: [L('2026-11-04', '23:55')] },
    '2026-11-05': { p: [L('2026-11-05', '08:00'), L('2026-11-05', '12:00')] }
  });
  const agora = new Date('2026-11-05T14:00');
  assert.equal(C.diaAberto(s, agora), null);
  const d = C.destinoBatida(s, agora);
  assert.deepEqual([d.k, d.acao, d.anterior], ['2026-11-05', 'entrada', false]);
  assert.deepEqual(d.pendente, { k: '2026-11-04', desde: L('2026-11-04', '23:55') });
  assert.equal(C.contaNoSaldo(s, '2026-11-04', agora), false);
  assert.equal(C.contaNoSaldo(s, '2026-11-05', agora), true);
  // a próxima batida (ENTRADA) entra em hoje, sem tocar no dia 04
  const dec = C.decidirBatida(s, agora, 'entrada', null);
  assert.equal(dec.ok, true);
  assert.equal(dec.destino.k, '2026-11-05');
  assert.equal(C.decidirBatida(s, agora, 'saida', null).motivo, 'paridade');
  // o saldo do mês ignora o 04 (pendente): só o 05 (ainda sem fechar = 4h de 8h30, 12:00 fechado → 240 − 510)
  assert.equal(C.resumo(s, '2026-11', agora).saldo, -270);
  // hoje aberto de novo: o anterior continua pendente
  C.dia(s, '2026-11-05').p.push(L('2026-11-05', '13:00'));
  const d2 = C.destinoBatida(s, new Date('2026-11-05T15:00'));
  assert.deepEqual([d2.k, d2.acao, d2.pendente && d2.pendente.k], ['2026-11-05', 'saida', '2026-11-04']);
});

test('r2 #1: dia ímpar de 3 dias atrás com hoje vazio fica pendente e fora do saldo até ser corrigido', () => {
  const s = S({ '2026-11-02': { p: [L('2026-11-02', '08:00')], tipo: '' }, '2026-11-03': { p: [L('2026-11-03', '08:00'), L('2026-11-03', '18:00')] } });
  const agora = new Date('2026-11-05T10:00');
  const d = C.destinoBatida(s, agora);
  assert.deepEqual([d.k, d.acao, d.anterior, d.pendente && d.pendente.k], ['2026-11-05', 'entrada', false, '2026-11-02']);
  assert.equal(C.contaNoSaldo(s, '2026-11-02', agora), false);
  // 02 (Finados, com marcação → regra normal) fica fora; 03 = 0; 04 = −510; 05 hoje vazio = −510
  assert.equal(C.resumo(s, '2026-11', agora).saldo, -1020);
  // corrigido no editor (saída às 17:00): volta a contar e deixa de ser pendente
  const ed = C.montarMarcacoes('2026-11-02', ['08:00', '17:00'], s.meses['2026-11'].dias['2026-11-02'].p, s.meses['2026-11'].dias['2026-11-02'].g);
  Object.assign(s.meses['2026-11'].dias['2026-11-02'], { p: ed.p, g: ed.g });
  assert.equal(C.contaNoSaldo(s, '2026-11-02', agora), true);
  assert.equal(C.destinoBatida(s, agora).pendente, null);
  assert.equal(C.resumo(s, '2026-11', agora).saldo, -1020 - 60);
});

test('r2 #1: o cartão Excel mostra o dia pendente sem saída e fora dos totais', async () => {
  const s = S({
    '2026-11-04': { p: [L('2026-11-04', '23:55')] },
    '2026-11-05': { p: [L('2026-11-05', '08:00'), L('2026-11-05', '18:00')] }
  });
  const agora = new Date('2026-11-05T20:00');
  const rb = new ExcelJS.Workbook();
  await rb.xlsx.load(await K.montar(ExcelJS, s, '2026-11', agora).xlsx.writeBuffer());
  const ws = rb.getWorksheet('Plan1'), linha = ws.getRow(6 + 4);
  assert.deepEqual([linha.getCell(2).value, linha.getCell(3).value], ['23:55', '_____ : _____ ']);
  const tot = {};
  for (let r = 7 + 30 + 2; r <= ws.rowCount; r++) tot[ws.getCell('A' + r).value] = ws.getCell('C' + r).value;
  assert.equal(tot['Saldo do mês'], C.formatHM(C.resumo(s, '2026-11', agora).saldo));
  assert.equal(C.resumo(s, '2026-11', agora).saldo, 0); // só o 05 (08–18 = 510 − 510); 04 pendente fora
});

test('r2 #A: mesclarMeses resolve campo a campo (cenário do Crivo)', () => {
  const g = { lat: -23.5, lon: -46.6, acc: 9 };
  const loc = { '2026-11': { dias: { '2026-11-04': { p: [L('2026-11-04', '08:00')], g: [g], obs: 'manual', ajuste: 30, banco: true } } } };
  const rem = { '2026-11': { dias: { '2026-11-04': { p: [L('2026-11-04', '08:00'), L('2026-11-04', '18:00')], g: [null, null], obs: 'cloud' } } } };
  const r = C.mesclarMeses(loc, rem), d = r.meses['2026-11'].dias['2026-11-04'];
  assert.deepEqual(d.p, rem['2026-11'].dias['2026-11-04'].p);   // mais marcações vence…
  // r3/f4 (autorizado pelo Farol): era [null, null] ("GPS anda junto com p"); agora g é por marcação,
  // o GPS local do 08:00 é mantido e o 18:00 (só remoto, sem GPS) fica null
  assert.deepEqual(d.g, [g, null]);
  assert.equal(d.obs, 'manual');                                // ambos com obs diferente: local
  assert.equal(d.banco, true);                                  // OU
  assert.equal(d.ajuste, 30);                                   // local ≠ 0
  assert.deepEqual(r.reenviar, ['2026-11']);
  // outros lados de cada regra
  const r2 = C.mesclarMeses(
    { '2026-11': { dias: { '2026-11-05': { p: [L('2026-11-05', '08:00'), L('2026-11-05', '18:00')], g: [g, null], obs: '', tipo: '', ajuste: 0, banco: false } } } },
    { '2026-11': { dias: { '2026-11-05': { p: [L('2026-11-05', '08:30')], obs: 'remoto', tipo: 'ATESTADO', ajuste: -15, banco: true } } } });
  const e = r2.meses['2026-11'].dias['2026-11-05'];
  assert.deepEqual([e.p.length, e.g[0]], [2, g]);                // local tem mais marcações: p+g locais
  assert.deepEqual([e.obs, e.tipo, e.ajuste, e.banco], ['remoto', 'ATESTADO', -15, true]);
  // empate de marcações: p+g locais
  const r3 = C.mesclarMeses(
    { '2026-11': { dias: { '2026-11-06': { p: [L('2026-11-06', '08:00')], g: [g] } } } },
    { '2026-11': { dias: { '2026-11-06': { p: [L('2026-11-06', '08:05')], g: [null] } } } });
  assert.deepEqual(r3.meses['2026-11'].dias['2026-11-06'].p, [L('2026-11-06', '08:00')]);
  assert.deepEqual(r3.meses['2026-11'].dias['2026-11-06'].g, [g]);
});
