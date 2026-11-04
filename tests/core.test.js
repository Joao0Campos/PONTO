'use strict';
// Rode também com TZ=America/Sao_Paulo: todas as marcações aqui são hora LOCAL.
const test = require('node:test');
const assert = require('node:assert/strict');
const C = require('../js/core');

/** ISO de um horário local. */
const L = (k, hhmm) => new Date(k + 'T' + hhmm + ':00').toISOString();
const S = (dias, perfil) => {
  const meses = {};
  for (const k of Object.keys(dias || {})) (meses[k.slice(0, 7)] = meses[k.slice(0, 7)] || { dias: {} }).dias[k] = dias[k];
  return C.normalizar({ meses, perfil: perfil || {} });
};
const dia = (k, hs, extra) => Object.assign({ p: hs.map(h => L(k, h)) }, extra || {});

test('trabalhado: 2 marcações descontam almoço; 4 não; banco; meio período; fim de semana; tipo', () => {
  const s = S({
    '2026-04-06': dia('2026-04-06', ['08:00', '18:00']),                       // seg
    '2026-04-07': dia('2026-04-07', ['08:00', '12:00', '13:00', '18:00']),     // ter, 4 marcações
    '2026-04-08': dia('2026-04-08', ['08:00', '18:00'], { banco: true }),
    '2026-04-09': dia('2026-04-09', ['08:00', '12:00']),                       // meio período (<= 6h)
    '2026-04-11': dia('2026-04-11', ['08:00', '18:00']),                       // sábado
    '2026-04-10': dia('2026-04-10', ['08:00', '18:00'], { tipo: 'FERIADO' })
  });
  const ag = new Date('2026-04-20T12:00');
  assert.equal(C.trabalhado(s, '2026-04-06', ag), 510);
  assert.equal(C.saldoDia(s, '2026-04-06', ag), 0);
  assert.equal(C.trabalhado(s, '2026-04-07', ag), 540);
  assert.equal(C.saldoDia(s, '2026-04-07', ag), 30);
  assert.equal(C.trabalhado(s, '2026-04-08', ag), 540);
  assert.equal(C.previsto(s, '2026-04-08'), 540);
  assert.equal(C.trabalhado(s, '2026-04-09', ag), 240);
  assert.equal(C.saldoDia(s, '2026-04-09', ag), -270);
  assert.equal(C.previsto(s, '2026-04-11'), 0);
  assert.equal(C.saldoDia(s, '2026-04-11', ag), 600);
  assert.equal(C.previsto(s, '2026-04-10'), 0);
  assert.equal(C.saldoDia(s, '2026-04-10', ag), 600);
  assert.equal(C.previsto(s, '2026-04-13'), 510); // dia útil sem nada
});

test('saldo do dia = trabalhado − previsto + ajuste', () => {
  const s = S({ '2026-04-06': dia('2026-04-06', ['08:00', '18:00'], { ajuste: 45 }), '2026-04-07': { p: [], ajuste: -30 } });
  assert.equal(C.saldoDia(s, '2026-04-06'), 45);
  assert.equal(C.saldoDia(s, '2026-04-07'), -540);
});

test('par aberto conta até agora só no dia de hoje', () => {
  const s = S({ '2026-04-06': dia('2026-04-06', ['08:00']), '2026-04-07': dia('2026-04-07', ['08:00']) });
  assert.equal(C.trabalhado(s, '2026-04-07', new Date('2026-04-07T10:30')), 150);
  assert.equal(C.trabalhado(s, '2026-04-06', new Date('2026-04-07T10:30')), 0);
  // hoje com marcação aberta não entra no saldo do mês
  const r = C.resumo(s, '2026-04', new Date('2026-04-07T10:30'));
  // r2/f3: era -510 (dia 06 aberto no passado contava 0 − 510). Regra nova, autorizada pelo Farol na f3:
  // dia passado com marcação ímpar fica pendente e NÃO entra no saldo até ser corrigido.
  assert.equal(r.saldo, 0);
  // o dia 06 aparece como pendente para correção
  const dest = C.destinoBatida(s, new Date('2026-04-07T10:30'));
  assert.equal(dest.pendente && dest.pendente.k, '2026-04-06');
  assert.equal(C.contaNoSaldo(s, '2026-04-06', new Date('2026-04-07T10:30')), false);
  // corrigido com a saída, volta a contar: 08:00–17:00 = 540 − 90 (almoço) − 510 = −60
  s.meses['2026-04'].dias['2026-04-06'].p.push(new Date('2026-04-06T17:00').toISOString());
  s.meses['2026-04'].dias['2026-04-06'].g.push(null);
  assert.equal(C.contaNoSaldo(s, '2026-04-06', new Date('2026-04-07T10:30')), true);
  assert.equal(C.resumo(s, '2026-04', new Date('2026-04-07T10:30')).saldo, -60);
  assert.equal(C.destinoBatida(s, new Date('2026-04-07T10:30')).pendente, null);
});

test('feriados: Sexta-feira Santa 2026-04-03 e Finados 2026-11-02', () => {
  assert.equal(C.pascoa(2026), '2026-04-05');
  assert.equal(C.pascoa(2025), '2025-04-20');
  assert.equal(C.nomeFeriado('2026-04-03'), 'Sexta-feira Santa');
  assert.equal(C.feriado('2026-04-02'), false);
  for (const md of ['01-01', '04-21', '05-01', '09-07', '10-12', '11-02', '11-15', '11-20', '12-25']) assert.ok(C.feriado('2026-' + md), md);
  const vazio = S({});
  assert.equal(C.previsto(vazio, '2026-04-03'), 0);
  assert.equal(C.previsto(vazio, '2026-11-02'), 0);                                   // sem marcação: previsto 0
  const comMarc = S({ '2026-11-02': dia('2026-11-02', ['08:00', '18:00']) });
  assert.equal(C.previsto(comMarc, '2026-11-02'), 510);                               // com marcação: regra normal
  assert.equal(C.saldoDia(comMarc, '2026-11-02', new Date('2026-11-03')), 0);
  const desligado = S({}, { feriadosAuto: false });
  assert.equal(C.previsto(desligado, '2026-11-02'), 510);
  assert.equal(C.feriadoAuto(desligado, '2026-11-02'), false);
  const comTipo = S({ '2026-11-02': { p: [], tipo: 'FOLGA' } });
  assert.equal(C.feriadoAuto(comTipo, '2026-11-02'), false);
  assert.equal(C.previsto(comTipo, '2026-11-02'), 0);
});

test('resumo e total: início, dias úteis vazios contam −previsto, extras/devidas, saldoIni, mês sem registro', () => {
  const s = S({
    '2026-04-01': dia('2026-04-01', ['08:00', '19:00']), // +60
    '2026-04-03': { p: [], obs: 'feriado' },               // Sexta Santa sem marcação: 0
    '2026-04-06': dia('2026-04-06', ['08:00', '17:00'])  // −60
  }, { saldoIni: 100 });
  assert.equal(C.inicio(s), '2026-04-01');
  const ag = new Date('2026-04-07T20:00');
  const r = C.resumo(s, '2026-04', ag);
  // 01 +60, 02 −510, 03 0, 06 −60, 07 −510 (hoje, sem marcação aberta)
  assert.deepEqual(r, { tr: 570 + 450, pv: 510 * 4, ex: 60, dv: 510 + 60 + 510, saldo: 60 - 1080 });
  assert.equal(C.total(s, ag), 100 + 60 - 1080);
  assert.equal(C.resumo(s, '2026-03', ag).saldo, 0); // antes do início não conta
  // mês sem nenhum registro entre o início e hoje também conta (dias úteis vazios)
  const s2 = S({ '2026-04-30': dia('2026-04-30', ['08:00', '18:00']) });
  const ag2 = new Date('2026-05-05T09:00');
  // 01/05 feriado, 04/05 −510, 05/05 (hoje, vazio) −510
  assert.deepEqual(C.mesesDoBanco(s2, ag2), ['2026-04', '2026-05']);
  assert.equal(C.total(s2, ag2), -1020);
  // ajuste e obs marcam o início
  assert.equal(C.inicio(S({ '2026-02-10': { p: [], ajuste: 30 }, '2026-02-12': { p: [], obs: 'x' } })), '2026-02-10');
});

test('migração: dado antigo (sem obs/ajuste/g/feriadosAuto) carrega e calcula igual', () => {
  const antigo = {
    meses: { '2026-01': { dias: { '2026-01-05': { p: [L('2026-01-05', '08:00'), L('2026-01-05', '18:00')], banco: false, tipo: '' },
      '2026-01-06': { p: [L('2026-01-06', '08:00'), L('2026-01-06', '17:00')], g: [{ lat: -23.5, lon: -46.6, acc: 10 }], banco: false, tipo: '', extra: 'mantido' } } } },
    perfil: { nome: 'Fulano', saldoIni: 60 }
  };
  const copia = JSON.parse(JSON.stringify(antigo));
  const s = C.normalizar(antigo);
  assert.deepEqual(antigo, copia, 'normalizar não muta a entrada');
  const d5 = s.meses['2026-01'].dias['2026-01-05'], d6 = s.meses['2026-01'].dias['2026-01-06'];
  assert.equal(d5.obs, ''); assert.equal(d5.ajuste, 0); assert.deepEqual(d5.g, [null, null]);
  assert.deepEqual(d6.g, [{ lat: -23.5, lon: -46.6, acc: 10 }, null]);
  assert.equal(d6.extra, 'mantido');
  assert.equal(s.perfil.feriadosAuto, true); assert.equal(s.perfil.ultimoBackup, null); assert.equal(s.perfil.nome, 'Fulano');
  assert.equal(C.total(s, new Date('2026-01-06T20:00')), 60 + 0 - 60);
  // lixo não quebra
  assert.deepEqual(C.normalizar(null).meses, {});
  assert.equal(C.normalizar({ meses: { '2026-01': null } }).meses['2026-01'].dias !== undefined, true);
  assert.deepEqual(C.normalizar({ meses: { '2026-01': { dias: { '2026-01-02': { p: 'x', ajuste: 'abc' } } } } }).meses['2026-01'].dias['2026-01-02'].p, []);
});

test('parseHM / formatHM', () => {
  assert.equal(C.parseHM('12:30'), 750);
  assert.equal(C.parseHM('-03:15'), -195);
  assert.equal(C.parseHM('+00:30'), 30);
  assert.equal(C.parseHM(' - 1:05 '), -65);
  assert.equal(C.parseHM('10:75'), null);
  assert.equal(C.parseHM('abc'), null);
  assert.equal(C.formatHM(-195), '-03:15');
  assert.equal(C.formatHM(0), '00:00');
  assert.equal(C.formatHMsinal(30), '+00:30');
  assert.equal(C.formatHMsinal(-30), '-00:30');
});

test('editor: montarMarcacoes usa hora local, virada da meia-noite (+1 dia) e não ordena', () => {
  const k = '2026-04-06';
  const a = C.montarMarcacoes(k, ['22:00', '02:00']);
  assert.deepEqual(a.dias, [0, 1]);
  assert.equal(new Date(a.p[0]).getHours(), 22);
  assert.equal(C.keyDate(new Date(a.p[1])), '2026-04-07');
  assert.equal(C.horaLocal(a.p[1]), '02:00');
  assert.equal(C.trabalhado(S({ [k]: { p: a.p } }), k), 240);
  // acumula a cada volta; vazios e inválidos são ignorados
  const b = C.montarMarcacoes(k, ['20:00', '', '23:00', '01:00', '00:30', '25:00', 'xx']);
  assert.deepEqual(b.dias, [0, 0, 1, 2]);
  assert.equal(C.keyDate(new Date(b.p[3])), '2026-04-08');
  // ordem digitada é mantida (sem sort)
  const c = C.montarMarcacoes(k, ['08:00', '18:00', '12:00']);
  assert.deepEqual(c.dias, [0, 0, 1]);
  // dia normal: 4 horários na mesma data, hora local exata
  const d = C.montarMarcacoes(k, ['08:00', '12:00', '13:30', '18:00']);
  assert.deepEqual(d.p, [L(k, '08:00'), L(k, '12:00'), L(k, '13:30'), L(k, '18:00')]);
  assert.deepEqual(d.dias, [0, 0, 0, 0]);
  assert.deepEqual(C.horariosDoDia(k, a.p), [{ hm: '22:00', dias: 0 }, { hm: '02:00', dias: 1 }]);
});

test('editor: GPS e segundos das marcações mantidas são preservados', () => {
  const k = '2026-04-06';
  const p0 = new Date(k + 'T08:00:37').toISOString(), p1 = new Date(k + 'T18:00:05').toISOString();
  const g0 = { lat: -23.1, lon: -46.1, acc: 8 }, g1 = { lat: -23.2, lon: -46.2, acc: 12 };
  const r = C.montarMarcacoes(k, ['08:00', '12:00', '13:00', '18:00'], [p0, p1], [g0, g1]);
  assert.equal(r.p[0], p0); assert.equal(r.p[3], p1);
  assert.deepEqual(r.g, [g0, null, null, g1]);
  // remover a primeira mantém o GPS da outra; mudar o horário perde o GPS daquela
  const r2 = C.montarMarcacoes(k, ['18:00'], [p0, p1], [g0, g1]);
  assert.deepEqual(r2.g, [g1]);
  const r3 = C.montarMarcacoes(k, ['08:05', '18:00'], [p0, p1], [g0, g1]);
  assert.deepEqual(r3.g, [null, g1]);
  // virada: marcação antiga do dia seguinte casa pelo deslocamento
  const pn = new Date('2026-04-07T02:00:10').toISOString(), gn = { lat: 1, lon: 2, acc: 3 };
  const r4 = C.montarMarcacoes(k, ['22:00', '02:00'], [L(k, '22:00'), pn], [null, gn]);
  assert.equal(r4.p[1], pn); assert.deepEqual(r4.g, [null, gn]);
});

test('nomes das marcações como no app original', () => {
  assert.deepEqual([0, 1].map(i => C.nomeMarcacao(i, 2)), ['Entrada', 'Saída']);
  assert.deepEqual([0, 1, 2, 3].map(i => C.nomeMarcacao(i, 4)), ['Entrada', 'Saída almoço', 'Volta almoço', 'Saída']);
  assert.deepEqual([4, 5].map(i => C.nomeMarcacao(i, 6)), ['Entrada', 'Saída']);
});

test('comDia não muta o estado original', () => {
  const s = S({ '2026-04-06': dia('2026-04-06', ['08:00', '18:00']) });
  const antes = JSON.stringify(s);
  const s2 = C.comDia(s, '2026-04-06', { p: [], ajuste: 10 });
  assert.equal(JSON.stringify(s), antes);
  assert.equal(C.saldoDia(s2, '2026-04-06'), -500);
});
