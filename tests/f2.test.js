'use strict';
// Achados da revisão r1 (f2). Hora LOCAL em tudo; rode também com TZ=America/Sao_Paulo.
const test = require('node:test');
const assert = require('node:assert/strict');
const ExcelJS = require('../vendor/exceljs.min');
const C = require('../js/core');
const I = require('../js/importar');

const L = (k, hhmm) => new Date(k + 'T' + hhmm + ':00').toISOString();
const S = (dias, perfil) => {
  const meses = {};
  for (const k of Object.keys(dias || {})) (meses[k.slice(0, 7)] = meses[k.slice(0, 7)] || { dias: {} }).dias[k] = dias[k];
  return C.normalizar({ meses, perfil: perfil || {} });
};
/** Simula a batida do app: decide o destino e grava. */
function bate(s, agoraTxt, acao, ultimo) {
  const agora = new Date(agoraTxt), dec = C.decidirBatida(s, agora, acao, ultimo);
  if (dec.ok) { const x = C.dia(s, dec.destino.k); x.p.push(agora.toISOString()); x.g.push(null); }
  return dec;
}

// ---------- 1. dia aberto atravessando a meia-noite ----------
test('meia-noite: ENTRADA 23:55 e SAÍDA 00:05 ficam no mesmo dia, com +1 dia', () => {
  const s = S({});
  assert.equal(bate(s, '2026-11-04T23:55', 'entrada').ok, true);
  const d = C.destinoBatida(s, new Date('2026-11-05T00:05'));
  assert.deepEqual([d.k, d.acao, d.anterior], ['2026-11-04', 'saida', true]);
  assert.equal(d.desde, L('2026-11-04', '23:55'));
  // enquanto aberto: conta até agora e não entra no saldo
  assert.equal(C.trabalhado(s, '2026-11-04', new Date('2026-11-05T00:03')), 8);
  assert.equal(C.contaNoSaldo(s, '2026-11-04', new Date('2026-11-05T00:03')), false);
  const r = bate(s, '2026-11-05T00:05', 'saida');
  assert.equal(r.ok, true);
  assert.equal(r.destino.k, '2026-11-04');
  const x = C.getDia(s, '2026-11-04');
  assert.deepEqual(x.p, [L('2026-11-04', '23:55'), L('2026-11-05', '00:05')]);
  assert.deepEqual(C.horariosDoDia('2026-11-04', x.p).map(h => h.dias), [0, 1]);   // "+1 dia"
  assert.equal(C.getDia(s, '2026-11-05'), null);                                      // dia novo não foi aberto
  assert.equal(C.trabalhado(s, '2026-11-04', new Date('2026-11-05T09:00')), 10);
  assert.equal(C.contaNoSaldo(s, '2026-11-04', new Date('2026-11-05T09:00')), true);
  // a próxima batida já é ENTRADA no dia novo
  const d2 = C.destinoBatida(s, new Date('2026-11-05T08:00'));
  assert.deepEqual([d2.k, d2.acao, d2.anterior, d2.pendente], ['2026-11-05', 'entrada', false, null]);
});

test('meia-noite: 22:00 → 07:00 (9 h) fecha o dia anterior', () => {
  const s = S({ '2026-11-04': { p: [L('2026-11-04', '22:00')] } });
  const r = bate(s, '2026-11-05T07:00', 'saida');
  assert.equal(r.ok, true);
  assert.equal(r.destino.k, '2026-11-04');
  assert.deepEqual(C.getDia(s, '2026-11-04').p, [L('2026-11-04', '22:00'), L('2026-11-05', '07:00')]);
  assert.equal(C.trabalhado(s, '2026-11-04', new Date('2026-11-05T12:00')), 540 - 90); // 2 marcações em dia útil: regra do almoço
});

test('marcação aberta de 3 dias atrás (ou há 16 h+): batida vai para hoje e o dia fica pendente', () => {
  const s = S({ '2026-11-01': { p: [L('2026-11-01', '08:00')] } });
  const d = C.destinoBatida(s, new Date('2026-11-04T10:00'));
  assert.deepEqual([d.k, d.acao, d.anterior], ['2026-11-04', 'entrada', false]);
  assert.deepEqual(d.pendente, { k: '2026-11-01', desde: L('2026-11-01', '08:00') });
  assert.equal(bate(s, '2026-11-04T10:00', 'saida').ok, false);   // SAÍDA não fecha um dia velho
  assert.equal(bate(s, '2026-11-04T10:00', 'entrada').ok, true);
  assert.deepEqual(C.getDia(s, '2026-11-04').p, [L('2026-11-04', '10:00')]);
  assert.deepEqual(C.getDia(s, '2026-11-01').p, [L('2026-11-01', '08:00')]);
  // limite exato de 16 h não continua o dia
  const s2 = S({ '2026-11-04': { p: [L('2026-11-04', '18:00')] } });
  assert.equal(C.destinoBatida(s2, new Date('2026-11-05T09:59')).anterior, true);
  const d16 = C.destinoBatida(s2, new Date('2026-11-05T10:00'));
  assert.deepEqual([d16.k, d16.acao, d16.pendente && d16.pendente.k], ['2026-11-05', 'entrada', '2026-11-04']);
});

test('desfazer depois de fechar o dia anterior reabre aquele dia', () => {
  const s = S({ '2026-11-04': { p: [L('2026-11-04', '23:55')] } });
  const r = bate(s, '2026-11-05T00:05', 'saida');
  const iso = new Date('2026-11-05T00:05').toISOString();
  assert.equal(C.removerUltima(s, r.destino.k, iso), true);
  assert.deepEqual(C.getDia(s, '2026-11-04').p, [L('2026-11-04', '23:55')]);
  assert.equal(C.removerUltima(s, r.destino.k, iso), false);           // não remove duas vezes
  const d = C.destinoBatida(s, new Date('2026-11-05T00:06'));
  assert.deepEqual([d.k, d.acao, d.anterior], ['2026-11-04', 'saida', true]);
  // desfazer a única marcação de um dia remove o dia (exceto se tiver "pagando banco")
  const s2 = S({ '2026-11-05': { p: [L('2026-11-05', '08:00')], banco: true } });
  assert.equal(C.removerUltima(s2, '2026-11-05', L('2026-11-05', '08:00')), true);
  assert.equal(C.getDia(s2, '2026-11-05').banco, true);
});

// ---------- 6. toque duplo ----------
test('toque duplo: paridade e trava de 1,5 s', () => {
  const s = S({});
  const t0 = new Date('2026-11-04T08:00:00');
  const a = bate(s, t0, 'entrada', null);
  assert.equal(a.ok, true);
  const ultimo = t0.getTime();
  // segundo toque no mesmo botão (ainda ENTRADA) 300 ms depois: repetido
  assert.deepEqual([bate(s, new Date(ultimo + 300), 'entrada', ultimo).ok, C.decidirBatida(s, new Date(ultimo + 300), 'entrada', ultimo).motivo], [false, 'repetido']);
  // mesmo depois de 1,5 s, ENTRADA com o dia aberto é recusada pela paridade
  assert.equal(C.decidirBatida(s, new Date(ultimo + 2000), 'entrada', ultimo).motivo, 'paridade');
  // SAÍDA antes de 1,5 s também é recusada (toque duplo atravessando a troca de botão)
  assert.equal(C.decidirBatida(s, new Date(ultimo + 1000), 'saida', ultimo).motivo, 'repetido');
  assert.equal(bate(s, new Date(ultimo + 1600), 'saida', ultimo).ok, true);
  assert.equal(C.getDia(s, '2026-11-04').p.length, 2);
  // SAÍDA com o dia fechado é recusada
  assert.equal(C.decidirBatida(s, new Date('2026-11-04T12:00'), 'saida', null).motivo, 'paridade');
  // com o dia anterior aberto (< 16 h), só SAÍDA vale
  const s2 = S({ '2026-11-04': { p: [L('2026-11-04', '23:55')] } });
  assert.equal(C.decidirBatida(s2, new Date('2026-11-05T00:05'), 'entrada', null).motivo, 'paridade');
});

// ---------- 2. backup ----------
const dia = { p: [L('2026-10-05', '08:00'), L('2026-10-05', '18:00')], g: [{ lat: -23.5, lon: -46.6, acc: 10 }, null], banco: false, tipo: '', obs: 'ok', ajuste: 15 };
const base = () => ({ meses: { '2026-10': { dias: { '2026-10-05': JSON.parse(JSON.stringify(dia)), '2026-10-06': { p: [], tipo: 'FOLGA' } } } }, perfil: { nome: 'X', saldoIni: 60, feriadosAuto: true, ultimoBackup: null } });
const env = dados => ({ app: 'ponto-vic', versao: 2, exportadoEm: '2026-11-04T12:00:00.000Z', dados });

test('backup: envelope novo e formato antigo aceitos, com resumo', () => {
  const a = C.validarBackup(env(base()));
  assert.equal(a.ok, true, a.erro);
  assert.deepEqual(a.resumo, { meses: 1, dias: 2, de: '2026-10-05', ate: '2026-10-06' });
  assert.equal(a.dados.meses['2026-10'].dias['2026-10-05'].ajuste, 15);
  const b = C.validarBackup(base());
  assert.equal(b.ok, true, b.erro);
  // export da f1b ({app:'Ponto VIC', versao, exportadoEm, meses, perfil}) também passa como formato antigo
  assert.equal(C.validarBackup(Object.assign({ app: 'Ponto VIC', versao: 2, exportadoEm: '2026-11-01T00:00:00Z' }, base())).ok, true);
});

test('backup: truncado, envelope ruim, data impossível, campos inválidos e GPS malicioso são rejeitados', () => {
  const erro = v => { const r = C.validarBackup(v); assert.equal(r.ok, false, 'devia rejeitar: ' + JSON.stringify(v).slice(0, 120)); return r.erro; };
  assert.match(erro({ meses: {} }), /perfil/);                       // truncado: sem perfil
  assert.match(erro({ perfil: {} }), /meses/);
  assert.match(erro(null), /backup/);
  assert.match(erro([]), /backup/);
  assert.match(erro({ app: 'ponto-vic', versao: 2, exportadoEm: '2026-11-04T12:00:00Z' }), /dados/);
  assert.match(erro(env(base()) && Object.assign(env(base()), { versao: 3 })), /versão/);
  assert.match(erro(Object.assign(env(base()), { exportadoEm: 'ontem' })), /exportação/);
  assert.match(erro(Object.assign(env(base()), { app: 'outro' })), /envelope/);
  const com = f => { const b = base(); f(b); return b; };
  assert.match(erro(com(b => { b.meses['2026-02'] = { dias: { '2026-02-31': { p: [] } } }; })), /impossível/);
  assert.match(erro(com(b => { b.meses['2026-13'] = { dias: {} }; })), /mês inválido/);
  assert.match(erro(com(b => { b.meses['2026-10'].dias['2026-11-01'] = { p: [] }; })), /está no mês/);
  assert.match(erro(com(b => { b.meses['2026-10'].dias['2026-10-05'].p = ['ontem']; })), /marcações inválidas/);
  assert.match(erro(com(b => { b.meses['2026-10'].dias['2026-10-05'].p.reverse(); })), /fora de ordem/);
  assert.match(erro(com(b => { b.meses['2026-10'].dias['2026-10-05'].tipo = 'FESTA'; })), /tipo/);
  assert.match(erro(com(b => { b.meses['2026-10'].dias['2026-10-05'].banco = 'sim'; })), /banco/);
  assert.match(erro(com(b => { b.meses['2026-10'].dias['2026-10-05'].ajuste = '30'; })), /ajuste/);
  assert.match(erro(com(b => { b.meses['2026-10'].dias['2026-10-05'].obs = 5; })), /observação/);
  assert.match(erro(com(b => { b.perfil.saldoIni = 'muito'; })), /saldo/);
  assert.match(erro(com(b => { b.meses['2026-10'].dias['2026-10-05'].g = [{ lat: '1"><img src=x onerror=alert(1)>', lon: '2', acc: 1 }]; })), /GPS/);
  assert.match(erro(com(b => { b.meses['2026-10'].dias['2026-10-05'].g = [{ lat: 95, lon: 0, acc: 1 }]; })), /GPS/);
  assert.match(erro(com(b => { b.meses['2026-10'].dias['2026-10-05'].g = [{ lat: 1, lon: 2, acc: -3 }]; })), /GPS/);
  assert.match(erro(com(b => { b.meses['2026-10'].dias['2026-10-05'].g = 'x'; })), /GPS/);
});

// ---------- 5. GPS no normalizar ----------
test('normalizar: GPS só com números finitos dentro dos limites', () => {
  const s = C.normalizar({ meses: { '2026-10': { dias: { '2026-10-05': {
    p: [L('2026-10-05', '08:00'), L('2026-10-05', '12:00'), L('2026-10-05', '13:00'), L('2026-10-05', '18:00'), L('2026-10-05', '19:00')],
    g: [{ lat: '1"><img src=x onerror=alert(1)>', lon: '2', acc: 1 }, { lat: 91, lon: 0, acc: 1 }, { lat: -23.5, lon: -46.6, acc: 8, x: '<b>' }, { lat: 1, lon: 2 }, { lat: NaN, lon: 1, acc: 1 }]
  } } } } });
  assert.deepEqual(s.meses['2026-10'].dias['2026-10-05'].g, [null, null, { lat: -23.5, lon: -46.6, acc: 8 }, null, null]);
  assert.equal(C.gpsValido({ lat: 0, lon: 180, acc: 0 }), true);
  assert.equal(C.gpsValido({ lat: 0, lon: 180.1, acc: 0 }), false);
});

// ---------- 3. merge da nuvem ----------
test('mesclarMeses: por dia, vence quem tem mais marcações; empate fica o local; reenvia o que mudou', () => {
  const loc = {
    '2026-11': { dias: {
      '2026-11-03': { p: [L('2026-11-03', '08:00'), L('2026-11-03', '18:00')] },
      '2026-11-04': { p: [L('2026-11-04', '08:00')] },                                 // batida feita durante a hidratação
      '2026-11-05': { p: [L('2026-11-05', '08:00')], obs: 'local' }
    } },
    '2026-12': { dias: { '2026-12-01': { p: [], tipo: 'FOLGA' } } }                       // só local
  };
  const rem = {
    '2026-11': { dias: {
      '2026-11-03': { p: [L('2026-11-03', '08:00'), L('2026-11-03', '12:00'), L('2026-11-03', '13:00'), L('2026-11-03', '18:00')] },
      '2026-11-05': { p: [L('2026-11-05', '08:30')], obs: 'remoto' },
      '2026-11-06': { p: [L('2026-11-06', '08:00')] }                                 // só remoto
    } },
    '2026-10': { dias: { '2026-10-01': { p: [L('2026-10-01', '08:00'), L('2026-10-01', '18:00')] } } }
  };
  const r = C.mesclarMeses(loc, rem);
  const d = r.meses['2026-11'].dias;
  assert.equal(d['2026-11-03'].p.length, 4);                 // remoto tem mais marcações
  assert.deepEqual(d['2026-11-04'].p, [L('2026-11-04', '08:00')]);   // batida local não some
  assert.equal(d['2026-11-05'].obs, 'local');                // empate: local
  assert.equal(d['2026-11-06'].p.length, 1);
  assert.equal(r.meses['2026-10'].dias['2026-10-01'].p.length, 2);
  assert.equal(r.meses['2026-12'].dias['2026-12-01'].tipo, 'FOLGA');
  assert.deepEqual(r.reenviar, ['2026-11', '2026-12']);      // outubro igual ao remoto: não reenvia
  // iguais: nada a reenviar
  assert.deepEqual(C.mesclarMeses(rem, rem).reenviar, []);
});

// ---------- 4. importação x "pagando banco" ----------
test('importação: dia só com "pagando banco" é conflito; com sobrescrever o banco local é preservado', async () => {
  const wb = new ExcelJS.Workbook(), ws = wb.addWorksheet('Out');
  ws.getRow(2).values = ['Dia', 'Dia', 'Entrada', 'Saída', 'Resumo', '', 'Obs'];
  ['2026-10-05', '2026-10-06', '2026-10-07'].forEach((k, i) => {
    const r = ws.getRow(3 + i);
    r.getCell(2).value = new Date(k + 'T00:00:00Z'); r.getCell(2).numFmt = 'dd/mm/yyyy';
    r.getCell(3).value = new Date('1899-12-30T08:00:00Z'); r.getCell(3).numFmt = 'hh:mm';
    r.getCell(4).value = new Date('1899-12-30T18:00:00Z'); r.getCell(4).numFmt = 'hh:mm';
  });
  const rb = new ExcelJS.Workbook(); await rb.xlsx.load(await wb.xlsx.writeBuffer());
  const { linhas } = I.lerPlanilha(rb);
  const s = C.normalizar({ meses: { '2026-10': { dias: { '2026-10-06': { p: [], banco: true } } } } });
  assert.equal(C.temDados(C.getDia(s, '2026-10-06')), true);
  assert.equal(C.inicio(s), null, 'banco sozinho não marca o início do banco de horas');
  const agora = new Date('2026-10-07T23:00');
  const p1 = I.planejar(s, linhas, { agora });
  assert.equal(p1.total.ignorados, 1);
  const o1 = I.aplicar(s, p1);
  assert.deepEqual(C.getDia(o1, '2026-10-06'), C.getDia(s, '2026-10-06'));
  const p2 = I.planejar(s, linhas, { sobrescrever: true, agora });
  const o2 = I.aplicar(s, p2), d = C.getDia(o2, '2026-10-06');
  assert.equal(d.banco, true);
  assert.equal(d.p.length, 2);
  // com banco o previsto é 540; o ajuste faz o dia bater com a planilha (0)
  assert.equal(C.saldoDia(o2, '2026-10-06', agora), 0);
  assert.equal(C.total(o2, agora), 0);
});
