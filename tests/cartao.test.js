'use strict';
const test = require('node:test');
const assert = require('node:assert/strict');
const ExcelJS = require('../vendor/exceljs.min');
const C = require('../js/core');
const K = require('../js/cartao');

const L = (k, hhmm) => new Date(k + 'T' + hhmm + ':00').toISOString();

test('cartão de ponto: modelo original completo, GPS, obs e totais com ajuste', async () => {
  const g = { lat: -23.55, lon: -46.63, acc: 9 };
  const s = C.normalizar({
    meses: { '2026-04': { dias: {
      '2026-04-01': { p: [L('2026-04-01', '08:00'), L('2026-04-01', '18:00')], g: [g, null] },
      '2026-04-02': { p: [L('2026-04-02', '22:00'), new Date('2026-04-03T02:00').toISOString()] },
      '2026-04-06': { p: [], ajuste: 30, obs: 'acerto' },
      '2026-04-07': { p: [], tipo: 'FOLGA' }
    } } },
    perfil: { nome: 'Pessoa Teste', saldoIni: 60 }
  });
  const agora = new Date('2026-04-30T20:00');
  const wb = K.montar(ExcelJS, s, '2026-04', agora);
  const buf = await wb.xlsx.writeBuffer();
  const rb = new ExcelJS.Workbook();
  await rb.xlsx.load(buf);
  const ws = rb.getWorksheet('Plan1');
  const v = a => ws.getCell(a).value;
  assert.equal(v('A1'), 'Cartão de Ponto');
  assert.match(v('A3'), /intervalo para refeições/);
  assert.equal(v('A4'), 'MÊS: ABRIL/2026');
  assert.equal(v('A5'), 'NOME:');
  assert.equal(v('B5'), 'Pessoa Teste');
  assert.deepEqual(ws.getRow(6).values.slice(1), ['Data ', 'Horário Entrada', 'Saída Almoço', 'Volta Almoço', 'Horário Saída', 'Rubrica Funcionário', 'Localização (GPS)', 'Observação']);
  assert.equal(ws.getCell('A6').fill.fgColor.argb, 'FFD9D9D9');
  const linha = d => ws.getRow(6 + d);
  assert.equal(linha(1).getCell(1).value.toISOString().slice(0, 10), '2026-04-01');
  assert.deepEqual([2, 3, 4, 5].map(c => linha(1).getCell(c).value), ['08:00', '—', '—', '18:00']);
  assert.match(linha(1).getCell(8).value, /Almoço de 01:30 descontado \(sem marcação\)/);
  assert.match(linha(1).getCell(7).value.hyperlink, /maps\.google\.com\/\?q=-23\.55,-46\.63/);
  assert.equal(linha(2).getCell(5).value, '02:00 (+1d)');
  assert.doesNotMatch(String(linha(2).getCell(8).value || ''), /Almoço/);   // 4h: sem almoço automático
  assert.equal(linha(3).getCell(2).value, 'FERIADO');         // Sexta-feira Santa automática
  assert.match(linha(3).getCell(8).value, /Sexta-feira Santa/);
  assert.equal(linha(4).getCell(2).value, 'SÁBADO');
  assert.equal(linha(5).getCell(3).value, 'DOMINGO');
  assert.equal(linha(7).getCell(2).value, 'FOLGA');
  assert.match(linha(6).getCell(8).value, /acerto · Ajuste \+00:30/);
  assert.equal(linha(8).getCell(2).value, '_____ : _____ ');
  assert.equal(v('A' + (7 + 30)), 'ASSINATURA:');
  const tot = {};
  for (let r = 7 + 30 + 2; r <= ws.rowCount; r++) tot[ws.getCell('A' + r).value] = ws.getCell('C' + r).value;
  const r = C.resumo(s, '2026-04', agora);
  assert.equal(tot['Saldo do mês'], C.formatHM(r.saldo));
  assert.equal(tot['Ajustes do mês (já incluídos)'], '00:30');
  assert.equal(tot['Banco de horas total'], C.formatHM(C.total(s, agora)));
  assert.equal(tot['Horas trabalhadas'], C.formatHM(r.tr));
  // o ajuste entra no saldo: sem ele o saldo seria 30 menor
  const semAjuste = C.normalizar(JSON.parse(JSON.stringify(s)));
  semAjuste.meses['2026-04'].dias['2026-04-06'].ajuste = 0;
  assert.equal(C.resumo(semAjuste, '2026-04', agora).saldo, r.saldo - 30);
  assert.equal(ws.pageSetup.paperSize, 9);
  assert.equal(ws.pageSetup.fitToPage, true);
  assert.equal(K.nomeArquivo('2026-04'), 'PONTO VIC 042026.xlsx');
});

test('cartão: almoço batido vai para Saída Almoço e Volta Almoço; 3 marcações deixam a saída em branco', async () => {
  const s = C.normalizar({ meses: { '2026-11': { dias: {
    '2026-11-03': { p: ['08:00', '12:00', '13:30', '18:00'].map(h => L('2026-11-03', h)) },
    '2026-11-04': { p: ['08:00', '12:10', '13:10'].map(h => L('2026-11-04', h)), banco: true },
    '2026-11-05': { p: ['08:00', '12:00'].map(h => L('2026-11-05', h)), almoco: true }
  } } }, perfil: {} });
  const ws = (await (async () => { const rb = new ExcelJS.Workbook(); await rb.xlsx.load(await K.montar(ExcelJS, s, '2026-11', new Date('2026-11-30T20:00')).xlsx.writeBuffer()); return rb.getWorksheet('Plan1'); })());
  const cel = (d, c) => ws.getRow(6 + d).getCell(c).value;
  assert.deepEqual([2, 3, 4, 5].map(c => cel(3, c)), ['08:00', '12:00', '13:30', '18:00']);
  assert.doesNotMatch(String(cel(3, 8) || ''), /descontado/);
  assert.deepEqual([2, 3, 4, 5].map(c => cel(4, c)), ['08:00', '12:10', '13:10', '_____ : _____ ']);
  assert.deepEqual([2, 3, 4, 5].map(c => cel(5, c)), ['08:00', '12:00', '_____ : _____ ', '_____ : _____ ']);
  assert.equal(cel(7, 3), 'SÁBADO');
  assert.equal(ws.getCell('A1').isMerged && ws.getCell('H1').master.address, 'A1');
});
