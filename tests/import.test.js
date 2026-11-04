'use strict';
// Planilhas SINTÉTICAS (nenhum dado real). Rode também com TZ=America/Sao_Paulo.
const test = require('node:test');
const assert = require('node:assert/strict');
const ExcelJS = require('../vendor/exceljs.min');
const C = require('../js/core');
const I = require('../js/importar');

const L = (k, hhmm) => new Date(k + 'T' + hhmm + ':00').toISOString();
const DATA = k => new Date(k + 'T00:00:00Z');                 // como o Excel entrega: UTC
const HORA = hhmm => new Date('1899-12-30T' + hhmm + ':00Z'); // hora sobre a época do Excel
const SEM = ['Domingo', 'Segunda', 'Terça', 'Quarta', 'Quinta', 'Sexta', 'Sábado'];

/** Workbook no formato "CONTROLE DE PONTO", salvo e relido (mesmo caminho do app). */
async function planilha() {
  const wb = new ExcelJS.Workbook();
  const abr = wb.addWorksheet('Abril');
  abr.getRow(1).values = [];
  abr.getRow(2).values = ['Dia', 'Dia', 'Entrada', 'Saída', 'Resumo', '', 'Obs'];
  const linhas = [
    ['2026-04-01', '08:00', '18:00'],              // normal: 0 / 0
    ['2026-04-02', '08:00', '12:00'],              // meio período: planilha −360, app −270
    ['2026-04-03'],                                // Sexta-feira Santa vazia: 0 / 0
    ['2026-04-04', '08:00', '12:00'],              // sábado trabalhado: 240 / 240
    ['2026-04-05'],                                // domingo vazio
    ['2026-04-06', 'Feriado'],                     // texto Feriado -> FERIADO
    ['2026-04-07', 'Compensado'],                  // texto Compensado -> FOLGA
    ['2026-04-08'],                                // dia útil vazio: planilha 0, app −510
    ['2026-04-09', '07:50', '18:30', 'médico'],    // obs em G: 40 / 40
    ['2026-04-10', 'Viagem'],                      // texto qualquer -> FOLGA + obs
    // 11..13 não aparecem em linha nenhuma (13 = dia útil: planilha 0, app −510)
    ['2026-04-14', '08:00', '18:00']
  ];
  linhas.forEach((l, i) => {
    const r = abr.getRow(3 + i), k = l[0];
    r.getCell(1).value = SEM[C.diaSemana(k)];
    r.getCell(2).value = DATA(k); r.getCell(2).numFmt = 'dd/mm/yyyy';
    for (const [c, v] of [[3, l[1]], [4, l[2]]]) {
      if (!v) continue;
      if (/^\d\d:\d\d$/.test(v)) { r.getCell(c).value = HORA(v); r.getCell(c).numFmt = 'hh:mm'; } else r.getCell(c).value = v;
    }
    if (l[3]) r.getCell(7).value = l[3];
    r.getCell(6).value = { formula: 'IF(OR(C3="",D3=""),"",1)', result: 0.5 }; // coluna F ignorada
  });
  // aba sem cabeçalho: B/C/D/G; hora como fração do dia e como texto
  const mai = wb.addWorksheet('Maio');
  const r = mai.getRow(5);
  r.getCell(2).value = DATA('2026-05-04'); r.getCell(2).numFmt = 'dd/mm/yyyy';
  r.getCell(3).value = 8 / 24;
  r.getCell(4).value = '18:30';
  wb.addWorksheet('Resumo').getRow(1).values = ['Total', 123];
  const buf = await wb.xlsx.writeBuffer();
  const wb2 = new ExcelJS.Workbook();
  await wb2.xlsx.load(buf);
  return wb2;
}

const AGORA = new Date('2026-05-04T23:00'); // hoje = último dia da planilha
const vazio = (perfil) => C.normalizar({ meses: {}, perfil: perfil || {} });
const dias = (s, k) => C.getDia(s, k);

test('lerPlanilha: cabeçalho com duas colunas Dia, datas/horas UTC, textos, obs, fallback sem cabeçalho, aba sem datas', async () => {
  const { linhas, avisos } = I.lerPlanilha(await planilha());
  const by = Object.fromEntries(linhas.map(l => [l.data, l]));
  assert.equal(linhas.length, 12);
  assert.deepEqual(by['2026-04-01'], { data: '2026-04-01', entrada: '08:00', saida: '18:00', tipo: '', obs: '', saldoPlanilha: 0, aba: 'Abril' });
  assert.equal(by['2026-04-02'].saldoPlanilha, -360);
  assert.equal(by['2026-04-04'].saldoPlanilha, 240);
  assert.equal(by['2026-04-06'].tipo, 'FERIADO');
  assert.equal(by['2026-04-07'].tipo, 'FOLGA');
  assert.equal(by['2026-04-09'].obs, 'médico');
  assert.equal(by['2026-04-09'].saldoPlanilha, 40);
  assert.deepEqual([by['2026-04-10'].tipo, by['2026-04-10'].obs], ['FOLGA', 'Viagem']);
  assert.equal(by['2026-04-03'].entrada, null);
  assert.deepEqual([by['2026-05-04'].entrada, by['2026-05-04'].saida, by['2026-05-04'].saldoPlanilha], ['08:00', '18:30', 30]);
  assert.ok(avisos.some(a => a.includes('Resumo')));
});

test('conversões de célula: Date com erro de ponto flutuante, número, texto', () => {
  assert.equal(I.minutos(new Date(HORA('18:12').getTime() - 1)), 18 * 60 + 12);
  assert.equal(I.minutos(0.5), 720);
  assert.equal(I.minutos('7:05'), 425);
  assert.equal(I.minutos({ formula: 'x', result: HORA('09:00') }), 540);
  assert.equal(I.minutos('Feriado'), null);
  assert.equal(I.data(new Date(DATA('2026-05-01').getTime() - 1)), '2026-05-01');
  assert.equal(I.data(46143), '2026-05-01');
  assert.equal(I.data('01/05/2026'), '2026-05-01');
  assert.equal(I.data('Domingo'), null);
  assert.equal(I.tipoDeTexto('Férias'), 'FÉRIAS');
  assert.equal(I.tipoDeTexto('ATESTADO médico'), 'ATESTADO');
  assert.equal(I.tipoDeTexto('folga'), 'FOLGA');
});

test('importar com ajuste: o C.total do estado APLICADO bate com a planilha, mês a mês', async () => {
  const { linhas } = I.lerPlanilha(await planilha());
  const s = vazio();
  const plano = I.planejar(s, linhas, { agora: AGORA });
  const out = I.aplicar(s, plano);
  assert.deepEqual(s, vazio(), 'aplicar não muta o estado recebido');
  // planilha: abril −360 + 240 + 40 = −80; maio +30
  assert.equal(plano.meses['2026-04'].planilha, -80);
  assert.equal(plano.meses['2026-05'].planilha, 30);
  assert.equal(C.resumo(out, '2026-04', AGORA).saldo, -80);
  assert.equal(C.resumo(out, '2026-05', AGORA).saldo, 30);
  assert.equal(C.total(out, AGORA), -50);
  assert.equal(plano.total.final, -50);
  assert.equal(plano.bancoDepois, -50);
  assert.equal(plano.meses['2026-04'].final, -80);
  // ajustes por dia (cada linha do diff de regras)
  assert.equal(dias(out, '2026-04-01').ajuste, 0);                              // normal
  assert.equal(dias(out, '2026-04-02').ajuste, -90);                            // meio período
  assert.ok(dias(out, '2026-04-02').obs.endsWith(I.SUFIXO));
  assert.equal(dias(out, '2026-04-03'), null);                                  // Sexta Santa: nada a ajustar
  assert.equal(dias(out, '2026-04-04').ajuste, 0);                              // fim de semana trabalhado
  assert.deepEqual([dias(out, '2026-04-06').tipo, dias(out, '2026-04-06').ajuste], ['FERIADO', 0]);
  assert.deepEqual([dias(out, '2026-04-07').tipo, dias(out, '2026-04-07').ajuste], ['FOLGA', 0]);
  assert.deepEqual([dias(out, '2026-04-08').ajuste, dias(out, '2026-04-08').obs], [510, I.SUFIXO]); // dia útil vazio
  assert.equal(dias(out, '2026-04-13').ajuste, 510);                            // data sem linha dentro do intervalo
  assert.equal(dias(out, '2026-04-21'), null);                                  // Tiradentes: 0 / 0
  assert.equal(dias(out, '2026-04-09').obs, 'médico');
  // hora local: 08:00 continua 08:00 em qualquer fuso
  const p = dias(out, '2026-04-01').p;
  assert.deepEqual(p, [L('2026-04-01', '08:00'), L('2026-04-01', '18:00')]);
  assert.equal(new Date(p[0]).getHours(), 8);
  assert.equal(C.horaLocal(dias(out, '2026-05-04').p[1]), '18:30');
});

test('importar sem ajuste: regra pura do app', async () => {
  const { linhas } = I.lerPlanilha(await planilha());
  const s = vazio();
  const plano = I.planejar(s, linhas, { ajustar: false, agora: AGORA });
  const out = I.aplicar(s, plano);
  // abril: −80 da planilha, +90 (meio período), −510 (08), −510 (13), −510 × 11 dias úteis de 15 a 30 (21 é feriado)
  const abril = -80 + 90 - 510 - 510 - 510 * 11;
  assert.equal(C.resumo(out, '2026-04', AGORA).saldo, abril);
  assert.equal(plano.meses['2026-04'].app, abril);
  assert.equal(plano.meses['2026-04'].final, abril);
  assert.equal(C.total(out, AGORA), abril + 30);
  assert.equal(dias(out, '2026-04-02').ajuste, 0);
  assert.equal(dias(out, '2026-04-08'), null);
  // e a coluna "app sem ajuste" do plano com ajuste mostra o mesmo número
  assert.equal(I.planejar(s, linhas, { agora: AGORA }).meses['2026-04'].app, abril);
});

test('conflito: dia já registrado é ignorado por padrão e substituído com sobrescrever', async () => {
  const { linhas } = I.lerPlanilha(await planilha());
  const g = { lat: -23, lon: -46, acc: 5 };
  const s = C.normalizar({ meses: { '2026-04': { dias: { '2026-04-09': { p: [L('2026-04-09', '08:00'), L('2026-04-09', '17:00')], g: [g, null] } } } } });
  const plano = I.planejar(s, linhas, { agora: AGORA });
  assert.equal(plano.meses['2026-04'].ignorados, 1);
  assert.equal(plano.total.ignorados, 1);
  const out = I.aplicar(s, plano);
  assert.deepEqual(dias(out, '2026-04-09').p, s.meses['2026-04'].dias['2026-04-09'].p);
  assert.deepEqual(dias(out, '2026-04-09').g, [g, null]);
  // abril = planilha (−80) − planilha do dia 09 (40) + saldo do app no dia 09 (540 − 90 − 510 = −60)
  assert.equal(C.resumo(out, '2026-04', AGORA).saldo, -180);
  assert.equal(plano.meses['2026-04'].final, -180);

  const plano2 = I.planejar(s, linhas, { sobrescrever: true, agora: AGORA });
  assert.equal(plano2.meses['2026-04'].ignorados, 0);
  const out2 = I.aplicar(s, plano2);
  assert.deepEqual(dias(out2, '2026-04-09').p, [L('2026-04-09', '07:50'), L('2026-04-09', '18:30')]);
  assert.deepEqual(dias(out2, '2026-04-09').g, [null, null]);
  assert.equal(C.total(out2, AGORA), -50);
  // dia só com tipo/obs também é preservado (não é sobrescrito sem pedir)
  const s3 = C.normalizar({ meses: { '2026-04': { dias: { '2026-04-08': { p: [], tipo: 'FÉRIAS' } } } } });
  const out3 = I.aplicar(s3, I.planejar(s3, linhas, { agora: AGORA }));
  assert.equal(dias(out3, '2026-04-08').tipo, 'FÉRIAS');
});

test('zerarSaldoIni: aviso de duplicidade, padrão zera, opção mantém', async () => {
  const { linhas } = I.lerPlanilha(await planilha());
  const s = vazio({ saldoIni: 367, nome: 'X' });
  const plano = I.planejar(s, linhas, { agora: AGORA });
  assert.equal(plano.avisoSaldoIni, true);
  assert.equal(plano.opcoes.zerarSaldoIni, true);
  assert.ok(plano.avisos.some(a => a.includes('duplicaria')));
  assert.equal(plano.bancoAntes, 367);
  const out = I.aplicar(s, plano);
  assert.equal(out.perfil.saldoIni, 0);
  assert.equal(out.perfil.nome, 'X');
  assert.equal(C.total(out, AGORA), -50);
  const out2 = I.aplicar(s, I.planejar(s, linhas, { zerarSaldoIni: false, agora: AGORA }));
  assert.equal(out2.perfil.saldoIni, 367);
  assert.equal(C.total(out2, AGORA), 317);
  assert.equal(I.planejar(vazio(), linhas, { agora: AGORA }).avisoSaldoIni, false);
});

test('aviso de dias úteis vazios entre o fim da planilha e hoje', async () => {
  const { linhas } = I.lerPlanilha(await planilha());
  const plano = I.planejar(vazio(), linhas, { agora: new Date('2026-05-06T12:00') });
  assert.ok(plano.avisos.some(a => a.startsWith('2 dia(s)')), plano.avisos.join(' | '));
  assert.equal(plano.bancoDepois, -50 - 1020);
});

test('planilha sem dados', () => {
  const plano = I.planejar(vazio(), [], {});
  assert.equal(plano.itens.length, 0);
  assert.equal(plano.avisos.length, 1);
});
