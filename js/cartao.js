/* Ponto VIC — cartão de ponto mensal em Excel (modelo original "Cartão de Ponto").
 * Node: module.exports · Navegador: window.PontoCartao (depende de window.PontoCore)
 * Recebe o construtor ExcelJS por parâmetro: o bundle é carregado sob demanda. */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory(require('./core'));
  else root.PontoCartao = factory(root.PontoCore);
})(typeof self !== 'undefined' ? self : globalThis, function (C) {
  'use strict';

  const B = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
  const CENTRO = { vertical: 'middle', horizontal: 'center' };
  const COLS = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H'];
  const VAZIO = '_____ : _____ ';

  const nomeArquivo = m => 'PONTO VIC ' + m.slice(5, 7) + m.slice(0, 4) + '.xlsx';

  /** Monta o workbook do mês m ('YYYY-MM'). */
  function montar(ExcelJS, estado, m, agora) {
    agora = agora ? new Date(agora) : new Date();
    const s = estado, y = Number(m.slice(0, 4)), mo = Number(m.slice(5, 7)) - 1, n = C.diasNoMes(m);
    const r = C.resumo(s, m, agora), ini = C.inicio(s);
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Plan1', { pageSetup: { paperSize: 9, orientation: 'portrait', fitToPage: true, fitToWidth: 1, fitToHeight: 1 } });
    ws.columns = [{ width: 14 }, { width: 18 }, { width: 18 }, { width: 18 }, { width: 18 }, { width: 26 }, { width: 30 }, { width: 40 }];

    const faixa = (row, txt, opt) => {
      ws.mergeCells('A' + row + ':H' + row);
      const c = ws.getCell('A' + row);
      c.value = txt;
      Object.assign(c, opt || {});
      COLS.forEach(l => { ws.getCell(l + row).border = B; });
    };
    faixa(1, 'Cartão de Ponto', { font: { bold: true, size: 14 }, alignment: CENTRO });
    faixa(2, '');
    faixa(3, ' Horario de intervalo para refeições livre limitado a 01:30 hora', { font: { bold: true } });
    faixa(4, 'MÊS: ' + C.MESES[mo] + '/' + y, { font: { bold: true } });
    ws.getCell('A5').value = 'NOME:';
    ws.getCell('A5').font = { bold: true };
    ws.mergeCells('B5:H5');
    ws.getCell('B5').value = (s.perfil && s.perfil.nome) || '';
    COLS.forEach(l => { ws.getCell(l + '5').border = B; });

    ws.getRow(6).values = ['Data ', 'Horário Entrada', 'Saída Almoço', 'Volta Almoço', 'Horário Saída', 'Rubrica Funcionário', 'Localização (GPS)', 'Observação'];
    ws.getRow(6).eachCell(c => {
      c.font = { bold: true }; c.alignment = CENTRO; c.border = B;
      c.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD9D9D9' } };
    });

    let ajustes = 0;
    for (let d = 1; d <= n; d++) {
      const k = m + '-' + C.pad(d), x = C.getDia(s, k) || C.diaVazio(), w = C.diaSemana(k), row = ws.getRow(6 + d);
      let e = VAZIO, as = VAZIO, av = VAZIO, sa = VAZIO;
      const hd = iso => { const off = C.horariosDoDia(k, [iso])[0].dias; return C.horaLocal(iso) + (off > 0 ? ' (+' + off + 'd)' : ''); };
      const q = x.p.length;
      if (q) {
        e = C.horaLocal(x.p[0]);
        if (q >= 3) { as = hd(x.p[1]); av = hd(x.p[2]); }  // almoço batido: 2ª e 3ª marcações
        else if (q === 2 && x.almoco) as = hd(x.p[1]);    // saiu para o almoço e não voltou (ainda)
        else if (q === 2) as = av = '—';                  // só entrada e saída: almoço descontado sozinho (obs)
        if ((q === 2 && !x.almoco) || q >= 4) sa = hd(x.p[q - 1]);
      } else if (x.tipo) e = as = av = sa = x.tipo;
      else if (C.feriadoAuto(s, k)) e = as = av = sa = 'FERIADO';
      else if (w === 6) e = as = av = sa = 'SÁBADO';
      else if (w === 0) e = as = av = sa = 'DOMINGO';
      const gs = (x.g || []).filter(Boolean);
      const obs = [];
      if (!x.p.length && !x.tipo && C.feriadoAuto(s, k)) obs.push(C.nomeFeriado(k));
      const auto = C.almocoAuto(s, k, agora);
      if (auto) obs.push('Almoço de ' + C.formatHM(auto) + ' descontado (sem marcação)');
      if (x.obs) obs.push(x.obs);
      if (x.ajuste) obs.push('Ajuste ' + C.formatHMsinal(x.ajuste));
      if (x.banco) obs.push('Pagando banco');
      if (x.ajuste && C.contaNoSaldo(s, k, agora, ini)) ajustes += x.ajuste;
      row.values = [
        new Date(Date.UTC(y, mo, d)), e, as, av, sa, '',
        gs.length ? { text: 'Ver no mapa (' + gs.length + ' marcações)', hyperlink: 'https://maps.google.com/?q=' + gs[0].lat + ',' + gs[0].lon } : '',
        obs.join(' · ')
      ];
      row.getCell(7).font = { color: { argb: 'FFF07D00' }, underline: !!gs.length };
      row.getCell(1).numFmt = 'dd/mm/yyyy';
      row.eachCell({ includeEmpty: true }, c => { c.border = B; c.alignment = CENTRO; });
      row.getCell(8).alignment = { vertical: 'middle', horizontal: 'left', wrapText: true };
    }

    const fim = 7 + n;
    faixa(fim, 'ASSINATURA:', { font: { bold: true } });
    ws.getRow(fim).height = 30;
    const hm = C.formatHM;
    const res = [
      ['Horas trabalhadas', hm(r.tr)], ['Horas previstas', hm(r.pv)], ['Horas extras', hm(r.ex)], ['Horas devidas', hm(r.dv)],
      ['Ajustes do mês (já incluídos)', hm(ajustes)], ['Saldo do mês', hm(r.saldo)], ['Banco de horas total', hm(C.total(s, agora))]
    ];
    res.forEach((v, i) => {
      const lin = fim + 2 + i, row = ws.getRow(lin);
      row.getCell(1).value = v[0];
      ws.mergeCells('A' + lin + ':B' + lin);
      row.getCell(3).value = v[1];
      row.getCell(1).font = { bold: true };
      [1, 2, 3].forEach(c => { row.getCell(c).border = B; });
      row.getCell(3).alignment = CENTRO;
    });
    return wb;
  }

  return { montar, nomeArquivo };
});
