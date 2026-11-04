/* Ponto VIC — importação da planilha antiga "CONTROLE DE PONTO" (puro, sem DOM).
 * Node: module.exports · Navegador: window.PontoImport (depende de window.PontoCore) */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory(require('./core'));
  else root.PontoImport = factory(root.PontoCore);
})(typeof self !== 'undefined' ? self : globalThis, function (C) {
  'use strict';

  const pad = C.pad;
  const EPOCA = Date.UTC(1899, 11, 30);
  const SUFIXO = '[ajuste p/ planilha]';
  const semAcento = s => String(s == null ? '' : s).normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().trim();

  /** Valor "cru" da célula do ExcelJS (resolve fórmula, rich text e hyperlink). */
  function valor(v) {
    if (v == null) return null;
    if (v instanceof Date || typeof v !== 'object') return v;
    if (Array.isArray(v.richText)) return v.richText.map(t => t.text).join('');
    if ('result' in v) return valor(v.result);
    if ('text' in v) return valor(v.text);
    return null;
  }
  const texto = v => { v = valor(v); return v == null || v instanceof Date ? '' : String(v).trim(); };

  /** Célula de hora -> minutos do dia (Date UTC sobre 1899-12-30, fração do dia, ou texto HH:MM). */
  function minutos(v) {
    v = valor(v);
    if (v instanceof Date) { const n = Math.round((v.getTime() - EPOCA) / 6e4); return ((n % 1440) + 1440) % 1440; }
    if (typeof v === 'number' && isFinite(v) && v >= 0) return Math.round(v * 1440) % 1440;
    if (typeof v === 'string') {
      const m = v.trim().match(/^(\d{1,2})\s*[:h]\s*(\d{2})(?::\d{2})?$/i);
      if (m && Number(m[1]) < 24 && Number(m[2]) < 60) return Number(m[1]) * 60 + Number(m[2]);
    }
    return null;
  }
  const hhmm = n => (n == null ? null : pad(Math.floor(n / 60)) + ':' + pad(n % 60));

  function chaveValida(y, m, d) {
    const dt = new Date(Date.UTC(y, m - 1, d));
    return y >= 1950 && y <= 2200 && dt.getUTCMonth() === m - 1 && dt.getUTCDate() === d ? y + '-' + pad(m) + '-' + pad(d) : null;
  }
  /** Célula de data -> 'YYYY-MM-DD' (Date UTC, número serial do Excel ou texto). */
  function data(v) {
    v = valor(v);
    if (v instanceof Date) { const d = new Date(Math.round(v.getTime() / 6e4) * 6e4); return chaveValida(d.getUTCFullYear(), d.getUTCMonth() + 1, d.getUTCDate()); }
    if (typeof v === 'number' && v > 18000 && v < 120000) { const d = new Date(EPOCA + Math.round(v) * 864e5); return chaveValida(d.getUTCFullYear(), d.getUTCMonth() + 1, d.getUTCDate()); }
    if (typeof v === 'string') {
      const s = v.trim();
      let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
      if (m) return chaveValida(+m[1], +m[2], +m[3]);
      m = s.match(/^(\d{1,2})[/.-](\d{1,2})[/.-](\d{2}|\d{4})$/);
      if (m) return chaveValida(m[3].length === 2 ? 2000 + +m[3] : +m[3], +m[2], +m[1]);
    }
    return null;
  }

  /** Texto em entrada/saída -> tipo do dia ('' se não reconhecido). */
  function tipoDeTexto(t) {
    t = semAcento(t);
    if (!t) return '';
    if (t.includes('feriado')) return 'FERIADO';
    if (t.includes('compensado') || t.includes('folga')) return 'FOLGA';
    if (t.includes('ferias')) return 'FÉRIAS';
    if (t.includes('atestado')) return 'ATESTADO';
    return '';
  }

  /** Acha cabeçalho e colunas de uma aba. */
  function colunas(ws) {
    const max = Math.min(ws.rowCount || 0, 6);
    for (let r = 1; r <= max; r++) {
      const vals = Array.from(ws.getRow(r).values || [], v => semAcento(texto(v))); // values é esparso
      const ent = vals.indexOf('entrada'), sai = vals.indexOf('saida');
      if (ent < 0 || sai < 0) continue;
      const cands = [];
      vals.forEach((v, i) => { if (v === 'dia' || v === 'data') cands.push(i); });
      let col = cands.length ? cands[0] : 2, melhor = -1;
      for (const c of cands) { // escolhe a coluna "Dia" que de fato tem datas
        let n = 0;
        for (let rr = r + 1; rr <= Math.min(ws.rowCount, r + 40); rr++) if (data(ws.getRow(rr).values[c])) n++;
        if (n > melhor) { melhor = n; col = c; }
      }
      let obs = vals.findIndex(v => v.startsWith('obs'));
      if (obs < 0 && !vals[7]) obs = 7; // sem cabeçalho "Obs": observação livre fica em G
      return { cab: r, data: col, ent, sai, obs: obs < 0 ? null : obs };
    }
    return { cab: 0, data: 2, ent: 3, sai: 4, obs: 7 };
  }

  /**
   * Lê o workbook ExcelJS.
   * @returns {{linhas:Array<{data,entrada,saida,tipo,obs,saldoPlanilha,aba}>, avisos:string[]}}
   */
  function lerPlanilha(wb) {
    const porData = new Map(), avisos = [];
    wb.eachSheet(ws => {
      const col = colunas(ws);
      let achou = 0;
      ws.eachRow((row, r) => {
        if (r <= col.cab) return;
        const v = row.values || [];
        const k = data(v[col.data]);
        if (!k) return;
        achou++;
        const ce = v[col.ent], cs = v[col.sai];
        const me = minutos(ce), ms = minutos(cs);
        const textos = [ce, cs].filter((c, i) => (i ? ms : me) == null).map(texto).filter(Boolean);
        let tipo = '';
        const obs = [];
        for (const t of textos) {
          const ti = tipoDeTexto(t);
          if (ti && !tipo) tipo = ti;
          else if (!ti) { if (!tipo) tipo = 'FOLGA'; obs.push(t); }
        }
        const livre = col.obs != null ? texto(v[col.obs]) : '';
        if (livre) obs.push(livre);
        // fórmula da planilha: entrada e saída preenchidas ? (fim de semana ? saída−entrada : saída−entrada−10:00) : 0
        let saldo = 0;
        if (me != null && ms != null) saldo = ms - me - (C.util(k) ? 600 : 0);
        if (porData.has(k)) avisos.push('Data ' + k + ' repetida (aba ' + ws.name + '); vale a última.');
        porData.set(k, { data: k, entrada: hhmm(me), saida: hhmm(ms), tipo, obs: obs.join(' · '), saldoPlanilha: saldo, aba: ws.name });
      });
      if (!achou) avisos.push('Aba "' + ws.name + '" ignorada: sem datas.');
    });
    const linhas = [...porData.values()].sort((a, b) => (a.data < b.data ? -1 : a.data > b.data ? 1 : 0));
    return { linhas, avisos };
  }

  const linhaTemDado = l => !!(l.entrada || l.saida || l.tipo || l.obs);
  const linhaVazia = k => ({ data: k, entrada: null, saida: null, tipo: '', obs: '', saldoPlanilha: 0, aba: '' });

  function aplicarItens(s, itens, campo, zerar) {
    const out = C.normalizar(s);
    for (const it of itens) {
      if (it.acao === 'ignorar') continue;
      const m = it.k.slice(0, 7), d = it[campo];
      if (C.temDados(d)) {
        out.meses[m] = out.meses[m] || { dias: {} };
        out.meses[m].dias[it.k] = C.clone(d);
      } else if (out.meses[m] && out.meses[m].dias[it.k]) {
        delete out.meses[m].dias[it.k];
      }
    }
    if (zerar) out.perfil.saldoIni = 0;
    return out;
  }

  /**
   * Plano de importação (puro).
   * opcoes: { ajustar=true, sobrescrever=false, zerarSaldoIni=(saldoIni≠0), agora=new Date() }
   * Conflito = data que já tem dados no app (marcações, tipo, obs, ajuste ou "pagando banco"): ignorada, salvo `sobrescrever`
   * (que preserva o "pagando banco" local).
   * Ajuste por dia = saldoPlanilha − saldoDia(estado com o dia NOVO sem ajuste).
   */
  function planejar(estado, linhas, opcoes) {
    const s = C.normalizar(estado);
    const o = Object.assign({ ajustar: true, sobrescrever: false, zerarSaldoIni: s.perfil.saldoIni !== 0 }, opcoes || {});
    const agora = o.agora ? new Date(o.agora) : new Date();
    o.agora = agora.toISOString();
    const hoje = C.keyDate(agora);
    const avisos = [];
    const comDado = (linhas || []).filter(l => l && l.data && linhaTemDado(l)).map(l => l.data).sort();
    const vazio = { itens: [], meses: {}, total: { dias: 0, ignorados: 0, planilha: 0, app: 0, final: 0 }, de: null, ate: null,
      opcoes: o, avisos: ['Nenhuma data com dados para importar.'], avisoSaldoIni: false, bancoAntes: C.total(s, agora), bancoDepois: C.total(s, agora) };
    if (!comDado.length) return vazio;
    const de = comDado[0], ate = comDado[comDado.length - 1];
    const porData = new Map((linhas || []).filter(l => l && l.data).map(l => [l.data, l]));

    const itens = [];
    for (const k of C.diasEntre(de, ate)) {
      const row = porData.get(k) || linhaVazia(k);
      const antigo = C.getDia(s, k);
      const conflito = C.temDados(antigo);
      if (conflito && !o.sobrescrever) { itens.push({ k, acao: 'ignorar', row }); continue; }
      const mont = C.montarMarcacoes(k, [row.entrada, row.saida].filter(Boolean));
      // "pagando banco" é escolha local que a planilha não representa: ao sobrescrever, é preservado
      const novo = { p: mont.p, g: mont.g, banco: !!(antigo && antigo.banco === true), tipo: row.tipo || '', obs: row.obs || '', ajuste: 0 };
      const saldoApp = C.saldoDia(C.comDia(s, k, novo), k, agora); // estado COM o dia novo, sem ajuste
      const d = C.clone(novo);
      const ajustavel = k < hoje || (k === hoje && d.p.length % 2 === 0);
      if (o.ajustar && ajustavel) {
        d.ajuste = row.saldoPlanilha - saldoApp;
        if (d.ajuste) d.obs = (d.obs ? d.obs + ' ' : '') + SUFIXO;
      }
      itens.push({ k, acao: conflito ? 'sobrescrever' : 'importar', row, d, dSemAjuste: novo, saldoApp, futuro: k > hoje });
    }

    const aplicado = aplicarItens(s, itens, 'd', o.zerarSaldoIni);
    const semAjuste = aplicarItens(s, itens, 'dSemAjuste', o.zerarSaldoIni);
    const meses = {};
    for (const it of itens) {
      const m = it.k.slice(0, 7);
      const T = meses[m] || (meses[m] = { dias: 0, ignorados: 0, planilha: 0, app: 0, final: 0 });
      if (it.acao === 'ignorar') T.ignorados++; else T.dias++;
      if (it.k <= hoje) T.planilha += it.row.saldoPlanilha;
    }
    const total = { dias: 0, ignorados: 0, planilha: 0, app: 0, final: 0 };
    for (const m of Object.keys(meses)) {
      meses[m].app = C.resumo(semAjuste, m, agora).saldo;
      meses[m].final = C.resumo(aplicado, m, agora).saldo;
      for (const q of Object.keys(total)) total[q] += meses[m][q];
    }

    const ign = itens.filter(i => i.acao === 'ignorar').length, sob = itens.filter(i => i.acao === 'sobrescrever').length;
    if (ign) avisos.push(ign + ' dia(s) já registrados no app foram mantidos como estão (ignorados).');
    if (sob) avisos.push(sob + ' dia(s) já registrados no app serão substituídos pela planilha (o GPS deles se perde).');
    const fut = itens.filter(i => i.futuro && i.acao !== 'ignorar').length;
    if (fut) avisos.push(fut + ' dia(s) depois de hoje importados sem ajuste (ainda não contam no saldo).');
    const avisoSaldoIni = s.perfil.saldoIni !== 0;
    if (avisoSaldoIni) avisos.push('Você tem um saldo anterior de ' + C.formatHM(s.perfil.saldoIni) + '. Ele provavelmente já é o saldo desta planilha; somar os dois duplicaria as horas.');
    // dias úteis vazios depois da planilha que passam a contar como devidos
    let novos = 0, min = 0;
    const iniAntes = C.inicio(s), iniDepois = C.inicio(aplicado);
    for (let k = C.somaDias(ate, 1); k <= hoje; k = C.somaDias(k, 1)) {
      const conta = C.contaNoSaldo(aplicado, k, agora, iniDepois), contava = C.contaNoSaldo(s, k, agora, iniAntes);
      if (conta && !contava) { const sd = C.saldoDia(aplicado, k, agora); if (sd) { novos++; min += sd; } }
    }
    if (novos) avisos.push(novos + ' dia(s) entre ' + C.somaDias(ate, 1) + ' e hoje, sem registro, passam a contar no saldo (' + C.formatHM(min) + ').');

    return { de, ate, itens, meses, total, opcoes: o, avisos, avisoSaldoIni, saldoIni: s.perfil.saldoIni,
      bancoAntes: C.total(s, agora), bancoDepois: C.total(aplicado, agora) };
  }

  /** Aplica o plano e devolve um estado NOVO (não muta o recebido). */
  function aplicar(estado, plano) {
    return aplicarItens(estado, (plano && plano.itens) || [], 'd', !!(plano && plano.opcoes && plano.opcoes.zerarSaldoIni));
  }

  return { lerPlanilha, planejar, aplicar, minutos, data, tipoDeTexto, SUFIXO };
});
