/* Ponto VIC — regras de cálculo (puras, sem DOM).
 * Node: module.exports · Navegador: window.PontoCore
 * Datas de dia são chaves 'YYYY-MM-DD' no fuso local; marcações são ISO absolutas. */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.PontoCore = factory();
})(typeof self !== 'undefined' ? self : globalThis, function () {
  'use strict';

  const MESES = ['JANEIRO', 'FEVEREIRO', 'MARÇO', 'ABRIL', 'MAIO', 'JUNHO', 'JULHO', 'AGOSTO', 'SETEMBRO', 'OUTUBRO', 'NOVEMBRO', 'DEZEMBRO'];
  const TIPOS = ['FERIADO', 'FOLGA', 'FÉRIAS', 'ATESTADO'];

  const pad = n => String(n).padStart(2, '0');
  /** Date -> 'YYYY-MM-DD' no fuso local. */
  const keyDate = d => d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate());
  const keyMes = d => keyDate(d).slice(0, 7);
  /** 'YYYY-MM-DD' -> Date ao meio-dia local (seguro contra horário de verão). */
  const dateKey = k => { const [y, m, d] = k.split('-').map(Number); return new Date(y, m - 1, d, 12); };
  /** Dia da semana (0 = domingo) a partir da chave, independente do fuso. */
  const diaSemana = k => { const [y, m, d] = k.split('-').map(Number); return new Date(Date.UTC(y, m - 1, d)).getUTCDay(); };
  const somaDias = (k, n) => { const d = dateKey(k); d.setDate(d.getDate() + n); return keyDate(d); };
  const diasNoMes = m => { const [y, mm] = m.split('-').map(Number); return new Date(y, mm, 0).getDate(); };
  const clone = x => JSON.parse(JSON.stringify(x));
  /** Dia útil (seg–sex). Aceita Date ou chave. */
  const util = d => { const w = typeof d === 'string' ? diaSemana(d) : d.getDay(); return w > 0 && w < 6; };
  /** 'HH:MM' local de um ISO. */
  const horaLocal = iso => { const d = new Date(iso); return pad(d.getHours()) + ':' + pad(d.getMinutes()); };

  // ---------- normalização / migração ----------
  const diaVazio = () => ({ p: [], g: [], banco: false, tipo: '', obs: '', ajuste: 0 });

  const num = v => typeof v === 'number' && Number.isFinite(v);
  /** GPS válido: lat ∈ [-90,90], lon ∈ [-180,180], acc ≥ 0 (todos números finitos). */
  const gpsValido = g => !!g && typeof g === 'object' && num(g.lat) && num(g.lon) && num(g.acc) &&
    g.lat >= -90 && g.lat <= 90 && g.lon >= -180 && g.lon <= 180 && g.acc >= 0;
  /** Só números; qualquer outra coisa vira null (evita injeção pelo backup ou pela nuvem). */
  const limparGPS = g => (gpsValido(g) ? { lat: g.lat, lon: g.lon, acc: g.acc } : null);
  const isoValido = v => typeof v === 'string' && /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}/.test(v) && !isNaN(new Date(v));

  function normalizarDia(x) {
    const d = Object.assign({}, x && typeof x === 'object' ? x : {});
    d.p = Array.isArray(d.p) ? d.p.filter(v => typeof v === 'string' && !isNaN(new Date(v))) : [];
    const g = Array.isArray(d.g) ? d.g : [];
    d.g = d.p.map((_, i) => limparGPS(g[i]));
    d.banco = !!d.banco;
    d.tipo = typeof d.tipo === 'string' ? d.tipo : '';
    d.obs = typeof d.obs === 'string' ? d.obs : '';
    d.ajuste = Number.isFinite(Number(d.ajuste)) ? Math.round(Number(d.ajuste)) : 0;
    return d;
  }

  /** Cópia completa e normalizada. Dados antigos sem obs/ajuste/g/feriadosAuto carregam normalmente; campos desconhecidos são mantidos. */
  function normalizar(input) {
    const s = clone(input && typeof input === 'object' ? input : {});
    s.meses = s.meses && typeof s.meses === 'object' ? s.meses : {};
    for (const m of Object.keys(s.meses)) {
      const M = s.meses[m] && typeof s.meses[m] === 'object' ? s.meses[m] : {};
      const dias = M.dias && typeof M.dias === 'object' ? M.dias : {};
      for (const k of Object.keys(dias)) dias[k] = normalizarDia(dias[k]);
      M.dias = dias;
      s.meses[m] = M;
    }
    const p = s.perfil && typeof s.perfil === 'object' ? s.perfil : {};
    s.perfil = Object.assign({}, p, {
      nome: typeof p.nome === 'string' ? p.nome : '',
      saldoIni: Number.isFinite(Number(p.saldoIni)) ? Math.round(Number(p.saldoIni)) : 0,
      feriadosAuto: p.feriadosAuto !== false,
      ultimoBackup: typeof p.ultimoBackup === 'string' ? p.ultimoBackup : null
    });
    return s;
  }

  /** Lê o dia sem criar nada. */
  function getDia(s, k) {
    const M = s.meses[k.slice(0, 7)];
    return (M && M.dias && M.dias[k]) || null;
  }
  /** Lê ou cria o dia (muta o estado). */
  function dia(s, k) {
    const m = k.slice(0, 7);
    const M = s.meses[m] || (s.meses[m] = { dias: {} });
    M.dias = M.dias || {};
    return M.dias[k] || (M.dias[k] = diaVazio());
  }
  /** Novo estado (cópia rasa do caminho) com o dia k substituído por d. Não muta s. */
  function comDia(s, k, d) {
    const m = k.slice(0, 7), M = s.meses[m] || { dias: {} };
    const dias = Object.assign({}, M.dias, { [k]: d });
    return Object.assign({}, s, { meses: Object.assign({}, s.meses, { [m]: Object.assign({}, M, { dias }) }) });
  }
  /** Registro que marca o início do banco: marcação, tipo, obs ou ajuste ≠ 0 (regra da spec). */
  const temRegistro = x => !!x && ((x.p && x.p.length > 0) || !!x.tipo || !!x.obs || (Number(x.ajuste) || 0) !== 0);
  /** Qualquer dado do usuário no dia, inclusive "pagando banco" (usado para conflito e para não apagar o dia). */
  const temDados = x => temRegistro(x) || (!!x && x.banco === true);

  // ---------- feriados nacionais ----------
  /** Domingo de Páscoa ('YYYY-MM-DD'). */
  function pascoa(y) {
    const a = y % 19, b = Math.floor(y / 100), c = y % 100, d = Math.floor(b / 4), e = b % 4, f = Math.floor((b + 8) / 25),
      g = Math.floor((b - f + 1) / 3), h = (19 * a + b - d - g + 15) % 30, i = Math.floor(c / 4), j = c % 4,
      l = (32 + 2 * e + 2 * i - h - j) % 7, m = Math.floor((a + 11 * h + 22 * l) / 451),
      mes = Math.floor((h + l - 7 * m + 114) / 31), di = ((h + l - 7 * m + 114) % 31) + 1;
    return y + '-' + pad(mes) + '-' + pad(di);
  }
  const FIXOS = {
    '01-01': 'Confraternização Universal', '04-21': 'Tiradentes', '05-01': 'Dia do Trabalho', '09-07': 'Independência',
    '10-12': 'Nossa Senhora Aparecida', '11-02': 'Finados', '11-15': 'Proclamação da República', '11-20': 'Consciência Negra', '12-25': 'Natal'
  };
  /** Nome do feriado nacional da data, ou ''. */
  function nomeFeriado(k) {
    if (FIXOS[k.slice(5)]) return FIXOS[k.slice(5)];
    if (somaDias(pascoa(Number(k.slice(0, 4))), -2) === k) return 'Sexta-feira Santa';
    return '';
  }
  const feriado = k => nomeFeriado(k) !== '';
  /** Feriado automático: opção ligada, data é feriado nacional e o dia não tem marcações nem tipo. */
  function feriadoAuto(s, k) {
    if (s.perfil && s.perfil.feriadosAuto === false) return false;
    const x = getDia(s, k);
    if (x && ((x.p && x.p.length) || x.tipo)) return false;
    return feriado(k);
  }

  // ---------- cálculo ----------
  function previsto(s, k) {
    const x = getDia(s, k) || {};
    if (!util(k) || x.tipo || feriadoAuto(s, k)) return 0;
    return x.banco ? 540 : 510;
  }

  const LIMITE_ABERTO = 16 * 60; // minutos: marcação aberta mais antiga que isso não é continuada

  /** Último dia antes de `antes` (chave) com alguma marcação, ou null. */
  function ultimoDiaComMarcacao(s, antes) {
    let u = null;
    for (const m of Object.keys(s.meses)) {
      if (m > antes.slice(0, 7)) continue;
      const dias = (s.meses[m] && s.meses[m].dias) || {};
      for (const k of Object.keys(dias)) if (k < antes && dias[k] && dias[k].p && dias[k].p.length && (!u || k > u)) u = k;
    }
    return u;
  }
  const minDesde = (iso, agora) => (new Date(agora) - new Date(iso)) / 6e4;

  /**
   * Dia com ponto aberto agora: hoje (se tem marcação ímpar) ou, se hoje não tem marcação,
   * o último dia anterior com marcação ímpar cuja última batida foi há menos de 16 h.
   */
  function diaAberto(s, agora) {
    const now = agora ? new Date(agora) : new Date(), hoje = keyDate(now), x = getDia(s, hoje);
    if (x && x.p && x.p.length) return x.p.length % 2 ? hoje : null;   // hoje tem marcações: só hoje pode estar aberto
    const u = ultimoDiaImpar(s, hoje);
    if (!u) return null;
    const y = getDia(s, u), min = minDesde(y.p[y.p.length - 1], now);
    return min >= 0 && min < LIMITE_ABERTO ? u : null;
  }

  /** Último dia antes de `antes` (chave) com número ímpar de marcações, ou null. */
  function ultimoDiaImpar(s, antes) {
    let u = null;
    for (const m of Object.keys(s.meses)) {
      if (m > antes.slice(0, 7)) continue;
      const dias = (s.meses[m] && s.meses[m].dias) || {};
      for (const k of Object.keys(dias)) if (k < antes && dias[k] && dias[k].p && dias[k].p.length % 2 === 1 && (!u || k > u)) u = k;
    }
    return u;
  }

  /**
   * Todos os dias anteriores a hoje com número ímpar de marcações que não são o dia aberto atual
   * (ponto incompleto: ficam fora do saldo até serem corrigidos), do mais recente para o mais antigo.
   * @returns {Array<{k:string, desde:string}>} desde = última marcação do dia
   */
  function pendentes(s, agora) {
    const now = agora ? new Date(agora) : new Date(), hoje = keyDate(now), ab = diaAberto(s, now), out = [];
    for (const m of Object.keys(s.meses)) {
      if (m > hoje.slice(0, 7)) continue;
      const dias = (s.meses[m] && s.meses[m].dias) || {};
      for (const k of Object.keys(dias)) {
        const x = dias[k];
        if (k < hoje && k !== ab && x && x.p && x.p.length % 2 === 1) out.push({ k, desde: x.p[x.p.length - 1] });
      }
    }
    return out.sort((a, b) => (a.k < b.k ? 1 : a.k > b.k ? -1 : 0));
  }

  /**
   * Para onde vai a próxima batida.
   * @returns {{k:string, acao:'entrada'|'saida', anterior:boolean, desde:string|null, pendente:{k:string,desde:string}|null}}
   *  anterior = a batida fecha um dia anterior (virada da meia-noite);
   *  pendente = último dia anterior com marcação ímpar que NÃO será fechado pela próxima batida
   *  (16 h ou mais, ou hoje já tem marcações): precisa de correção manual e fica fora do saldo.
   */
  function destinoBatida(s, agora) {
    const now = agora ? new Date(agora) : new Date(), hoje = keyDate(now), ab = diaAberto(s, now);
    const u = ultimoDiaImpar(s, hoje);
    const pendente = u && u !== ab ? { k: u, desde: getDia(s, u).p[getDia(s, u).p.length - 1] } : null;
    if (ab) { const x = getDia(s, ab); return { k: ab, acao: 'saida', anterior: ab !== hoje, desde: x.p[x.p.length - 1], pendente }; }
    return { k: hoje, acao: 'entrada', anterior: false, desde: null, pendente };
  }

  /**
   * Decide se um toque em ENTRADA/SAÍDA deve virar marcação (item toque duplo).
   * @param acao 'entrada'|'saida' (rótulo do botão tocado)
   * @param ultimoToque ms (Date.now) da última batida aceita, ou null
   */
  function decidirBatida(s, agora, acao, ultimoToque) {
    const now = agora ? new Date(agora) : new Date(), destino = destinoBatida(s, now);
    if (ultimoToque != null && now.getTime() - ultimoToque < 1500) return { ok: false, motivo: 'repetido', destino };
    if (acao !== destino.acao) return { ok: false, motivo: 'paridade', destino };
    return { ok: true, motivo: '', destino };
  }

  /** Remove a marcação `iso` se ela ainda for a última do dia k (desfazer). Muta s. */
  function removerUltima(s, k, iso) {
    const x = getDia(s, k);
    if (!x || x.p[x.p.length - 1] !== iso) return false;
    x.p.pop(); x.g.length = x.p.length;
    if (!temDados(x)) delete s.meses[k.slice(0, 7)].dias[k];
    return true;
  }

  /** Minutos trabalhados. Par aberto conta até `agora` só no dia aberto atual (hoje ou virada < 16 h). */
  function bruto(s, k, agora) {
    const x = getDia(s, k);
    if (!x || !x.p || !x.p.length) return { x, t: 0, contaAberto: false };
    const now = agora ? new Date(agora) : null, contaAberto = !!now && x.p.length % 2 === 1 && diaAberto(s, now) === k;
    let t = 0;
    for (let i = 0; i < x.p.length; i += 2) {
      const a = new Date(x.p[i]);
      const b = x.p[i + 1] ? new Date(x.p[i + 1]) : (contaAberto ? now : null);
      if (b) t += (b - a) / 6e4;
    }
    return { x, t, contaAberto };
  }

  /** Almoço descontado automaticamente (min): só entrada/saída, ou só a entrada com o dia aberto ("se sair agora"). */
  function almocoAuto(s, k, agora) {
    const { x, t, contaAberto } = bruto(s, k, agora);
    if (!x || !x.p) return 0;
    const soEntradaSaida = x.p.length === 2 || (x.p.length === 1 && contaAberto);
    return soEntradaSaida && util(k) && !x.tipo && t > 360 ? (x.banco ? 60 : 90) : 0;
  }

  function trabalhado(s, k, agora) {
    const { t } = bruto(s, k, agora);
    return Math.max(0, Math.round(t - almocoAuto(s, k, agora)));
  }

  /** saldo do dia = trabalhado − previsto + ajuste */
  function saldoDia(s, k, agora) {
    const x = getDia(s, k) || {};
    return trabalhado(s, k, agora) - previsto(s, k) + (Number(x.ajuste) || 0);
  }

  /** Primeiro dia com marcação, tipo, obs ou ajuste ≠ 0. */
  function inicio(s) {
    let ini = null;
    for (const m of Object.keys(s.meses)) {
      const dias = (s.meses[m] && s.meses[m].dias) || {};
      for (const k of Object.keys(dias)) if (temRegistro(dias[k]) && (!ini || k < ini)) ini = k;
    }
    return ini;
  }

  /**
   * O dia entra no saldo? Entre o início e hoje, exceto dia com número ímpar de marcações:
   * o dia aberto agora (hoje ou virada da meia-noite) e o dia passado pendente, até ser corrigido.
   */
  function contaNoSaldo(s, k, agora, ini) {
    const hoje = keyDate(agora ? new Date(agora) : new Date());
    if (ini === undefined) ini = inicio(s);
    if (!ini || k < ini || k > hoje) return false;
    const x = getDia(s, k);
    return !(x && x.p && x.p.length % 2);
  }

  function resumo(s, m, agora) {
    agora = agora ? new Date(agora) : new Date();
    const ini = inicio(s), n = diasNoMes(m);
    let tr = 0, pv = 0, ex = 0, dv = 0;
    for (let d = 1; d <= n; d++) {
      const k = m + '-' + pad(d);
      if (!contaNoSaldo(s, k, agora, ini)) continue;
      const t = trabalhado(s, k, agora), v = previsto(s, k), sd = saldoDia(s, k, agora);
      tr += t; pv += v;
      if (sd > 0) ex += sd; else dv -= sd;
    }
    return { tr, pv, ex, dv, saldo: ex - dv };
  }

  /** Meses com dados no estado mais todos os meses de `inicio` até o mês de `agora`. */
  function mesesDoBanco(s, agora) {
    const set = new Set(Object.keys(s.meses)), ini = inicio(s);
    if (ini) {
      const fim = keyMes(agora ? new Date(agora) : new Date());
      let [y, m] = ini.slice(0, 7).split('-').map(Number);
      for (let k = ini.slice(0, 7); k <= fim; k = y + '-' + pad(m)) { set.add(k); if (++m > 12) { m = 1; y++; } }
    }
    return [...set].sort();
  }

  /** Banco total = saldo anterior + saldo de todos os meses (do início até hoje). */
  function total(s, agora) {
    return (Number(s.perfil.saldoIni) || 0) + mesesDoBanco(s, agora).reduce((a, m) => a + resumo(s, m, agora).saldo, 0);
  }

  // ---------- backup ----------
  const chaveDiaValida = k => {
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(k);
    if (!m) return false;
    const d = new Date(Date.UTC(+m[1], +m[2] - 1, +m[3]));
    return d.getUTCFullYear() === +m[1] && d.getUTCMonth() === +m[2] - 1 && d.getUTCDate() === +m[3];
  };
  const objeto = v => !!v && typeof v === 'object' && !Array.isArray(v);

  /**
   * Valida um backup. Aceita o envelope {app:'ponto-vic', versao:2, exportadoEm, dados:{meses,perfil}}
   * e o formato antigo {meses, perfil}. Devolve {ok, erro, dados, resumo}.
   */
  function validarBackup(v) {
    const falha = erro => ({ ok: false, erro, dados: null, resumo: null });
    if (!objeto(v)) return falha('o arquivo não é um backup do Ponto VIC.');
    let dados = v;
    if ('dados' in v || v.app === 'ponto-vic') {
      if (v.app !== 'ponto-vic') return falha('envelope de backup desconhecido.');
      if (v.versao !== 2) return falha('versão de backup não suportada (' + String(v.versao).slice(0, 10) + ').');
      if (!isoValido(v.exportadoEm)) return falha('data de exportação inválida.');
      if (!objeto(v.dados)) return falha('backup sem dados.');
      dados = v.dados;
    }
    if (!objeto(dados.meses)) return falha('falta a lista de meses (arquivo truncado?).');
    if (!objeto(dados.perfil)) return falha('falta o perfil (arquivo truncado?).');
    const P = dados.perfil;
    if (P.nome !== undefined && typeof P.nome !== 'string') return falha('nome inválido no perfil.');
    if (P.saldoIni !== undefined && !num(P.saldoIni)) return falha('saldo anterior inválido no perfil.');
    if (P.feriadosAuto !== undefined && typeof P.feriadosAuto !== 'boolean') return falha('opção de feriados inválida no perfil.');
    if (P.ultimoBackup != null && !isoValido(P.ultimoBackup)) return falha('data do último backup inválida no perfil.');
    let nDias = 0, de = null, ate = null;
    for (const m of Object.keys(dados.meses)) {
      if (!/^\d{4}-(0[1-9]|1[0-2])$/.test(m)) return falha('mês inválido: ' + m.slice(0, 20));
      const M = dados.meses[m];
      if (!objeto(M) || !objeto(M.dias)) return falha('mês ' + m + ' sem dias.');
      for (const k of Object.keys(M.dias)) {
        const br = k.slice(8, 10) + '/' + k.slice(5, 7) + '/' + k.slice(0, 4);
        if (!chaveDiaValida(k)) return falha('data impossível: ' + k.slice(0, 20));
        if (k.slice(0, 7) !== m) return falha('o dia ' + br + ' está no mês ' + m + '.');
        const x = M.dias[k];
        if (!objeto(x)) return falha('dia ' + br + ' inválido.');
        if (x.p !== undefined) {
          if (!Array.isArray(x.p) || x.p.some(t => !isoValido(t))) return falha('marcações inválidas em ' + br + '.');
          for (let i = 1; i < x.p.length; i++) if (new Date(x.p[i]) < new Date(x.p[i - 1])) return falha('marcações fora de ordem em ' + br + '.');
        }
        if (x.g !== undefined && (!Array.isArray(x.g) || x.g.some(g => g !== null && !gpsValido(g)))) return falha('localização (GPS) inválida em ' + br + '.');
        if (x.tipo !== undefined && x.tipo !== '' && !TIPOS.includes(x.tipo)) return falha('tipo do dia inválido em ' + br + '.');
        if (x.banco !== undefined && typeof x.banco !== 'boolean') return falha('"pagando banco" inválido em ' + br + '.');
        if (x.ajuste !== undefined && !num(x.ajuste)) return falha('ajuste inválido em ' + br + '.');
        if (x.obs !== undefined && typeof x.obs !== 'string') return falha('observação inválida em ' + br + '.');
        if (temDados(normalizarDia(x))) { nDias++; if (!de || k < de) de = k; if (!ate || k > ate) ate = k; }
      }
    }
    const limpo = normalizar({ meses: dados.meses, perfil: dados.perfil });
    return { ok: true, erro: '', dados: limpo, resumo: { meses: Object.keys(limpo.meses).length, dias: nDias, de, ate } };
  }

  /** Comparação estável (ordem das chaves não importa). */
  function estavel(v) {
    if (Array.isArray(v)) return '[' + v.map(estavel).join(',') + ']';
    if (v && typeof v === 'object') return '{' + Object.keys(v).sort().map(k => JSON.stringify(k) + ':' + estavel(v[k])).join(',') + '}';
    return JSON.stringify(v);
  }

  /**
   * Conflito no mesmo dia (a = local, b = remoto), campo a campo:
   * p: mais marcações vence, empate fica o local; g: por marcação (GPS válido do lado com o mesmo ISO, local primeiro); obs e tipo: o não vazio, se ambos, o local;
   * banco: OU; ajuste: o local se ≠ 0, senão o remoto. Outros campos: local por cima do remoto.
   */
  function mesclarDia(a, b) {
    const pg = b.p.length > a.p.length ? b : a;
    // GPS por marcação: para cada ISO do p vencedor, o GPS válido de quem tem esse ISO (local primeiro)
    const gpsDe = (lado, iso) => { const i = lado.p.indexOf(iso); return i >= 0 && gpsValido(lado.g[i]) ? limparGPS(lado.g[i]) : null; };
    return Object.assign({}, b, a, {
      p: pg.p.slice(), g: pg.p.map(iso => gpsDe(a, iso) || gpsDe(b, iso) || null),
      obs: a.obs || b.obs || '',
      tipo: a.tipo || b.tipo || '',
      banco: a.banco === true || b.banco === true,
      ajuste: a.ajuste ? a.ajuste : (b.ajuste || 0)
    });
  }

  /**
   * Merge por dia entre os meses locais e o snapshot remoto (nuvem).
   * Dia só de um lado: entra. Dia dos dois lados e diferente: mesclarDia (campo a campo).
   * @returns {{meses:Object, reenviar:string[]}} reenviar = meses cujo resultado difere do remoto.
   */
  function mesclarMeses(locais, remotos) {
    const L = normalizar({ meses: locais || {} }).meses, R = normalizar({ meses: remotos || {} }).meses, out = {}, reenviar = [];
    for (const m of new Set([...Object.keys(L), ...Object.keys(R)])) {
      const dl = (L[m] && L[m].dias) || {}, dr = (R[m] && R[m].dias) || {}, dias = {};
      for (const k of new Set([...Object.keys(dl), ...Object.keys(dr)])) {
        const a = dl[k], b = dr[k];
        if (!a) dias[k] = b;
        else if (!b) dias[k] = a;
        else if (estavel(a) === estavel(b)) dias[k] = b;
        else dias[k] = mesclarDia(a, b);
      }
      out[m] = Object.assign({}, R[m] || {}, L[m] || {}, { dias });
      if (!R[m] || estavel(dias) !== estavel(dr)) reenviar.push(m);
    }
    return { meses: out, reenviar: reenviar.sort() };
  }

  // ---------- horas ----------
  /** '12:30', '-03:15', '+0:30', '45' -> minutos; inválido -> null. */
  function parseHM(v) {
    const m = String(v == null ? '' : v).trim().match(/^([+-])?\s*(\d+)(?::(\d{1,2}))?$/);
    if (!m || (m[3] !== undefined && Number(m[3]) > 59)) return null;
    return (m[1] === '-' ? -1 : 1) * (Number(m[2]) * 60 + Number(m[3] || 0));
  }
  function formatHM(v) {
    v = Math.round(Number(v) || 0);
    const a = Math.abs(v);
    return (v < 0 ? '-' : '') + pad(Math.floor(a / 60)) + ':' + pad(a % 60);
  }
  const formatHMsinal = v => (v > 0 ? '+' : '') + formatHM(v);

  function diasEntre(a, b) {
    const out = [];
    for (let k = a; k <= b; k = somaDias(k, 1)) out.push(k);
    return out;
  }

  /** Nome da marcação i num dia com n marcações (igual ao app original). */
  function nomeMarcacao(i, n) {
    const nomes = n <= 2 ? ['Entrada', 'Saída'] : ['Entrada', 'Saída almoço', 'Volta almoço', 'Saída'];
    return nomes[i] || (i % 2 ? 'Saída' : 'Entrada');
  }

  // ---------- editor de dia ----------
  /** Marcações -> [{hm:'HH:MM', dias:deslocamento em dias a partir de k}] (hora local). */
  function horariosDoDia(k, p) {
    const base = dateKey(k);
    return (p || []).map(iso => {
      const d = new Date(iso), meio = new Date(d.getFullYear(), d.getMonth(), d.getDate(), 12);
      return { hm: horaLocal(iso), dias: Math.round((meio - base) / 864e5) };
    });
  }

  /**
   * Monta as marcações do dia k a partir dos horários do editor ('HH:MM', na ordem digitada; vazios são ignorados).
   * Virada da meia-noite: cada horário menor que o anterior vai para o dia seguinte (acumulando). Sem sort.
   * Marcação antiga com mesmo HH:MM no mesmo dia é reaproveitada (ISO original com segundos + GPS).
   * @returns {{p:string[], g:Array, dias:number[]}} dias[i] = deslocamento (0 = mesmo dia, 1 = +1 dia...)
   */
  function montarMarcacoes(k, horarios, pAnt, gAnt) {
    const antigos = horariosDoDia(k, pAnt || []), usados = new Set();
    const p = [], g = [], dias = [];
    let off = 0, ant = null;
    for (const raw of horarios || []) {
      const m = String(raw == null ? '' : raw).trim().match(/^(\d{1,2}):(\d{2})$/);
      if (!m || Number(m[1]) > 23 || Number(m[2]) > 59) continue;
      const min = Number(m[1]) * 60 + Number(m[2]);
      if (ant !== null && min < ant) off++;
      ant = min;
      const hh = pad(m[1]) + ':' + m[2];
      const j = antigos.findIndex((a, i) => !usados.has(i) && a.hm === hh && a.dias === off);
      if (j >= 0) {
        usados.add(j);
        p.push(pAnt[j]);
        g.push((gAnt && gAnt[j]) || null);
      } else {
        p.push(new Date(somaDias(k, off) + 'T' + hh + ':00').toISOString()); // hora LOCAL (sem Z)
        g.push(null);
      }
      dias.push(off);
    }
    return { p, g, dias };
  }

  return {
    MESES, TIPOS, pad, keyDate, keyMes, dateKey, diaSemana, somaDias, diasNoMes, clone, util, horaLocal,
    diaVazio, normalizarDia, normalizar, getDia, dia, comDia, temDados, temRegistro, gpsValido, limparGPS,
    ultimoDiaComMarcacao, ultimoDiaImpar, diaAberto, pendentes, destinoBatida, mesclarDia, decidirBatida, removerUltima, LIMITE_ABERTO,
    validarBackup, mesclarMeses, estavel,
    pascoa, nomeFeriado, feriado, feriadoAuto,
    previsto, almocoAuto, trabalhado, saldoDia, inicio, contaNoSaldo, resumo, mesesDoBanco, total,
    parseHM, formatHM, formatHMsinal, diasEntre, nomeMarcacao, horariosDoDia, montarMarcacoes
  };
});
