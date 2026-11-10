/* Ponto VIC — UI, armazenamento e eventos. Depende de PontoCore, PontoImport e PontoCartao. */
(function () {
  'use strict';
  const C = window.PontoCore, I = window.PontoImport, K = window.PontoCartao;
  const $ = id => document.getElementById(id);
  const pad = C.pad, hm = C.formatHM, hora = C.horaLocal;
  const LK = 'pontoVIC2', LK_PRE = 'pontoVIC2_preimport', LK_PREREST = 'pontoVIC2_prerestore', LK_TEMA = 'pontoTema';
  const XLSX_MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  const esc = t => String(t == null ? '' : t).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const dataBR = k => k.slice(8, 10) + '/' + k.slice(5, 7) + '/' + k.slice(0, 4);
  const sinal = v => (v >= 0 ? 'pos' : 'neg');

  // ---------- armazenamento: local, com nuvem opcional (window.claude, só no claude.ai) ----------
  const claude = window.claude || { use: async () => null };
  const usar = async nome => { try { return await claude.use(nome); } catch (e) { return null; } };
  let DB = null, UID = null, fila = Promise.resolve();
  let hidratando = false; // true enquanto a nuvem (window.claude) carrega: ponto bloqueado
  const TEMPO_NUVEM = 8000;

  function carregar() {
    let raw = null;
    try { raw = localStorage.getItem(LK); } catch (e) { /* sem storage */ }
    if (!raw) return C.normalizar({});
    try { return C.normalizar(JSON.parse(raw)); } catch (e) {
      try { localStorage.setItem(LK + '_corrompido', raw); } catch (e2) { /* ignora */ }
      return C.normalizar({});
    }
  }
  let state = carregar();

  function local() {
    try { localStorage.setItem(LK, JSON.stringify({ meses: state.meses, perfil: state.perfil })); return true; } catch (e) {
      aviso('Não foi possível salvar neste aparelho (armazenamento cheio ou bloqueado).');
      return false;
    }
  }
  function syncMes(m) {
    if (!DB) return;
    const dias = C.clone((state.meses[m] && state.meses[m].dias) || {});
    fila = fila.then(() => DB.collection('data/users/' + UID).doc(m).set({ dias })).catch(e => aviso('Falha ao sincronizar: ' + (e.message || e.code)));
  }
  function syncPerfil() {
    if (!DB) return;
    const p = C.clone(state.perfil);
    fila = fila.then(() => DB.collection('data/users/' + UID).doc('perfil').set(p)).catch(() => { });
  }
  /** Grava localmente e sincroniza os meses indicados (e o perfil, se pedido). */
  function salvar(meses, perfil) {
    const ok = local();
    (meses || []).forEach(syncMes);
    if (perfil) syncPerfil();
    return ok;
  }
  const todosMeses = (...estados) => [...new Set(estados.flatMap(s => Object.keys(s.meses)))];

  /** Conecta à nuvem (só no claude.ai). Bloqueia o ponto até o snapshot chegar ou 8 s passarem; depois faz merge por dia. */
  async function conectar() {
    const SO_LOCAL = 'Os registros ficam salvos neste celular. Exporte o backup (.json) e o cartão em Excel todo mês.';
    if (!window.claude) { $('sync').textContent = SO_LOCAL; return; }
    hidratando = true; render();
    const carregarNuvem = (async () => {
      const db = await usar('db'), user = await usar('user');
      if (!db || !user) return { tipo: 'local' };
      const uid = await user.id();
      const snap = await db.collection('data/users/' + uid).get();
      return { tipo: 'ok', db, uid, snap };
    })().catch(e => ({ tipo: 'erro', e }));
    const r = await Promise.race([carregarNuvem, new Promise(res => setTimeout(() => res({ tipo: 'tempo' }), TEMPO_NUVEM))]);
    hidratando = false;
    if (r.tipo !== 'ok') {
      $('sync').textContent = r.tipo === 'local' ? SO_LOCAL : 'Não foi possível conectar à nuvem; os registros ficam salvos neste aparelho.';
      if (r.tipo === 'tempo') aviso('A nuvem não respondeu em 8 s. Seguindo só neste aparelho.');
      render(); return;
    }
    const remotos = {}; let perfilRemoto = null;
    r.snap.docs.forEach(d => { const v = d.data(); if (d.id === 'perfil') perfilRemoto = v; else remotos[d.id] = v; });
    const mx = C.mesclarMeses(state.meses, remotos);
    state = C.normalizar({ meses: mx.meses, perfil: perfilRemoto || state.perfil });
    DB = r.db; UID = r.uid;
    local();
    mx.reenviar.forEach(syncMes);           // meses em que o local venceu ou que só existem aqui
    if (!perfilRemoto) syncPerfil();
    $('sync').textContent = 'Sincronizado na nuvem: os registros aparecem em qualquer aparelho onde você abrir o app.';
    preencherCfg(); render();
  }

  // ---------- avisos e diálogos ----------
  let avisoT = null;
  function aviso(t) { $('msg').textContent = t; clearTimeout(avisoT); avisoT = setTimeout(() => { $('msg').textContent = ''; }, 4500); }
  /** Confirmação dentro da UI (substitui confirm()). */
  function confirmar(texto, ok) {
    return new Promise(res => {
      const d = $('dlgConf');
      $('confTxt').textContent = texto;
      $('confOk').textContent = ok || 'Confirmar';
      d.returnValue = '';
      d.onclose = () => { d.onclose = null; res(d.returnValue === 'ok'); };
      d.showModal();
    });
  }

  async function baixar(nome, blob) {
    const dl = await usar('downloads');
    if (dl) {
      try { await dl.save({ filename: nome, data: blob }); } catch (err) { if (err && err.code !== 'declined') aviso('Não foi possível baixar: ' + (err.message || err.code)); return false; }
      return true;
    }
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = nome;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 60000);
    return true;
  }

  let excelP = null;
  function carregarExcel() {
    if (window.ExcelJS) return Promise.resolve();
    if (!excelP) excelP = new Promise((res, rej) => {
      const s = document.createElement('script');
      s.src = 'vendor/exceljs.min.js';
      s.onload = () => res();
      s.onerror = () => { excelP = null; s.remove(); rej(new Error('não foi possível carregar o leitor de Excel')); };
      document.head.appendChild(s);
    });
    return excelP;
  }

  // ---------- tela Hoje ----------
  const CIRC = 326.73; // 2π·52 (anel do SVG)
  const durHM = m => { m = Math.max(0, Math.round(m)); return Math.floor(m / 60) + 'h' + pad(m % 60); };
  const compacto = v => (v > 0 ? '+' : v < 0 ? '-' : '') + Math.floor(Math.abs(v) / 60) + ':' + pad(Math.abs(v) % 60);
  function linkGPS(g) {
    if (!C.gpsValido(g)) return '';
    const url = 'https://maps.google.com/?q=' + encodeURIComponent(g.lat.toFixed(6) + ',' + g.lon.toFixed(6));
    return '<a class="gps" href="' + esc(url) + '" target="_blank" rel="noopener" aria-label="' + esc('Ver no mapa (precisão ~' + Math.round(g.acc) + ' m)') + '"><svg aria-hidden="true"><use href="#i-pin"/></svg></a>';
  }
  function setHTML(el, h) { if (el._h !== h) { el.innerHTML = h; el._h = h; } } // evita recriar a lista a cada segundo

  function renderHoje() {
    const now = new Date(), hoje = C.keyDate(now), dest = C.destinoBatida(state, now);
    const k = dest.anterior ? dest.k : hoje, x = C.getDia(state, k) || C.diaVazio(), open = dest.acao === 'saida';
    $('clock').textContent = pad(now.getHours()) + ':' + pad(now.getMinutes()) + ':' + pad(now.getSeconds());
    const dtx = now.toLocaleDateString('pt-BR', { weekday: 'long', day: '2-digit', month: 'long' });
    $('date').textContent = dtx.charAt(0).toUpperCase() + dtx.slice(1);
    // botões do almoço: SAÍDA ALMOÇO só no 1º período do dia; VOLTA ALMOÇO enquanto está no almoço
    const podeAlmoco = open && !dest.anterior && x.p.length === 1, noAlmoco = !open && !dest.anterior && C.voltaDoAlmoco(state, k) !== null;
    $('bIn').hidden = open || noAlmoco; $('bOut').hidden = !open; $('bAlm').hidden = !podeAlmoco; $('bVolta').hidden = !noAlmoco;
    $('bIn').disabled = open || noAlmoco || hidratando; $('bOut').disabled = !open || hidratando;
    $('bAlm').disabled = !podeAlmoco || hidratando; $('bVolta').disabled = !noAlmoco || hidratando;
    // o interruptor vale para o dia que a tela mostra (o dia anterior aberto, quando a SAÍDA vai fechá-lo)
    $('banco').checked = (C.getDia(state, k) || {}).banco === true;
    $('bancoRot').textContent = k === hoje ? 'Pagando banco hoje' : 'Pagando banco em ' + dataBR(k).slice(0, 5);
    const ad = $('avisoDia');
    if (dest.anterior) {
      const quando = (dest.k === C.somaDias(hoje, -1) ? 'ontem ' : dataBR(dest.k) + ' ') + hora(dest.desde);
      $('avisoDiaTxt').textContent = 'Ponto aberto desde ' + quando + ' · SAÍDA fecha o dia ' + dataBR(dest.k).slice(0, 5);
      ad.hidden = false;
    } else ad.hidden = true;
    renderPendentes(now);
    $('todayTit').textContent = dest.anterior ? 'Marcações de ' + dataBR(k).slice(0, 5) : 'Marcações de hoje';
    const w = C.trabalhado(state, k, now), v = C.previsto(state, k), sd = C.saldoDia(state, k, now), fer = C.feriadoAuto(state, k);
    const va = !open && !dest.anterior && !hidratando ? C.voltaDoAlmoco(state, k) : null;
    let st;
    if (hidratando) st = 'Sincronizando…';
    else if (va) st = 'No almoço desde ' + hora(va.saiu);
    else if (open) st = 'Trabalhando há ' + durHM((now - new Date(x.p[x.p.length - 1])) / 6e4);
    else if (x.p.length) st = 'Fora desde ' + hora(x.p[x.p.length - 1]);
    else if (x.tipo) st = x.tipo.charAt(0) + x.tipo.slice(1).toLowerCase();
    else if (fer) st = C.nomeFeriado(k);
    else if (!C.util(k)) st = 'Fim de semana';
    else st = 'Sem marcação hoje';
    $('status').textContent = st;
    const alm = open ? C.almocoAuto(state, k, now) : 0;
    const atraso = va ? (now - new Date(va.volta)) / 6e4 : 0;
    $('hProg').textContent = va
      ? (atraso < 1 ? 'Volte às ' + hora(va.volta) + ' (almoço de ' + durHM(va.almoco) + ')' : 'Volta prevista às ' + hora(va.volta) + ' · passou ' + durHM(atraso))
      : (v ? hm(w) + ' de ' + hm(v) : hm(w) + ' trabalhadas') + (alm ? ' · almoço de ' + durHM(alm) + ' descontado' : '');
    const frac = v ? Math.min(1, w / v) : (w > 0 ? 1 : 0), anel = $('anel');
    $('anelProg').style.strokeDashoffset = (CIRC * (1 - frac)).toFixed(2);
    anel.classList.toggle('cheio', v > 0 && w >= v);
    anel.classList.toggle('aberto', open);
    anel.setAttribute('role', 'img');
    anel.setAttribute('aria-label', (dest.anterior ? 'Trabalhado no dia ' + dataBR(k) + ' ' : 'Trabalhado hoje ') + hm(w) + (v ? ' de ' + hm(v) : ''));
    $('hSaldo').textContent = 'Saldo do dia ' + C.formatHMsinal(sd);
    $('hSaldo').className = 'chip ' + sinal(sd);
    $('hPrev').textContent = v ? 'Jornada ' + hm(v) + (x.banco ? ' · banco' : '') : 'Sem jornada prevista';
    const offs = C.horariosDoDia(k, x.p);
    let h = x.p.length ? x.p.map((t, i) => '<div class="tl"><span class="dot' + (i % 2 ? ' out' : '') + '"></span><span class="n">' + C.nomeMarcacao(i, x.p.length, x.almoco) +
      '</span><span class="t">' + linkGPS(x.g && x.g[i]) + hora(t) + (offs[i].dias > 0 ? ' <span class="tag mais1">+' + offs[i].dias + ' dia</span>' : '') + '</span></div>').join('')
      : '<div class="vazio">Nenhuma marcação hoje. Toque em ENTRADA para começar.</div>';
    if (va) h += '<div class="tl total prevista"><span class="dot vazio-dot"></span><span class="n">Volta almoço (prevista)</span><span class="t">' + hora(va.volta) + '</span></div>';
    if (x.ajuste) h += '<div class="tl total"><span class="n">Ajuste</span><span class="t">' + C.formatHMsinal(x.ajuste) + '</span></div>';
    if (x.obs) h += '<div class="tl total"><span class="n">' + esc(x.obs) + '</span></div>';
    setHTML($('today'), h);
    const r = C.resumo(state, C.keyMes(now), now), tot = C.total(state, now);
    $('kMes').textContent = C.formatHMsinal(r.saldo); $('kMes').className = sinal(r.saldo);
    $('kExt').textContent = hm(r.ex);
    $('kTot').textContent = C.formatHMsinal(tot); $('kTot').className = 'kbig ' + sinal(tot);
    renderLembrete(now);
  }

  // dias com ponto incompleto (marcação ímpar no passado): ficam fora do saldo até serem corrigidos
  const MAX_PEND = 5;
  let verTodosPend = false;
  function renderPendentes(now) {
    const lista = C.pendentes(state, now), box = $('pendBox');
    box.hidden = !lista.length;
    if (!lista.length) { verTodosPend = false; setHTML($('pendLista'), ''); return; }
    const n = lista.length;
    $('pendTit').textContent = n + (n === 1 ? ' dia com ponto incompleto, fora do saldo' : ' dias com ponto incompleto, fora do saldo');
    const mostra = verTodosPend ? lista : lista.slice(0, MAX_PEND);
    setHTML($('pendLista'), mostra.map(p => {
      const dt = C.dateKey(p.k).toLocaleDateString('pt-BR', { weekday: 'short' }).replace('.', '');
      return '<div class="pend-item"><span><b>' + esc(dataBR(p.k).slice(0, 5)) + '</b> ' + esc(dt) + '<small>última marcação ' + esc(hora(p.desde)) + ', sem saída</small></span>' +
        '<button class="btn" type="button" data-k="' + esc(p.k) + '" aria-label="' + esc('Corrigir dia ' + dataBR(p.k).slice(0, 5)) + '">Corrigir</button></div>';
    }).join(''));
    $('pendTodos').hidden = n <= MAX_PEND;
    $('pendTodos').textContent = verTodosPend ? 'Mostrar menos' : 'Ver todos (' + n + ')';
  }

  function diasDesdeBackup(now) {
    const u = state.perfil.ultimoBackup;
    return u ? Math.floor((now - new Date(u)) / 864e5) : null;
  }
  function renderLembrete(now) {
    const d = diasDesdeBackup(now), precisa = !!C.inicio(state) && (d === null || d > 30);
    $('lembrete').hidden = !precisa;
    if (precisa) $('lembreteTxt').textContent = d === null ? 'Você ainda não fez backup dos seus registros.' : 'Último backup há ' + d + ' dias.';
  }

  // ---------- marcar ponto ----------
  let desfazerT = null;
  const vibrar = ms => { try { if (navigator.vibrate) navigator.vibrate(ms); } catch (e) { /* ignora */ } };
  function esconderDesfazer() { clearTimeout(desfazerT); document.body.classList.remove('com-snack'); $('desfazerBox').hidden = true; $('desfazerBox').classList.remove('corre'); $('desfazer').onclick = null; }
  function mostrarDesfazer(k, iso, texto) {
    esconderDesfazer();
    $('desfTxt').textContent = texto;
    const box = $('desfazerBox');
    box.hidden = false; document.body.classList.add('com-snack');
    void box.offsetWidth; box.classList.add('corre'); // reinicia a barra de 10 s
    $('desfazer').onclick = () => {
      if (C.removerUltima(state, k, iso)) { salvar([k.slice(0, 7)]); render(); vibrar(15); aviso('Marcação desfeita.'); }
      esconderDesfazer();
    };
    desfazerT = setTimeout(esconderDesfazer, 10000);
  }

  let ultimoToque = null; // ms da última batida aceita (trava de toque duplo)
  /** @param almoco true nos botões SAÍDA ALMOÇO e VOLTA ALMOÇO (a regra de entrada/saída é a mesma). */
  function bater(acao, almoco) {
    if (hidratando) { aviso('Aguarde: sincronizando com a nuvem…'); return; }
    const n = new Date(), dec = C.decidirBatida(state, n, acao, ultimoToque);
    if (!dec.ok) { render(); return; } // toque repetido (< 1,5 s) ou botão que não vale para o estado atual
    ultimoToque = n.getTime();
    const k = dec.destino.k, m = k.slice(0, 7), x = C.dia(state, k), iso = n.toISOString(), rotulo = (acao === 'entrada' ? (almoco ? 'Volta almoço' : 'Entrada') : (almoco ? 'Saída almoço' : 'Saída'));
    x.p.push(iso); x.g.push(null);
    if (almoco && acao === 'saida') x.almoco = true;
    salvar([m]); render();
    mostrarDesfazer(k, iso, rotulo + ' às ' + hora(iso) + (dec.destino.anterior ? ' · fechou o dia ' + dataBR(k).slice(0, 5) : ''));
    vibrar(35);
    const novo = ['bIn', 'bAlm', 'bOut', 'bVolta'].map($).find(b => !b.hidden) || $('bIn');
    novo.classList.remove('bateu'); void novo.offsetWidth; novo.classList.add('bateu');
    if (!navigator.geolocation) { aviso('GPS indisponível neste aparelho.'); return; }
    aviso('Obtendo localização…');
    const atual = () => { const y = C.getDia(state, k); const j = y ? y.p.indexOf(iso) : -1; return j >= 0 ? { y, j } : null; };
    navigator.geolocation.getCurrentPosition(pos => {
      const a = atual(); if (!a) return;
      a.y.g[a.j] = C.limparGPS({ lat: +pos.coords.latitude.toFixed(6), lon: +pos.coords.longitude.toFixed(6), acc: Math.round(pos.coords.accuracy) });
      salvar([m]); render(); aviso('Localização registrada.');
    }, err => {
      if (atual()) aviso('Marcação salva sem localização (' + (err.code === 1 ? 'permissão negada' : 'GPS não respondeu') + ').');
    }, { enableHighAccuracy: true, timeout: 15000, maximumAge: 0 });
  }
  $('bIn').onclick = () => bater('entrada');
  $('bOut').onclick = () => bater('saida');
  $('bAlm').onclick = () => bater('saida', true);
  $('bVolta').onclick = () => bater('entrada', true);
  $('pendLista').onclick = e => { const b = e.target.closest('button[data-k]'); if (b) abrirEditor(b.dataset.k); };
  $('pendTodos').onclick = () => { verTodosPend = !verTodosPend; renderPendentes(new Date()); };
  /** Chave do dia que a tela Hoje mostra: o dia aberto anterior quando a próxima SAÍDA vai fechá-lo; senão hoje. */
  function diaDaTela(now) { const d = C.destinoBatida(state, now); return d.anterior ? d.k : C.keyDate(now); }
  $('banco').onchange = e => {
    const k = diaDaTela(new Date()), x = C.dia(state, k);
    x.banco = e.target.checked;
    if (!C.temDados(x)) delete state.meses[k.slice(0, 7)].dias[k];
    salvar([k.slice(0, 7)]); render();
  };

  // ---------- tela Mês ----------
  const ver = new Date(); ver.setDate(1);
  const LK_VISAO = 'pontoVisaoMes';
  let visao = 'lista';
  try { if (localStorage.getItem(LK_VISAO) === 'cal') visao = 'cal'; } catch (e) { /* ignora */ }
  function setVisao(v) { visao = v; try { localStorage.setItem(LK_VISAO, v); } catch (e) { /* ignora */ } renderMes(); }
  $('vLista').onclick = () => setVisao('lista');
  $('vCal').onclick = () => setVisao('cal');

  function infoDia(k, now, ini, hoje) {
    const x = C.getDia(state, k) || C.diaVazio(), fer = C.feriadoAuto(state, k), w = C.diaSemana(k);
    return { k, x, fer, w, futuro: k > hoje, conta: C.contaNoSaldo(state, k, now, ini), sd: C.saldoDia(state, k, now) };
  }
  function renderMes() {
    const now = new Date(), m = C.keyMes(ver), r = C.resumo(state, m, now), hoje = C.keyDate(now), ini = C.inicio(state);
    const pend = new Set(C.pendentes(state, now).map(p => p.k));
    const nomeMes = C.MESES[ver.getMonth()];
    $('mTit').textContent = nomeMes.charAt(0) + nomeMes.slice(1).toLowerCase() + ' ' + ver.getFullYear();
    $('mTrab').textContent = hm(r.tr); $('mPrev').textContent = hm(r.pv); $('mExt').textContent = hm(r.ex); $('mDev').textContent = hm(r.dv);
    $('mSaldo').textContent = C.formatHMsinal(r.saldo); $('mSaldo').className = 'kbig ' + sinal(r.saldo);
    $('lData').max = hoje;
    if (!$('lData').value) $('lData').value = hoje;
    const cal = visao === 'cal';
    $('vLista').classList.toggle('on', !cal); $('vLista').setAttribute('aria-selected', !cal);
    $('vCal').classList.toggle('on', cal); $('vCal').setAttribute('aria-selected', cal);
    $('mDias').hidden = cal; $('mCal').hidden = !cal;
    const n = C.diasNoMes(m);
    if (!cal) {
      let h = '';
      for (let d = 1; d <= n; d++) {
        const D = infoDia(m + '-' + pad(d), now, ini, hoje), k = D.k, x = D.x, fer = D.fer, w = D.w;
        if (D.futuro && !C.temDados(x)) continue;
        if (!x.p.length && !C.util(k) && !C.temDados(x)) continue;
        const tags = [];
        if (pend.has(k)) tags.push('<span class="tag incompleto">INCOMPLETO</span>');
        if (x.tipo) tags.push('<span class="tag tipo">' + esc(x.tipo) + '</span>');
        else if (fer) tags.push('<span class="tag fer" title="' + esc(C.nomeFeriado(k)) + '">FERIADO</span>');
        else if (w === 6 || w === 0) tags.push('<span class="tag">' + (w === 6 ? 'SÁBADO' : 'DOMINGO') + '</span>');
        if (x.banco) tags.push('<span class="tag">BANCO</span>');
        if (x.ajuste) tags.push('<span class="tag aj' + (x.ajuste < 0 ? ' neg' : '') + '">AJUSTE ' + C.formatHMsinal(x.ajuste) + '</span>');
        const offs = C.horariosDoDia(k, x.p), ult = x.p.length - 1;
        const horas = x.p.length ? hora(x.p[0]) + ' – ' + (x.p.length > 1 ? hora(x.p[ult]) + (offs[ult].dias > 0 ? ' (+' + offs[ult].dias + 'd)' : '') : (pend.has(k) ? 'sem saída' : 'em aberto')) +
          (x.p.length === 4 ? ' · almoço ' + hora(x.p[1]) + '–' + hora(x.p[2]) : x.p.length > 2 ? ' · ' + x.p.length + ' marcações' : '') : (fer && !x.tipo ? C.nomeFeriado(k) : 'Sem marcação');
        h += '<div class="day' + (k === hoje ? ' hoje' : '') + '" data-k="' + k + '" role="button" tabindex="0" aria-label="Editar ' + dataBR(k) + '">' +
          '<div class="d-num"><b>' + pad(d) + '</b><span>' + esc(C.dateKey(k).toLocaleDateString('pt-BR', { weekday: 'short' }).replace('.', '')) + '</span></div>' +
          '<div class="d-mid"><div class="d-h">' + esc(horas) + '</div>' + (tags.length ? '<div class="d-tags">' + tags.join('') + '</div>' : '') +
          (x.obs ? '<div class="d-obs">' + esc(x.obs) + '</div>' : '') + '</div>' +
          (D.conta ? '<span class="pill ' + sinal(D.sd) + '">' + C.formatHMsinal(D.sd) + '</span>' : '<span class="pill mut">—</span>') + '</div>';
      }
      $('mDias').innerHTML = h || '<div class="vazio">Nenhum registro neste mês.</div>';
    } else {
      let c = '<div class="cal-g">' + ['D', 'S', 'T', 'Q', 'Q', 'S', 'S'].map(l => '<div class="cal-h" aria-hidden="true">' + l + '</div>').join('');
      for (let i = 0; i < C.diaSemana(m + '-01'); i++) c += '<div></div>';
      for (let d = 1; d <= n; d++) {
        const D = infoDia(m + '-' + pad(d), now, ini, hoje), k = D.k, x = D.x, fer = D.fer;
        const cls = ['cal-d'];
        if (!C.util(k)) cls.push('fds');
        if (k === hoje) cls.push('hoje');
        if (x.tipo || fer) cls.push('marca');
        let small = '', estilo = '', rot = dataBR(k);
        if (pend.has(k)) {
          cls.push('incompleto');
          small = 'INCOMP';
          rot += ', INCOMPLETO: ponto sem saída, fora do saldo';
        } else if (D.conta && (x.p.length || C.util(k) || C.temDados(x))) {
          cls.push(D.sd > 0 ? 'pos' : D.sd < 0 ? 'neg' : 'zero');
          small = compacto(D.sd);
          estilo = ' style="--a:' + (0.10 + Math.min(Math.abs(D.sd), 300) / 300 * 0.16).toFixed(3) + '"';
          rot += ', saldo ' + C.formatHMsinal(D.sd);
        } else if (x.tipo || fer) small = (x.tipo || 'FER').slice(0, 3);
        if (x.tipo) rot += ', ' + x.tipo; else if (fer) rot += ', feriado';
        const bloq = D.futuro && !C.temDados(x);
        c += '<button type="button" class="' + cls.join(' ') + '" data-k="' + k + '"' + estilo + (bloq ? ' disabled' : '') + ' aria-label="' + esc(rot) + '"><b>' + d + '</b><small>' + esc(small) + '</small></button>';
      }
      c += '</div><div class="cal-leg"><span><i style="background:var(--pos-bg);box-shadow:inset 0 0 0 1px var(--pos)"></i>extra</span><span><i style="background:var(--neg-bg);box-shadow:inset 0 0 0 1px var(--neg)"></i>devida</span><span><i style="background:var(--acc);border-radius:50%"></i>feriado/folga</span><span><i style="box-shadow:inset 0 0 0 2px var(--neg)"></i>INCOMP = incompleto, fora do saldo</span></div>';
      $('mCal').innerHTML = c;
    }
  }
  $('prev').onclick = () => { ver.setMonth(ver.getMonth() - 1); renderMes(); };
  $('next').onclick = () => { ver.setMonth(ver.getMonth() + 1); renderMes(); };
  $('mDias').onclick = e => { const r = e.target.closest('.day'); if (r) abrirEditor(r.dataset.k); };
  $('mDias').onkeydown = e => { if (e.key === 'Enter' || e.key === ' ') { const r = e.target.closest('.day'); if (r) { e.preventDefault(); abrirEditor(r.dataset.k); } } };
  $('mCal').onclick = e => { const b = e.target.closest('.cal-d'); if (b && !b.disabled) abrirEditor(b.dataset.k); };
  $('lancar').onclick = () => {
    const v = $('lData').value, hoje = C.keyDate(new Date());
    if (!/^\d{4}-\d{2}-\d{2}$/.test(v)) { aviso('Escolha uma data.'); return; }
    if (v > hoje) { aviso('Só dá para lançar até hoje.'); return; }
    abrirEditor(v);
  };

  // ---------- editor de dia ----------
  let ed = null; // { k, pAnt, gAnt }
  function horariosEditor() { return [...$('dMarc').querySelectorAll('input[type=time]')].map(i => i.value); }
  function linhasEditor(valores) {
    const n = valores.length;
    $('dMarc').innerHTML = valores.map((v, i) => {
      const nome = C.nomeMarcacao(i, n, (C.getDia(state, ed.k) || {}).almoco);
      return '<div class="mrow" data-i="' + i + '"><label for="dm' + i + '">' + nome + '</label>' +
        '<div class="cmp"><input type="time" id="dm' + i + '" value="' + esc(v) + '"><span class="mais1 tag" hidden></span><span class="mgps"></span></div>' +
        '<button type="button" class="btn ico rm" data-i="' + i + '" aria-label="Remover ' + nome + '"><svg aria-hidden="true"><use href="#i-x"/></svg></button></div>';
    }).join('') || '<div class="vazio">Sem marcações. Use "Adicionar marcação".</div>';
    previaEditor();
  }
  /** Dia que resultará do editor; null se o ajuste for inválido. */
  function diaDoEditor() {
    const ant = C.getDia(state, ed.k) || C.diaVazio();
    const mont = C.montarMarcacoes(ed.k, horariosEditor(), ed.pAnt, ed.gAnt);
    const aTxt = $('dAjuste').value.trim(), aj = aTxt ? C.parseHM(aTxt) : 0;
    const d = Object.assign({}, ant, { p: mont.p, g: mont.g, almoco: ant.almoco === true && mont.p.length >= 2, tipo: $('dTipo').value, banco: $('dBanco').checked, obs: $('dObs').value.trim(), ajuste: aj === null ? 0 : aj });
    return { d, mont, ajusteInvalido: aj === null };
  }
  function previaEditor() {
    const { d, mont, ajusteInvalido } = diaDoEditor();
    // "+1 dia" e GPS ao lado de cada horário preenchido
    let j = 0;
    $('dMarc').querySelectorAll('.mrow').forEach(row => {
      const inp = row.querySelector('input'), b = row.querySelector('.mais1'), gps = row.querySelector('.mgps');
      if (/^\d{1,2}:\d{2}$/.test(inp.value)) {
        const off = mont.dias[j];
        b.hidden = !(off > 0); b.textContent = '+' + off + ' dia';
        gps.innerHTML = linkGPS(mont.g[j]);
        j++;
      } else { b.hidden = true; gps.innerHTML = ''; }
    });
    const now = new Date(), s2 = C.comDia(state, ed.k, d);
    const t = C.trabalhado(s2, ed.k, now), v = C.previsto(s2, ed.k), sd = C.saldoDia(s2, ed.k, now);
    const det = ['Trabalhado ' + hm(t) + ' · Previsto ' + hm(v)];
    if (d.ajuste) det.push('Ajuste ' + C.formatHMsinal(d.ajuste));
    if (d.p.length % 2) det.push('Última marcação em aberto');
    if (C.feriadoAuto(s2, ed.k)) det.push('Feriado: ' + esc(C.nomeFeriado(ed.k)));
    $('dPrevia').innerHTML = '<div class="det">' + det.join('<br>') + '</div><div class="sd"><small>Saldo do dia</small><b class="' + sinal(sd) + '">' + C.formatHMsinal(sd) + '</b></div>';
    $('dErro').textContent = ajusteInvalido ? 'Ajuste inválido. Use ±HH:MM, ex.: 00:30 ou -01:15.' : '';
  }
  function abrirEditor(k) {
    const x = C.getDia(state, k) || C.diaVazio();
    ed = { k, pAnt: x.p.slice(), gAnt: x.g.slice() };
    const tit = C.dateKey(k).toLocaleDateString('pt-BR', { weekday: 'long', day: '2-digit', month: '2-digit', year: 'numeric' });
    $('dTit').textContent = tit.charAt(0).toUpperCase() + tit.slice(1);
    $('dTipo').value = x.tipo || '';
    $('dBanco').checked = !!x.banco;
    $('dObs').value = x.obs || '';
    $('dAjuste').value = x.ajuste ? C.formatHMsinal(x.ajuste) : '';
    const hs = C.horariosDoDia(k, x.p).map(h => h.hm);
    linhasEditor(hs.length ? hs : ['', '']);
    $('dlg').showModal();
  }
  $('dMarc').addEventListener('input', previaEditor);
  $('dMarc').addEventListener('click', e => {
    const b = e.target.closest('.rm'); if (!b) return;
    const vals = horariosEditor(); vals.splice(Number(b.dataset.i), 1); linhasEditor(vals);
  });
  $('dAdd').onclick = () => {
    const vals = horariosEditor(); vals.push(''); linhasEditor(vals);
    const ins = $('dMarc').querySelectorAll('input'); if (ins.length) ins[ins.length - 1].focus();
  };
  ['dTipo', 'dBanco', 'dObs', 'dAjuste'].forEach(id => $(id).addEventListener('input', previaEditor));
  $('dTipo').addEventListener('change', previaEditor);
  $('dBanco').addEventListener('change', previaEditor);
  function gravarDia(k, d) {
    const m = k.slice(0, 7);
    if (!C.temDados(d)) { if (state.meses[m]) delete state.meses[m].dias[k]; }
    else { state.meses[m] = state.meses[m] || { dias: {} }; state.meses[m].dias[k] = d; }
    salvar([m]); render();
  }
  $('dSalvar').onclick = () => {
    const { d, ajusteInvalido } = diaDoEditor();
    if (ajusteInvalido) { $('dErro').textContent = 'Ajuste inválido. Use ±HH:MM, ex.: 00:30 ou -01:15.'; $('dAjuste').focus(); return; }
    gravarDia(ed.k, d);
    $('dlg').close('ok');
    aviso('Dia ' + dataBR(ed.k) + ' salvo.');
  };
  $('dLimpar').onclick = async () => {
    if (!(await confirmar('Limpar o dia ' + dataBR(ed.k) + '?\nApaga marcações, GPS, tipo, observação e ajuste.', 'Limpar dia'))) return;
    gravarDia(ed.k, C.diaVazio());
    $('dlg').close('x');
    aviso('Dia ' + dataBR(ed.k) + ' limpo.');
  };

  // ---------- cartão de ponto (Excel) ----------
  $('xlsx').onclick = async () => {
    const b = $('xlsx'); b.disabled = true;
    try {
      await carregarExcel();
      const m = C.keyMes(ver), wb = K.montar(window.ExcelJS, state, m, new Date());
      const buf = await wb.xlsx.writeBuffer();
      await baixar(K.nomeArquivo(m), new Blob([buf], { type: XLSX_MIME }));
    } catch (e) { aviso('Falha ao gerar o Excel: ' + (e.message || e)); } finally { b.disabled = false; }
  };

  // ---------- backup ----------
  async function exportarBackup() {
    const agora = new Date();
    state.perfil.ultimoBackup = agora.toISOString();
    const env = { app: 'ponto-vic', versao: 2, exportadoEm: agora.toISOString(), dados: { meses: state.meses, perfil: state.perfil } };
    const ok = await baixar('ponto-vic-backup-' + C.keyDate(agora) + '.json', new Blob([JSON.stringify(env, null, 1)], { type: 'application/json' }));
    if (ok) { salvar([], true); render(); preencherCfg(); aviso('Backup exportado.'); }
  }
  $('exportar').onclick = exportarBackup;
  $('lembreteBtn').onclick = exportarBackup;

  function descrever(r, s) {
    const per = r.de ? ', de ' + dataBR(r.de) + ' a ' + dataBR(r.ate) : '';
    return r.meses + ' mês(es), ' + r.dias + ' dia(s) com registro' + per + '; banco total ' + C.formatHMsinal(C.total(s, new Date()));
  }
  function resumoEstado(s) {
    const v = C.validarBackup({ meses: s.meses, perfil: s.perfil });
    return v.ok ? v.resumo : { meses: Object.keys(s.meses).length, dias: 0, de: null, ate: null };
  }
  /** Troca o estado inteiro de forma transacional: só troca se a gravação local der certo. */
  function substituirEstado(novo) {
    try { localStorage.setItem(LK, JSON.stringify({ meses: novo.meses, perfil: novo.perfil })); } catch (e) {
      aviso('Não foi possível gravar neste aparelho; nada foi alterado.'); return false;
    }
    const antes = state;
    state = novo;
    todosMeses(antes, novo).forEach(syncMes); syncPerfil();
    preencherCfg(); render();
    return true;
  }
  /** Lê uma cópia de segurança interna (pré-importação / pré-restauração). */
  function lerCopia(chave) {
    let b = null;
    try { b = JSON.parse(localStorage.getItem(chave)); } catch (e) { b = null; }
    const v = C.validarBackup(b);
    return v.ok ? { dados: v.dados, em: b._em || b._preimportEm || null } : null;
  }
  function guardarCopia(chave) {
    try { localStorage.setItem(chave, JSON.stringify({ meses: state.meses, perfil: state.perfil, _em: new Date().toISOString() })); return true; } catch (e) { return false; }
  }
  $('restaurar').onclick = () => $('arquivoBackup').click();
  $('arquivoBackup').onchange = async e => {
    const f = e.target.files[0]; e.target.value = '';
    if (!f) return;
    let v;
    try { v = JSON.parse(await f.text()); } catch (err) { aviso('Backup inválido: não é um JSON completo (arquivo truncado?).'); return; }
    const val = C.validarBackup(v);
    if (!val.ok) { aviso('Backup inválido: ' + val.erro); return; }
    const novo = val.dados;
    const ok = await confirmar('Restaurar o backup "' + f.name + '"?\n\nBackup: ' + descrever(val.resumo, novo) +
      '.\nAtual: ' + descrever(resumoEstado(state), state) + '.\n\nOs dados atuais serão substituídos (dá para desfazer em Ajustes).', 'Restaurar');
    if (!ok) return;
    if (!guardarCopia(LK_PREREST)) { aviso('Não foi possível guardar a cópia de segurança; restauração cancelada.'); return; }
    if (!substituirEstado(novo)) { try { localStorage.removeItem(LK_PREREST); } catch (err) { /* ignora */ } return; }
    $('restResultado').textContent = 'Backup restaurado.';
    preencherCfg();
    aviso('Backup restaurado.');
  };
  $('desfazerRestore').onclick = async () => {
    const c = lerCopia(LK_PREREST);
    if (!c) { aviso('Não há restauração para desfazer.'); preencherCfg(); return; }
    const quando = c.em ? new Date(c.em).toLocaleString('pt-BR') : 'antes da restauração';
    if (!(await confirmar('Desfazer a restauração e voltar aos dados de ' + quando + '?\nO que foi registrado depois disso também será desfeito.', 'Desfazer restauração'))) return;
    if (!substituirEstado(c.dados)) return;
    try { localStorage.removeItem(LK_PREREST); } catch (e) { /* ignora */ }
    $('restResultado').textContent = 'Restauração desfeita.';
    preencherCfg();
    aviso('Restauração desfeita.');
  };

  // ---------- importação da planilha antiga ----------
  let imp = null; // { linhas, avisos, nome, plano }
  $('importar').onclick = () => $('arquivoXlsx').click();
  $('arquivoXlsx').onchange = async e => {
    const f = e.target.files[0]; e.target.value = '';
    if (!f) return;
    try {
      aviso('Lendo a planilha…');
      await carregarExcel();
      const wb = new window.ExcelJS.Workbook();
      await wb.xlsx.load(await f.arrayBuffer());
      const leitura = I.lerPlanilha(wb);
      if (!leitura.linhas.length) { aviso('Nenhuma data encontrada na planilha.'); return; }
      imp = { linhas: leitura.linhas, avisos: leitura.avisos, nome: f.name, plano: null };
      $('impArq').textContent = f.name + ' · ' + leitura.linhas.length + ' linhas com data';
      $('oAjustar').checked = true;
      $('oSobrescrever').checked = false;
      $('oZerarBox').hidden = state.perfil.saldoIni === 0;
      $('oZerar').checked = state.perfil.saldoIni !== 0;
      $('oZerarVal').textContent = hm(state.perfil.saldoIni);
      replanejar();
      $('msg').textContent = '';
      $('dlgImp').showModal();
    } catch (err) { aviso('Falha ao ler a planilha: ' + (err.message || err)); }
  };
  function replanejar() {
    imp.plano = I.planejar(state, imp.linhas, { ajustar: $('oAjustar').checked, sobrescrever: $('oSobrescrever').checked, zerarSaldoIni: $('oZerar').checked && state.perfil.saldoIni !== 0 });
    const P = imp.plano, td = v => '<td class="' + (v < 0 ? 'neg' : '') + '">' + hm(v) + '</td>';
    const qtd = T => '<small>' + T.dias + ' dia' + (T.dias === 1 ? '' : 's') + (T.ignorados ? ' · ' + T.ignorados + ' ign.' : '') + '</small>';
    let h = '<thead><tr><th>Mês</th><th>Planilha</th><th>App sem ajuste</th><th>Final</th></tr></thead><tbody>';
    for (const m of Object.keys(P.meses).sort()) {
      const T = P.meses[m];
      h += '<tr><td><b>' + C.MESES[Number(m.slice(5)) - 1].slice(0, 3) + '/' + m.slice(2, 4) + '</b>' + qtd(T) + '</td>' + td(T.planilha) + td(T.app) + td(T.final) + '</tr>';
    }
    const T = P.total;
    h += '</tbody><tfoot><tr><th>Total' + qtd(T) + '</th>' + td(T.planilha) + td(T.app) + td(T.final) + '</tr></tfoot>';
    $('impTab').innerHTML = h;
    $('impBanco').innerHTML = '<div class="rotulo">Banco de horas total</div><b>' + C.formatHMsinal(P.bancoAntes) + '</b> → <b class="' + sinal(P.bancoDepois) + '">' + C.formatHMsinal(P.bancoDepois) + '</b>' +
      (P.de ? '<div class="mut peq">Período: ' + dataBR(P.de) + ' a ' + dataBR(P.ate) + '</div>' : '');
    $('impAvisos').innerHTML = imp.avisos.concat(P.avisos.filter(a => !(P.avisoSaldoIni && a.startsWith('Você tem um saldo anterior')))).map(a => '<li>' + esc(a) + '</li>').join('');
    $('impOk').disabled = !P.itens.some(i => i.acao !== 'ignorar') && !P.opcoes.zerarSaldoIni;
  }
  ['oAjustar', 'oSobrescrever', 'oZerar'].forEach(id => { $(id).onchange = replanejar; });
  $('impOk').onclick = () => {
    if (!imp || !imp.plano) return;
    replanejar(); // garante plano com o estado e o horário atuais
    if (!guardarCopia(LK_PRE)) { aviso('Não foi possível guardar a cópia de segurança neste aparelho; importação cancelada.'); return; }
    const P = imp.plano, n = P.itens.filter(i => i.acao !== 'ignorar').length, ign = P.total.ignorados;
    if (!substituirEstado(I.aplicar(state, P))) { try { localStorage.removeItem(LK_PRE); } catch (e) { /* ignora */ } return; }
    $('dlgImp').close('ok');
    const txt = 'Importação concluída: ' + n + ' dia(s) importado(s)' + (ign ? ', ' + ign + ' ignorado(s)' : '') + '. Banco de horas total: ' + hm(C.total(state, new Date())) + '.';
    $('impResultado').textContent = txt;
    aviso(txt);
    imp = null;
  };
  $('desfazerImport').onclick = async () => {
    const c = lerCopia(LK_PRE);
    if (!c) { aviso('Não há importação para desfazer.'); preencherCfg(); return; }
    const quando = c.em ? new Date(c.em).toLocaleString('pt-BR') : 'antes da importação';
    if (!(await confirmar('Desfazer a importação e voltar aos dados de ' + quando + '?\nO que foi registrado depois disso também será desfeito.', 'Desfazer importação'))) return;
    if (!substituirEstado(c.dados)) return;
    try { localStorage.removeItem(LK_PRE); } catch (e) { /* ignora */ }
    $('impResultado').textContent = 'Importação desfeita.';
    preencherCfg();
    aviso('Importação desfeita.');
  };

  // ---------- ajustes ----------
  function preencherCfg() {
    $('cNome').value = state.perfil.nome || '';
    $('cSaldo').value = state.perfil.saldoIni ? hm(state.perfil.saldoIni) : '';
    $('cFeriados').checked = state.perfil.feriadosAuto !== false;
    const d = diasDesdeBackup(new Date());
    $('cBackupInfo').textContent = d === null ? 'Nenhum backup exportado ainda.' : 'Último backup: ' + new Date(state.perfil.ultimoBackup).toLocaleDateString('pt-BR') + (d > 0 ? ' (há ' + d + ' dias)' : ' (hoje)') + '.';
    let tem = false, temR = false;
    try { tem = !!localStorage.getItem(LK_PRE); temR = !!localStorage.getItem(LK_PREREST); } catch (e) { /* ignora */ }
    $('desfazerImport').hidden = !tem;
    $('desfazerRestore').hidden = !temR;
  }
  $('cSalvar').onclick = () => {
    const v = $('cSaldo').value.trim();
    let saldo = 0;
    if (v) { saldo = C.parseHM(v); if (saldo === null) { aviso('Saldo inválido. Use o formato 12:30 ou -03:15.'); return; } }
    state.perfil.saldoIni = saldo;
    state.perfil.nome = $('cNome').value.trim();
    salvar([], true); preencherCfg(); render(); aviso('Salvo.');
  };
  $('cFeriados').onchange = e => { state.perfil.feriadosAuto = e.target.checked; salvar([], true); render(); aviso(e.target.checked ? 'Feriados automáticos ligados.' : 'Feriados automáticos desligados.'); };

  // ---------- navegação e tema ----------
  document.querySelectorAll('.tabs button').forEach(b => {
    b.onclick = () => {
      document.querySelectorAll('.tabs button').forEach(x => { x.classList.toggle('on', x === b); if (x === b) x.setAttribute('aria-current', 'page'); else x.removeAttribute('aria-current'); });
      ['vHoje', 'vMes', 'vCfg'].forEach(v => {
        const el = $(v), mostra = v === b.dataset.v;
        if (mostra && el.hidden) { el.classList.remove('entra'); void el.offsetWidth; el.classList.add('entra'); }
        el.hidden = !mostra;
      });
      window.scrollTo(0, 0);
      if (b.dataset.v === 'vCfg') preencherCfg();
      render();
    };
  });
  function render() { renderHoje(); if (!$('vMes').hidden) renderMes(); }

  function temaSalvo() { try { return localStorage.getItem(LK_TEMA) || ''; } catch (e) { return ''; } }
  function escuro() { const t = document.documentElement.dataset.theme; return t ? t === 'dark' : matchMedia('(prefers-color-scheme: dark)').matches; }
  function aplicaTema(t) {
    if (t) document.documentElement.dataset.theme = t; else delete document.documentElement.dataset.theme;
    const e = escuro();
    $('tema').innerHTML = '<svg aria-hidden="true"><use href="#' + (e ? 'i-sol' : 'i-lua') + '"/></svg><span class="sr">' + (e ? 'Claro' : 'Escuro') + '</span>';
    $('tema').setAttribute('aria-label', e ? 'Usar tema claro' : 'Usar tema escuro');
    document.querySelectorAll('#temaSel button').forEach(b => { const on = b.dataset.t === (t || ''); b.classList.toggle('on', on); b.setAttribute('aria-checked', on); });
  }
  function setTema(t) { try { if (t) localStorage.setItem(LK_TEMA, t); else localStorage.removeItem(LK_TEMA); } catch (e) { /* ignora */ } aplicaTema(t); }
  aplicaTema(temaSalvo());
  $('tema').onclick = () => setTema(escuro() ? 'light' : 'dark');
  document.querySelectorAll('#temaSel button').forEach(b => { b.onclick = () => setTema(b.dataset.t); });
  try { matchMedia('(prefers-color-scheme: dark)').addEventListener('change', () => { if (!temaSalvo()) aplicaTema(''); }); } catch (e) { /* navegador antigo */ }

  // Enter num campo não deve acionar o botão padrão (Cancelar) dos formulários dos diálogos
  ['dlg', 'dlgImp'].forEach(id => $(id).addEventListener('keydown', e => { if (e.key === 'Enter' && e.target.tagName === 'INPUT') e.preventDefault(); }));

  // ---------- início ----------
  try { if (navigator.storage && navigator.storage.persist) navigator.storage.persist().catch(() => { }); } catch (e) { /* ignora */ }
  preencherCfg();
  render();
  setInterval(renderHoje, 1000);
  conectar();
})();
