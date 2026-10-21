/* Leitura direta da planilha NOVA BD ASBUILT no SharePoint pelo Microsoft Graph.
   Converte as linhas das tabelas PBI e PBI_9 no mesmo formato dos CSV de dados/ e entrega ao painel. */
(function () {
  const CFG = window.ASBUILT_CONFIG || {};
  const configurado = !!(CFG.clientId && window.msal);
  const GRAPH = 'https://graph.microsoft.com/v1.0';
  let app, account, busy = false;

  // ---------- interface ----------
  const css = document.createElement('style');
  css.textContent = `.ab-sync{position:relative;display:flex;flex-direction:column;align-items:flex-end;gap:6px;margin-left:auto}
.ab-sync button{font:inherit;font-size:12.5px;font-weight:700;color:#fff;background:rgba(255,255,255,.1);border:1px solid rgba(255,255,255,.3);border-radius:20px;padding:6px 14px;cursor:pointer}
.ab-sync button:hover{background:rgba(255,255,255,.18)}.ab-sync button:disabled{opacity:.6;cursor:wait}
.ab-sync span{font-size:11.5px;color:#c7d6ea;text-align:right;max-width:22rem}.ab-sync span b{color:#f89b32}`;
  document.head.append(css);
  let box, btn, msg;
  function ui() {
    const host = document.querySelector('.ab-band-in'); if (!host || box) return;
    box = document.createElement('div'); box.className = 'ab-sync';
    btn = document.createElement('button'); btn.type = 'button';
    msg = document.createElement('span');
    box.append(btn, msg); host.append(box);
    btn.addEventListener('click', () => (!configurado ? avisoConfig(true) : account ? atualizar(true) : entrar()));
  }
  const status = (b, m) => { ui(); if (!box) return; btn.textContent = b; msg.innerHTML = m; };

  function avisoConfig(clicou) {
    status('Conectar ao SharePoint', clicou
      ? 'Falta configurar o acesso: preencha <b>clientId</b> e <b>tenantId</b> em config.js com os códigos do aplicativo registrado pela TI (veja PEDIDO_TI.md).'
      : 'Mostrando a <b>cópia salva</b>. A leitura direta do SharePoint ainda não foi configurada.');
  }
  // ---------- autenticação ----------
  async function iniciar() {
    ui();
    if (!configurado) { avisoConfig(false); return; }
    try {
    app = new msal.PublicClientApplication({
      auth: { clientId: CFG.clientId, authority: 'https://login.microsoftonline.com/' + (CFG.tenantId || 'organizations'), redirectUri: location.origin + location.pathname },
      cache: { cacheLocation: 'localStorage' }
    });
    await app.initialize();
    const r = await app.handleRedirectPromise().catch(() => null);
    account = (r && r.account) || app.getAllAccounts()[0] || null;
    if (account) atualizar(false);
    else status('Entrar com Microsoft', 'Mostrando a <b>última cópia salva</b>. Entre para ler a planilha do SharePoint.');
    } catch (e) { status('Entrar com Microsoft', 'Erro ao iniciar o login Microsoft: ' + (e.errorMessage || e.message || e)); }
  }
  async function entrar() {
    try { const r = await app.loginPopup({ scopes: CFG.escopos }); account = r.account; atualizar(true); }
    catch (e) { status('Entrar com Microsoft', 'Não foi possível entrar: ' + (e.errorMessage || e.message || e)); }
  }
  async function token() {
    try { return (await app.acquireTokenSilent({ scopes: CFG.escopos, account })).accessToken; }
    catch (e) { return (await app.acquireTokenPopup({ scopes: CFG.escopos, account })).accessToken; }
  }
  async function graph(path, tk) {
    const r = await fetch(path.startsWith('http') ? path : GRAPH + path, { headers: { Authorization: 'Bearer ' + tk } });
    if (!r.ok) { let t = ''; try { t = (await r.json()).error.message; } catch (_) {} throw new Error(r.status + ' ' + (t || r.statusText)); }
    return r.json();
  }
  const shareId = (url) => 'u!' + btoa(unescape(encodeURIComponent(url))).replace(/=+$/, '').replace(/\//g, '_').replace(/\+/g, '-');

  async function lerTabela(base, nome, tk) {
    const head = await graph(`${base}/workbook/tables('${encodeURIComponent(nome)}')/headerRowRange?$select=values`, tk);
    const cols = head.values[0].map((c) => String(c).trim());
    const out = []; const N = 2000;
    for (let skip = 0; ; skip += N) {
      const pg = await graph(`${base}/workbook/tables('${encodeURIComponent(nome)}')/rows?$top=${N}&$skip=${skip}&$select=values`, tk);
      for (const row of pg.value) { const v = row.values[0], o = {}; cols.forEach((c, i) => { o[c] = v[i]; }); out.push(o); }
      status('Atualizando…', `Lendo ${nome}: ${out.length.toLocaleString('pt-BR')} linhas`);
      if (pg.value.length < N) break;
    }
    return out;
  }

  // ---------- conversão (mesmas regras usadas para gerar os CSV) ----------
  const CANON = Object.fromEntries(['AS Built Finalizado', 'Aprovado Pela Vale', 'Aguardando análise Vale', 'Pendente análise CDOC', 'Reprovado Pela Vale', 'Reprovado pelo CDOC', 'Não Emitido', 'Cancelado', 'Cancelado pela Vale', 'Sem Numeração Correspondente', 'Não se Aplica', 'Pendência da Implantação', 'Finalizado'].map((s) => [s.toLowerCase(), s]));
  const ATIVO = { 'CORREGO BALDAIA': [9], 'CRIXÁS AÇU': [8], 'GO-154': [7], 'GO-347': [1], 'GO-556': [3], 'PF SOBRE O CORREGO PEITO DE MOÇA': [29], 'PI GO-156': [11], 'TRAVESSIA GO-156': [11], 'PI GO-336': [13], 'TRAVESSIA GO-336': [13], 'PI SOB A MT-326': [28], 'TRAVESSIA GO-164': [17], 'TRAVESSIA MT-100': [21], 'PONTE CORIXO DA SAUDADE': [22], 'PONTE CORIXO MOLHA MALA': [24], 'PONTE CÓRREGO BARREIRO': [15], 'PONTE SOBRE CORREGO BARREIRO': [15], "PONTE RIBEIRÃO D'ANTA": [12], "RIBEIRÃO D'ANTA": [12], 'RIBEIRAO DANTA': [12], 'RIBEIRÃO DANTA': [12], 'PONTE RIO ARAGUAIA': [20], 'PONTE SOBRE O RIO ARAGUAIA': [20], 'PONTE RIO CRISTALINO': [23], 'PONTE RIO CRISTALINO E PONTE CORIXO MOLHA MALA': [23, 24], 'PONTE RIO CRIXÁS-MIRIM': [16], 'PONTE RIO DAS MORTES': [26], 'PONTE RIO PEIXE III': [19], 'RIO DO PEIXE I': [6], 'PONTE RIO SANTA MARIA': [14], 'PONTE RIO VERMELHO': [10], 'RIO VERMELHO': [10], 'RIBEIRÃO DO MEIO': [5], 'RIO DOS BOIS': [2], 'RIO PRETO': [18] };
  const SEGNAME = { 'PONTE RIO ARAGUAIA': [20], 'PONTE SOBRE O RIO ARAGUAIA': [20], 'PONTE RIO SANTA MARIA': [14], 'PONTE SOBRE CORREGO BARREIRO': [15], 'PF RIO PRETO': [18], 'PF RIO DO PEIXE III': [19], 'PONTE SOBRE RIO CRIXAS MIRIM': [16], 'RIO VERMELHO': [10], "RIBEIRÃO D'ANTA": [12], 'PI GO-156': [11], 'PI GO-336': [13], 'PI SOB A RODOVIA MT-100': [21], 'PI GO-164': [17] };
  const txt = (v) => { if (v == null) return ''; const s = String(v).replace(/\s+/g, ' ').trim(); return s === '0' || s === '0.0' ? '' : s; };
  const iso = (d) => d.toISOString().slice(0, 10);
  function dataDe(v) {
    if (v == null || v === '') return '';
    if (typeof v === 'number') { if (v < 36526) return ''; return iso(new Date(Date.UTC(1899, 11, 30) + Math.round(v) * 864e5)); }
    const m = String(v).match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/); if (m && +m[3] >= 2000) return `${m[3]}-${m[2].padStart(2, '0')}-${m[1].padStart(2, '0')}`;
    const n = Number(v); return isFinite(n) ? dataDe(n) : '';
  }
  function mapear(segAtivo, segTxt, ordem, oaeSeg) {
    const at = segAtivo.toUpperCase().trim(), sg0 = segTxt.toUpperCase().trim();
    let ids = ATIVO[at] || [];
    if (at === 'GO-347' && segTxt.includes('1.4B')) ids = [4];
    let segs = [];
    if (sg0 && sg0 !== 'GERAL') {
      if (SEGNAME[sg0]) segs = SEGNAME[sg0].map((i) => oaeSeg[i]);
      else {
        let u = sg0.replace('3.A.1', '3A.1').replace('3.1A', '3.1').replace(/SEGMENTO|SEG\.|SEG/g, '').replace(/PACOTE \d+ -/g, '');
        let codes = (u.match(/ALÇA|\d[A]?\.\d[AB]?/g) || []).filter((c) => ordem.includes(c));
        if (codes.length === 2) { const rest = u.replace(codes[0], '').replace(codes[1], ''); if (/\bAO?\b/.test(rest)) { const a = ordem.indexOf(codes[0]), b = ordem.indexOf(codes[1]); codes = ordem.slice(Math.min(a, b), Math.max(a, b) + 1); } }
        segs = [...new Set(codes)];
      }
    }
    if (!segs.length && ids.length) segs = [...new Set(ids.map((i) => oaeSeg[i]))];
    if (!ids.length && SEGNAME[sg0]) ids = SEGNAME[sg0];
    return { segs: segs.filter(Boolean).join('|'), oaes: ids.join('|') };
  }
  function converter(rows, ehBD, ref) {
    const out = []; let id = 0;
    for (const r of rows) {
      const nec = txt(r['Necessário Emissão de As Built?']).toUpperCase();
      if (ehBD && nec !== 'SIM' && nec !== 'AVALIAR') continue;
      let sd = txt(r['Status do As Built Detalhado']); sd = CANON[sd.toLowerCase()] || sd;
      let pend = txt(r['Pendência na implantação?']).toUpperCase(); pend = pend === 'SIM' ? 'SIM' : pend.startsWith('N') ? 'NÃO' : '';
      let disc = txt(r['Disciplina Descrição']); if (disc && disc === disc.toUpperCase()) disc = disc.toLowerCase().replace(/(^|[\s/])\S/g, (m) => m.toUpperCase());
      const seg = txt(r['Segmento']), ativo = txt(r['Ativo']).toUpperCase();
      const mp = mapear(ativo, seg, ref.ordem, ref.oaeSeg);
      const o = { 'ID': ++id, 'Número Vale': txt(r['Número Vale']), 'Título': txt(r['Título']) };
      if (ehBD) o['Necessário Emissão'] = nec;
      Object.assign(o, { 'Disciplina': disc, 'Especialidade': txt(r['Especialidade de Engenharia']), 'Lote': txt(r['Lote']).toUpperCase(), 'Pacote': txt(r['Pacote']).toUpperCase(), 'Segmento': seg, 'Ativo': ativo,
        'Empresa Responsável': txt(r['Empresa Responsável Pelo As Built']), 'Empresa Colaboradora': txt(r['Empresa Colaboradora']), 'NAP Associada': txt(r['NAP Associada']),
        'Data Prevista': dataDe(r['Data Prevista']), 'Data Reprogramada': dataDe(r['Data Reprogramada']), 'Data Ação CDOC': dataDe(r['Data da Ação do CDOC']), 'Data Tendência': dataDe(r['Data Tendência']), 'Data Conclusão': dataDe(r['Data Conclusão']),
        'Nº Reprovações': String(parseInt(r['Número de Reprovações'], 10) || 0), 'Pendência Implantação': pend, 'Status Documento': txt(r['Status Documento']), 'Status Detalhado': sd,
        'Segmentos (mapa)': mp.segs, 'OAE (mapa)': mp.oaes });
      if (!o['Número Vale'] && !o['Título']) continue;
      out.push(o);
    }
    out.columns = out.length ? Object.keys(out[0]) : [];
    return out;
  }
  function referencia() {
    const s = dash.data('tracado_segmentos'), o = dash.data('tracado_oaes');
    if (s.status !== 'ok' || o.status !== 'ok') return null;
    return { ordem: s.data.map((r) => r['Segmento']), oaeSeg: Object.fromEntries(o.data.map((r) => [+r['ID OAE'], r['Segmento']])) };
  }

  async function atualizar(manual) {
    if (busy) return; busy = true; ui(); if (btn) btn.disabled = true;
    try {
      status('Atualizando…', 'Conectando ao SharePoint…');
      let ref = referencia(); for (let i = 0; !ref && i < 50; i++) { await new Promise((r) => setTimeout(r, 200)); ref = referencia(); }
      if (!ref) throw new Error('tabelas de segmentos não carregaram');
      const tk = await token();
      const item = await graph(`/shares/${shareId(CFG.arquivo)}/driveItem?$select=id,name,lastModifiedDateTime,lastModifiedBy,parentReference`, tk);
      const base = `/drives/${item.parentReference.driveId}/items/${item.id}`;
      const bd = converter(await lerTabela(base, CFG.tabelaDocumentos, tk), true, ref);
      const ob = converter(await lerTabela(base, CFG.tabelaObrigatorios, tk), false, ref);
      dash._set('documentos', bd); dash._set('obrigatorios', ob);
      const quando = new Date(item.lastModifiedDateTime).toLocaleString('pt-BR', { dateStyle: 'short', timeStyle: 'short' });
      const quem = item.lastModifiedBy && item.lastModifiedBy.user ? ' por ' + item.lastModifiedBy.user.displayName : '';
      status('Atualizar agora', `Dados <b>ao vivo</b> do SharePoint · planilha salva em ${quando}${quem}`);
    } catch (e) {
      status(account ? 'Tentar de novo' : 'Entrar com Microsoft', 'Não foi possível ler o SharePoint (' + (e.message || e) + '). Mostrando a última cópia salva.');
    } finally { busy = false; if (btn) btn.disabled = false; }
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', iniciar); else iniciar();
})();
