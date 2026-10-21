/* Substitui a API "dash" do painel do Claude para rodar como site estático (GitHub Pages).
   Lê os arquivos da pasta dados/ e guarda os filtros na própria página. */
window.dash = (function () {
  const SOURCES = {
    documentos: 'dados/asbuilt_documentos.csv',
    obrigatorios: 'dados/asbuilt_obrigatorios.csv',
    tracado_segmentos: 'dados/tracado_segmentos.csv',
    tracado_oaes: 'dados/tracado_oaes.csv',
    tracado_municipios: 'dados/tracado_municipios.csv',
    tracado_eixo: 'dados/tracado_eixo.csv'
  };
  const DEFAULTS = { emp: null, lote: null, pac: null, seg: null, esp: null, sdoc: null, sdet: null, pend: null,
    ob_emp: null, ob_lote: null, ob_pac: null, ob_ativo: null, ref: null, mbase: 'bd', mmet: 'pf', mseg: null, mfundo: 'street' };
  let params = Object.assign({}, DEFAULTS);
  const state = {}, listeners = [];
  let raf = 0;
  const notify = () => { if (raf) return; raf = requestAnimationFrame(() => { raf = 0; listeners.forEach((fn) => { try { fn(); } catch (e) { console.error(e); } }); }); };
  function load(id) {
    const url = SOURCES[id];
    if (!url) { state[id] = { status: 'missing', data: [], columns: [], message: 'Fonte desconhecida' }; return; }
    state[id] = { status: 'loading', data: [], columns: [] };
    fetch(url, { cache: 'no-cache' }).then((r) => { if (!r.ok) throw new Error(r.status + ' ' + r.statusText); return r.text(); })
      .then((txt) => {
        let data, columns;
        if (url.endsWith('.json')) { data = JSON.parse(txt); columns = data.length ? Object.keys(data[0]) : []; }
        else { data = d3.csvParse(txt); columns = data.columns; }
        state[id] = { status: 'ok', data, columns, rows: data, meta: [] };
        notify();
      })
      .catch((e) => { state[id] = { status: 'error', data: [], columns: [], message: String(e) }; notify(); });
  }
  window.addEventListener('resize', notify);
  return {
    data(id) { if (!state[id]) load(id); return state[id]; },
    onData(fn) { listeners.push(fn); fn(); },
    params() { return params; },
    setParams(o) { params = Object.assign({}, params, o); notify(); },
    resetParams() { params = Object.assign({}, DEFAULTS); notify(); },
    setLink(name) { try { history.replaceState(null, '', '#' + name); } catch (e) {} },
    refresh(id) { load(id); },
    _set(id, data) { state[id] = { status: 'ok', data, columns: data.columns || [], rows: data, meta: [] }; notify(); },
    calc() {}, loader() {},
    colors: ['#3578bc', '#6fa1d8', '#4fae7b', '#7c8aa3', '#e07d12', '#45566f', '#7c8aa3', '#6b3db5']
  };
})();
