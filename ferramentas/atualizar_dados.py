"""Atualiza dados/asbuilt_documentos.csv e dados/asbuilt_obrigatorios.csv a partir da
planilha NOVA BD ASBUILT-PAINEL.xlsx (tabelas PBI e PBI_9).

Mesmas regras de conversão do sharepoint.js, para que a cópia salva e a leitura ao vivo
gerem exatamente os mesmos dados.

Uso:  python ferramentas/atualizar_dados.py "NOVA BD ASBUILT-PAINEL.xlsx"
"""
import csv, datetime as dt, json, re, sys
from pathlib import Path
import openpyxl

RAIZ = Path(__file__).resolve().parent.parent
DADOS = RAIZ / 'dados'

CANON = {s.lower(): s for s in ['AS Built Finalizado', 'Aprovado Pela Vale', 'Aguardando análise Vale', 'Pendente análise CDOC',
         'Reprovado Pela Vale', 'Reprovado pelo CDOC', 'Não Emitido', 'Cancelado', 'Cancelado pela Vale',
         'Sem Numeração Correspondente', 'Não se Aplica', 'Pendência da Implantação', 'Finalizado']}
ATIVO = {  # nomes novos da planilha (nov/2026)
         'PONTE CORREGO BALDAIA': [9], 'PONTE CORREGO PEITO DE MOÇA': [29], 'PONTE CRIXÁS AÇU': [8], 'PONTE RIBEIRÃO DO MEIO': [5], 'PONTE RIO DO PEIXE I': [6], 'PONTE RIO DOS BOIS': [2], 'PONTE RIO PRETO': [18], 'PI GO-154': [7], 'PI GO-164': [17], 'PI GO-347': [1], 'PI GO-556': [3], 'PI MT-100': [21], 'PI MT-326': [28],
         # nomes anteriores, mantidos para versões antigas da planilha
         'CORREGO BALDAIA': [9], 'CRIXÁS AÇU': [8], 'GO-154': [7], 'GO-347': [1], 'GO-556': [3], 'PF SOBRE O CORREGO PEITO DE MOÇA': [29],
         'PI GO-156': [11], 'TRAVESSIA GO-156': [11], 'PI GO-336': [13], 'TRAVESSIA GO-336': [13], 'PI SOB A MT-326': [28], 'TRAVESSIA GO-164': [17],
         'TRAVESSIA MT-100': [21], 'PONTE CORIXO DA SAUDADE': [22], 'PONTE CORIXO MOLHA MALA': [24], 'PONTE CÓRREGO BARREIRO': [15],
         'PONTE SOBRE CORREGO BARREIRO': [15], "PONTE RIBEIRÃO D'ANTA": [12], "RIBEIRÃO D'ANTA": [12], 'RIBEIRAO DANTA': [12], 'RIBEIRÃO DANTA': [12],
         'PONTE RIO ARAGUAIA': [20], 'PONTE SOBRE O RIO ARAGUAIA': [20], 'PONTE RIO CRISTALINO': [23], 'PONTE RIO CRISTALINO E PONTE CORIXO MOLHA MALA': [23, 24],
         'PONTE RIO CRIXÁS-MIRIM': [16], 'PONTE RIO DAS MORTES': [26], 'PONTE RIO PEIXE III': [19], 'RIO DO PEIXE I': [6], 'PONTE RIO SANTA MARIA': [14],
         'PONTE RIO VERMELHO': [10], 'RIO VERMELHO': [10], 'RIBEIRÃO DO MEIO': [5], 'RIO DOS BOIS': [2], 'RIO PRETO': [18]}
SEGNAME = {'PONTE RIO ARAGUAIA': [20], 'PONTE SOBRE O RIO ARAGUAIA': [20], 'PONTE RIO SANTA MARIA': [14], 'PONTE SOBRE CORREGO BARREIRO': [15],
           'PF RIO PRETO': [18], 'PF RIO DO PEIXE III': [19], 'PONTE SOBRE RIO CRIXAS MIRIM': [16], 'RIO VERMELHO': [10], "RIBEIRÃO D'ANTA": [12],
           'PI GO-156': [11], 'PI GO-336': [13], 'PI SOB A RODOVIA MT-100': [21], 'PI GO-164': [17]}
EXCEL0 = dt.datetime(1899, 12, 30)
# Lote definido pelo pacote: LOTE 01 = PCT 1-3; LOTE 02 = PCT 4-8; LOTE 03 = PCT 9-11
LOTE_PCT = {**{n: 'LOTE 01' for n in (1, 2, 3)}, **{n: 'LOTE 02' for n in (4, 5, 6, 7, 8)}, **{n: 'LOTE 03' for n in (9, 10, 11)}}


def pacote_de(p):
    """PACOTE 1A e 1B são tratados como PACOTE 1."""
    return re.sub(r'^PACOTE 1[AB]$', 'PACOTE 1', p)


def lote_de(pacote, lote):
    """Lote pelo número do pacote; sem pacote, mantém o da planilha se for LOTE 1, 2, 3 (como LOTE 01, 02, 03) ou GERAL; senão fica vazio."""
    ls = {LOTE_PCT.get(int(n)) for n in re.findall(r'\d+', pacote)}
    if len(ls) == 1 and None not in ls: return ls.pop()
    lote = re.sub(r'^LOTE 0?([123])$', r'LOTE 0\1', lote)
    return lote if lote in ('LOTE 01', 'LOTE 02', 'LOTE 03', 'GERAL') else ''


def txt(v):
    if v is None: return ''
    if isinstance(v, float) and v.is_integer(): v = int(v)
    if isinstance(v, (dt.datetime, dt.date)): v = v.strftime('%d/%m/%Y')
    s = re.sub(r'\s+', ' ', str(v)).strip()
    return '' if s in ('0', '0.0') else s


def data_de(v):
    if v is None or v == '': return ''
    if isinstance(v, dt.datetime): return v.date().isoformat() if v.year >= 2000 else ''
    if isinstance(v, dt.date): return v.isoformat() if v.year >= 2000 else ''
    if isinstance(v, (int, float)):
        return '' if v < 36526 else (EXCEL0 + dt.timedelta(days=round(v))).date().isoformat()
    m = re.search(r'(\d{1,2})/(\d{1,2})/(\d{4})', str(v))
    if m and int(m[3]) >= 2000: return f'{m[3]}-{int(m[2]):02d}-{int(m[1]):02d}'
    if isinstance(v, (dt.time, dt.timedelta)): return ''
    try: return data_de(float(v))
    except (TypeError, ValueError): return ''


def titulo(s):
    return re.sub(r'(^|[\s/])(\S)', lambda m: m[1] + m[2].upper(), s.lower())


def mapear(ativo, seg_txt, ordem, oae_seg):
    at, sg0 = ativo.upper().strip(), seg_txt.upper().strip()
    ids = list(ATIVO.get(at, []))
    if at in ('GO-347', 'PI GO-347') and '1.4B' in seg_txt: ids = [4]
    segs = []
    if sg0 and sg0 != 'GERAL':
        if sg0 in SEGNAME: segs = [oae_seg.get(i) for i in SEGNAME[sg0]]
        else:
            u = sg0.replace('3.A.1', '3A.1').replace('3.1A', '3.1')
            u = re.sub(r'SEGMENTO|SEG\.|SEG', '', u)
            u = re.sub(r'PACOTE \d+ -', '', u)
            codes = [c for c in re.findall(r'ALÇA|\d[A]?\.\d[AB]?', u) if c in ordem]
            if len(codes) == 2:
                rest = u.replace(codes[0], '', 1).replace(codes[1], '', 1)
                if re.search(r'\bAO?\b', rest):
                    a, b = ordem.index(codes[0]), ordem.index(codes[1])
                    codes = ordem[min(a, b):max(a, b) + 1]
            segs = list(dict.fromkeys(codes))
    if not segs and ids: segs = list(dict.fromkeys(oae_seg.get(i) for i in ids))
    if not ids and sg0 in SEGNAME: ids = SEGNAME[sg0]
    return '|'.join(s for s in segs if s), '|'.join(str(i) for i in ids)


def converter(linhas, eh_bd, ordem, oae_seg):
    out = []
    for r in linhas:
        g = lambda k: r.get(k)
        nec = txt(g('Necessário Emissão de As Built?')).upper()
        if eh_bd and nec not in ('SIM', 'AVALIAR'): continue
        if not eh_bd and nec == 'AVALIAR': continue  # obrigatórios: AVALIAR fica fora do painel
        sd = txt(g('Status do As Built Detalhado')); sd = CANON.get(sd.lower(), sd)
        pend = txt(g('Pendência na implantação?')).upper()
        pend = 'SIM' if pend == 'SIM' else 'NÃO' if pend.startswith('N') else ''
        disc = txt(g('Disciplina Descrição'))
        if disc and disc == disc.upper(): disc = titulo(disc)
        seg, ativo = txt(g('Segmento')), txt(g('Ativo')).upper()
        segs, oaes = mapear(ativo, seg, ordem, oae_seg)
        try: repr_ = str(int(float(g('Número de Reprovações'))))
        except (TypeError, ValueError): repr_ = '0'
        o = {'ID': 0, 'Número Vale': txt(g('Número Vale')), 'Título': txt(g('Título'))}
        if eh_bd: o['Necessário Emissão'] = nec
        o.update({'Disciplina': disc, 'Especialidade': txt(g('Especialidade de Engenharia')), 'Lote': lote_de(pacote_de(txt(g('Pacote')).upper()), txt(g('Lote')).upper()),
                  'Pacote': pacote_de(txt(g('Pacote')).upper()), 'Segmento': seg, 'Ativo': ativo,
                  'Empresa Responsável': txt(g('Empresa Responsável Pelo As Built')), 'Empresa Colaboradora': txt(g('Empresa Colaboradora')),
                  'NAP Associada': txt(g('NAP Associada')), 'Data Prevista': data_de(g('Data Prevista')),
                  'Data Reprogramada': data_de(g('Data Reprogramada')), 'Data Ação CDOC': data_de(g('Data da Ação do CDOC')),
                  'Data Tendência': data_de(g('Data Tendência')), 'Data Conclusão': data_de(g('Data Conclusão')),
                  'Nº Reprovações': repr_, 'Pendência Implantação': pend, 'Status Documento': txt(g('Status Documento')),
                  'Status Detalhado': sd, 'Segmentos (mapa)': segs, 'OAE (mapa)': oaes})
        if not o['Número Vale'] and not o['Título']: continue
        out.append(o)
    for i, o in enumerate(out, 1): o['ID'] = i
    return out


def localizar_tabelas(caminho):
    """Lê do pacote .xlsx onde fica cada tabela (nome -> aba, intervalo), sem abrir a planilha inteira."""
    import posixpath, zipfile
    from html import unescape
    z = zipfile.ZipFile(caminho)
    rid_alvo = dict(re.findall(r'Id="([^"]+)"[^>]*Target="([^"]+)"', z.read('xl/_rels/workbook.xml.rels').decode()))
    rid_alvo.update({k: v for v, k in re.findall(r'Target="([^"]+)"[^>]*Id="([^"]+)"', z.read('xl/_rels/workbook.xml.rels').decode())})
    abas = {}
    for tag in re.findall(r'<sheet [^>]*/>', z.read('xl/workbook.xml').decode()):
        nome, rid = re.search(r'name="([^"]+)"', tag)[1], re.search(r'r:id="([^"]+)"', tag)[1]
        abas[posixpath.normpath(posixpath.join('xl', rid_alvo[rid].lstrip('/').removeprefix('xl/')))] = unescape(nome)
    tabelas = {}
    for n in z.namelist():
        m = re.match(r'xl/worksheets/_rels/(sheet\d+\.xml)\.rels$', n)
        if not m: continue
        for alvo in re.findall(r'Target="([^"]+)"', z.read(n).decode()):
            if 'tables/' not in alvo: continue
            tx = z.read(posixpath.normpath(posixpath.join('xl/worksheets', alvo))).decode('utf-8', 'ignore')
            cab = re.search(r'<table [^>]*>', tx)[0]
            nome, ref = re.search(r'\bname="([^"]+)"', cab)[1], re.search(r'\bref="([^"]+)"', cab)[1]
            tabelas[nome.lower()] = (abas.get('xl/worksheets/' + m[1]), ref)
    return tabelas


def ler_tabela(wb, tabelas, nome):
    from openpyxl.utils.cell import range_boundaries
    if nome.lower() not in tabelas: raise SystemExit(f'Tabela {nome} não encontrada na planilha.')
    aba, ref = tabelas[nome.lower()]
    c0, r0, c1, r1 = range_boundaries(ref)
    it = wb[aba].iter_rows(min_row=r0, max_row=r1, min_col=c0, max_col=c1, values_only=True)
    cols = [str(c).strip() if c is not None else '' for c in next(it)]
    return [dict(zip(cols, l)) for l in it], aba


def referencia():
    with open(DADOS / 'tracado_segmentos.csv', encoding='utf-8') as f: ordem = [r['Segmento'] for r in csv.DictReader(f)]
    with open(DADOS / 'tracado_oaes.csv', encoding='utf-8') as f: oae_seg = {int(r['ID OAE']): r['Segmento'] for r in csv.DictReader(f)}
    return ordem, oae_seg


def gravar(nome, linhas):
    with open(DADOS / nome, 'w', encoding='utf-8', newline='') as f:
        w = csv.DictWriter(f, fieldnames=list(linhas[0].keys()), lineterminator='\n')
        w.writeheader(); w.writerows(linhas)


def main(caminho):
    cfg = {'documentos': 'PBI', 'obrigatorios': 'PBI_9'}
    tabelas = localizar_tabelas(caminho)
    wb = openpyxl.load_workbook(caminho, data_only=True, read_only=True)
    ordem, oae_seg = referencia()
    resumo = {}
    for chave, arq, eh_bd in (('documentos', 'asbuilt_documentos.csv', True), ('obrigatorios', 'asbuilt_obrigatorios.csv', False)):
        linhas, aba = ler_tabela(wb, tabelas, cfg[chave])
        conv = converter(linhas, eh_bd, ordem, oae_seg)
        gravar(arq, conv)
        resumo[chave] = {'aba': aba, 'linhas': len(conv)}
    stamp = {'atualizado_em': dt.datetime.now().astimezone().isoformat(timespec='minutes'), 'arquivo': re.sub(r'^[0-9a-f]{8}-', '', Path(caminho).name).replace('_', ' '), **resumo}
    (DADOS / 'atualizacao.json').write_text(json.dumps(stamp, ensure_ascii=False, indent=1), encoding='utf-8')
    print(json.dumps(stamp, ensure_ascii=False))


if __name__ == '__main__':
    main(sys.argv[1] if len(sys.argv) > 1 else 'NOVA BD ASBUILT-PAINEL.xlsx')
