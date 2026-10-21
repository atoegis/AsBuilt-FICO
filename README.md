# Painel As Built — FICO

Painel de controle da documentação As Built (Ferrovia de Integração Centro-Oeste), com cinco abas:
Visão geral, Contratadas, Curva S e aderência, Documentos obrigatórios e Mapa de avanço
(mapa em MapLibre com fundos OpenStreetMap, Humanitário, OpenTopoMap e satélite).

## Arquivos
- `index.html` — o painel.
- `dash-shim.js` — carrega os dados da pasta `dados/` e controla os filtros.
- `fonts.css` — fontes DM Sans e Bebas Neue embutidas.
- `config.js` — configuração da leitura direta do SharePoint.
- `sharepoint.js` — login Microsoft e leitura da planilha pelo Microsoft Graph.
- `vendor/msal-browser.min.js` — biblioteca oficial de login da Microsoft (MSAL 3.27).
- `PEDIDO_TI.md` — o que pedir à TI.
- `dados/` — bases em CSV:
  - `asbuilt_documentos.csv` e `asbuilt_obrigatorios.csv`: planilha NOVA BD ASBUILT (abas BD ASBUILT e Documentos Obrigatorios),
    com as colunas "Segmentos (mapa)" e "OAE (mapa)" calculadas a partir da Tabela de segmentos rev9;
  - `tracado_segmentos.csv`, `tracado_oaes.csv`, `tracado_municipios.csv`: Tabela de segmentos rev9;
  - `tracado_eixo.csv`: eixo georreferenciado km 0 a 292 (mesma geometria do Mapa Superestrutura).

## Publicar (GitHub + Firebase Hosting)
1. Crie um repositório público.
2. Envie todo o conteúdo desta pasta para a raiz do repositório, mantendo a subpasta `dados/`.
3. Publique a raiz do repositório no Firebase Hosting (public: `.`), como os demais sites.

O painel precisa ser aberto pelo endereço publicado. Abrindo o index.html direto do computador,
o navegador bloqueia a leitura dos CSV.

## Dados ao vivo do SharePoint
Com `clientId` e `tenantId` preenchidos em `config.js`, o painel mostra o botão **Entrar com Microsoft**.
Depois do login, ele lê as tabelas `PBI` (aba BD ASBUILT) e `PBI_9` (aba Documentos Obrigatorios)
da planilha NOVA BD ASBUILT no SharePoint, a cada abertura ou no botão **Atualizar agora**.
Veja em `PEDIDO_TI.md` o que pedir à TI para gerar esses dois valores.

Sem login (ou sem `config.js` preenchido), o painel usa a cópia salva em `dados/`.

## Regras que o painel aplica na planilha
- Usa as tabelas `PBI` e `PBI_9`. Não renomeie as tabelas nem os cabeçalhos das colunas.
- Na BD ASBUILT, considera apenas "Necessário Emissão de As Built?" = SIM ou AVALIAR.
- Localiza cada documento no traçado pela coluna Segmento ou, se vazia, pelo Ativo (lista de OAEs da Tabela de segmentos rev9).
  Nomes novos de OAE ou de segmento precisam ser incluídos em `sharepoint.js` (listas ATIVO e SEGNAME).
