# Painel As Built — FICO

Painel de controle da documentação As Built (Ferrovia de Integração Centro-Oeste), com cinco abas:
Visão geral, Contratadas, Documentos obrigatórios e Mapa de avanço
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
  - `asbuilt_documentos.csv` e `asbuilt_obrigatorios.csv`: planilha NOVA BD ASBUILT-PAINEL (abas BD ASBUILT e Documentos Obrigatorios),
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
da planilha NOVA BD ASBUILT-PAINEL no SharePoint, a cada abertura ou no botão **Atualizar agora**.
Veja em `PEDIDO_TI.md` o que pedir à TI para gerar esses dois valores.

Sem login (ou sem `config.js` preenchido), o painel usa a cópia salva em `dados/`.

## Regras que o painel aplica na planilha
- Usa as tabelas `PBI` e `PBI_9`. Não renomeie as tabelas nem os cabeçalhos das colunas.
- Na BD ASBUILT, "Necessário Emissão de As Built?" = OBRIGATÓRIO conta como SIM. A coluna "Obrigatório" (SIM/NÃO) dos dados guarda essa marcação e alimenta o filtro "Obrigatório".
- Nos Documentos Obrigatórios, linhas com "Necessário Emissão de As Built?" = AVALIAR ficam fora do painel.
- PACOTE 1A e 1B são tratados como PACOTE 1 (também no diagrama linear).
- Lote definido pelo Pacote: LOTE 01 = PCT 1, 2 e 3; LOTE 02 = PCT 4, 5, 6, 7 e 8; LOTE 03 = PCT 9, 10 e 11.
  Sem pacote, o painel mantém o lote da planilha se for LOTE 1, 2, 3 (gravado como LOTE 01, 02, 03) ou GERAL; outros valores (ex.: LOTE 4, LOTE 5) ficam sem lote.
- Nos Documentos Obrigatórios, linhas com "Necessário Emissão de As Built?" = AVALIAR ficam fora do painel.
- Na BD ASBUILT, considera apenas "Necessário Emissão de As Built?" = SIM ou AVALIAR.
- Localiza cada documento no traçado pela coluna Segmento ou, se vazia, pelo Ativo (lista de OAEs da Tabela de segmentos rev9).
  Nomes novos de OAE ou de segmento precisam ser incluídos em `sharepoint.js` (listas ATIVO e SEGNAME).
