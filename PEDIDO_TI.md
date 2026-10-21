# Pedido à TI — leitura da planilha NOVA BD ASBUILT pelo Painel As Built

## Objetivo
O Painel As Built (site estático hospedado no Firebase) precisa ler, em tempo real, a planilha
`NOVA BD ASBUILT.xlsx` do SharePoint:

https://globalvale.sharepoint.com/sites/LACAL/Shared Documents/FICO/FICO - AS BUILT/NOVA BD ASBUILT.xlsx

A leitura é feita no navegador de quem abre o painel, com a própria conta Microsoft do usuário
(permissão delegada). Ninguém vê dados aos quais já não tenha acesso no SharePoint.
O painel só lê: não grava nada e não guarda nenhuma senha ou segredo.

## O que precisa ser criado no Microsoft Entra ID (Azure AD) do locatário da Vale
1. **Registro de aplicativo**
   - Nome sugerido: `Painel As Built FICO`
   - Tipos de conta: *Somente contas neste diretório organizacional* (locatário único)
2. **Autenticação → Adicionar plataforma → Aplicativo de página única (SPA)**
   - URI de redirecionamento: o endereço do painel no Firebase, exatamente como é aberto
     (ex.: `https://NOME-DO-PROJETO.web.app/` e, se usado, `https://NOME-DO-PROJETO.firebaseapp.com/`)
3. **Permissões de API → Microsoft Graph → Permissões delegadas**
   - `Files.Read.All` (ler arquivos que o usuário pode acessar)
   - `User.Read` (padrão, para o login)
   - Se o consentimento de usuário estiver bloqueado no locatário: **Conceder consentimento do administrador**.
4. Se os usuários forem convidados (contas de outras empresas, como EGIS, adicionadas como B2B),
   eles entram normalmente desde que sejam convidados no locatário da Vale e tenham acesso ao site LACAL.

## O que a TI devolve
- **ID do aplicativo (cliente)** — Application (client) ID
- **ID do diretório (locatário)** — Directory (tenant) ID

Esses dois valores vão no arquivo `config.js` do painel. Nenhum segredo de cliente é necessário.

## Chamadas feitas pelo painel (somente leitura)
- `GET /shares/{url-codificada}/driveItem` — localizar o arquivo
- `GET /drives/{drive}/items/{item}/workbook/tables('PBI')/headerRowRange`
- `GET /drives/{drive}/items/{item}/workbook/tables('PBI')/rows?$top=2000&$skip=…`
- idem para a tabela `PBI_9`
