/* Configuração da leitura direta do SharePoint (Microsoft Graph).
   Enquanto clientId estiver vazio, o painel usa os CSV da pasta dados/.
   Preencha com os dados do aplicativo registrado pela TI no Microsoft Entra ID (Azure AD). */
window.ASBUILT_CONFIG = {
  clientId: '',            // "ID do aplicativo (cliente)" do registro no Entra ID
  tenantId: '',            // "ID do diretório (locatário)" da Vale (ou 'organizations')
  arquivo: 'https://globalvale.sharepoint.com/sites/LACAL/Shared%20Documents/FICO/FICO%20-%20AS%20BUILT/NOVA%20BD%20ASBUILT.xlsx',
  tabelaDocumentos: 'PBI',      // tabela da aba BD ASBUILT
  tabelaObrigatorios: 'PBI_9',  // tabela da aba Documentos Obrigatorios
  escopos: ['Files.Read.All']
};
