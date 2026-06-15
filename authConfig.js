/*
  Configuração do MSAL.js para Nested App Authentication (NAA).

  NAA é o método de SSO atualmente recomendado/exigido pela Microsoft
  para add-ins do Outlook (legacy Office SSO está sendo descontinuado).

  Preencha CLIENT_ID e TENANT_ID com os valores do seu App Registration
  no Entra ID (Aplicativos > Registros de aplicativo > NFe-Outlook-Addin).
*/

// Application (client) ID do seu App Registration
const CLIENT_ID = "3a6a98b3-f597-4711-8b2d-ffef4ea53850";

// Directory (tenant) ID — restringe o login à sua organização (tenant).
const TENANT_ID = "51873eff-cfd8-4715-839f-10a25cdcbec9";

// Domínio onde o add-in está hospedado (GitHub Pages)
const DOMINIO_ADDIN = "elisontown-star.github.io/NFe-Outlook-Addin";

// Scope da SUA API (Focus NFe / nfe-api). Ajuste para o Application ID
// URI definido em "Expose an API" no App Registration.
// Exemplo: "api://SEU_DOMINIO/00000000-0000-0000-0000-000000000000/access_as_user"
const API_SCOPE = "api://" + DOMINIO_ADDIN + "/" + CLIENT_ID + "/access_as_user";

const msalConfig = {
  auth: {
    clientId: CLIENT_ID,
    authority: `https://login.microsoftonline.com/${TENANT_ID}`,
    // Redirect URI registrado como SPA no App Registration.
    // Para NAA, o host (Outlook) intercepta este fluxo — não é uma
    // navegação real de página.
    redirectUri: `brk-multihub://${DOMINIO_ADDIN}`,
  },
  cache: {
    cacheLocation: "localStorage",
  },
  system: {
    allowNativeBroker: false, // NAA cuida do broker; deixar false
  },
};

// Escopos solicitados por padrão (perfil do usuário + sua API)
const LOGIN_SCOPES = ["User.Read", "openid", "profile"];
const API_SCOPES = [API_SCOPE];
