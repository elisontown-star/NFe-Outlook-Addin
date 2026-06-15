# Add-in do Outlook — Emissor de NF (atalhos para portais oficiais)

Suplemento (task pane) para Outlook Desktop e Outlook na Web que exibe
botões de atalho para abrir os portais oficiais de emissão de **NF-e**,
**NFC-e** e **NFS-e** em uma nova janela.

## Por que não embute o site do governo

Os portais gov.br/SEFAZ bloqueiam ser exibidos dentro de um iframe de
outro domínio (proteção `X-Frame-Options` / `frame-ancestors` contra
clickjacking). Por isso este add-in não tenta renderizar o portal dentro
do painel — ele abre uma **janela própria** via
`Office.context.ui.displayDialogAsync`, que é a forma suportada pela
plataforma Office.

## Estrutura

```
NFe-Outlook-Addin/
├── manifest.xml        # Manifest do add-in (XML, formato clássico)
├── taskpane.html        # Painel exibido no Outlook
├── taskpane.css         # Estilos
├── taskpane.js          # Lógica dos botões / abertura dos portais
├── commands.html        # Function file exigido pelo manifest
├── auth.html            # Fallback de login (diálogo MSAL)
├── auth.js              # Lógica de login NAA / MSAL
├── authConfig.js        # Configuração MSAL (client ID, tenant ID, etc.)
└── assets/
    ├── icon.svg
    ├── icon-16.png
    ├── icon-32.png
    ├── icon-64.png
    ├── icon-80.png
    └── icon-128.png
```

Estrutura plana (sem subpastas) propositalmente, para facilitar o upload
direto ao GitHub e a hospedagem via GitHub Pages na raiz do repositório.

## 0. Autenticação (SSO via Nested App Authentication)

A Microsoft está descontinuando o SSO legado dos add-ins do Outlook em
favor do **NAA (Nested App Authentication)**, usando MSAL.js. Os arquivos
em os arquivos de autenticação (`auth.js`, `auth.html`, `authConfig.js`) já implementam esse fluxo, com fallback para diálogo em
hosts mais antigos.

### Dados já configurados neste projeto

- **Client ID**: `3a6a98b3-f597-4711-8b2d-ffef4ea53850`
- **Tenant ID**: `51873eff-cfd8-4715-839f-10a25cdcbec9`
- **Domínio (GitHub Pages)**: `elisontown-star.github.io/NFe-Outlook-Addin`

Esses valores já estão preenchidos em `authConfig.js` e em
`manifest.xml`.

### Checklist no App Registration (Entra ID)

No registro `NFe-Outlook-Addin` (entra.microsoft.com → Aplicativos →
Registros de aplicativo), confirme:

1. **Autenticação** → plataforma **SPA**, com as seguintes Redirect URIs:
   - `brk-multihub://elisontown-star.github.io/NFe-Outlook-Addin`
   - `https://elisontown-star.github.io/NFe-Outlook-Addin/auth.html`
2. **Permissões de API**: `User.Read`, `openid`, `profile` (Microsoft
   Graph) — vêm por padrão.
3. (Opcional, para autenticar chamadas à sua API NFe) em **Expose an
   API**:
   - Application ID URI: `api://elisontown-star.github.io/NFe-Outlook-Addin/3a6a98b3-f597-4711-8b2d-ffef4ea53850`
   - Scope `access_as_user`, estado **Habilitado**.
   - Esse valor já corresponde ao `API_SCOPE` configurado em `authConfig.js`.

### Comportamento no painel

- Ao abrir, o painel tenta obter a conta já autenticada no Outlook via
  NAA (sem pedir login novamente).
- Se não houver sessão, exibe um botão **Entrar** que dispara o
  `acquireTokenPopup` (NAA) ou, em hosts sem suporte, abre
  `auth.html` em um diálogo do Office com o fluxo MSAL
  tradicional (`loginRedirect`).
- O e-mail/nome do usuário aparece no cabeçalho do painel.
- `obterTokenApi()` (em `auth.js`) retorna um access token para
  o `API_SCOPE` configurado, pronto para autenticar chamadas à sua API
  NFe (Focus NFe) em nome do usuário logado.

## 1. Configurar os links dos portais

Edite `taskpane.js`, objeto `PORTAIS`:

```js
const PORTAIS = {
  nfe:  "https://www.nfe.fazenda.gov.br/portal/principal.aspx",  // ajuste para a SEFAZ do seu estado
  nfce: "https://www.nfce.fazenda.gov.br/portal/principal.aspx", // ajuste para a SEFAZ do seu estado
  nfse: "https://www.nfse.gov.br/EmissorNacional/",
};
```

NF-e e NFC-e têm portais próprios por estado — substitua pelas URLs da
SEFAZ do seu UF. NFS-e usa o Emissor Nacional (padrão para a maioria dos
municípios na Reforma Tributária).

## 2. Hospedar os arquivos (GitHub Pages)

Este repositório já está configurado para
`https://elisontown-star.github.io/NFe-Outlook-Addin`.

Para publicar via GitHub Pages:

1. Faça push deste conteúdo para o repositório
   `elisontown-star/NFe-Outlook-Addin` (branch `main`).
2. No GitHub: **Settings** → **Pages** → **Source**: `Deploy from a
   branch` → Branch: `main` / pasta `/ (root)` → **Save**.
3. Aguarde alguns minutos; o site ficará disponível em
   `https://elisontown-star.github.io/NFe-Outlook-Addin/`.
4. Adicione os ícones (16/32/80/64/128px) em `assets/`
   com os nomes `icon-16.png`, `icon-32.png`, `icon-80.png`,
   `icon-64.png`, `icon-128.png` — o manifest já referencia esses
   caminhos.

> **Importante**: o GitHub Pages serve HTTPS automaticamente, requisito
> obrigatório para Office Add-ins.

## 3. Testar localmente (sideload)

### Outlook na Web
1. Hospede os arquivos (passo 2) ou use `npx office-addin-dev-certs` +
   um servidor local HTTPS para testes (`https://localhost:3000`).
2. Acesse outlook.office.com → **Configurações** → **Ver todas as
   configurações do Outlook** → **Suplementos** → **Meus suplementos**
   → **Adicionar um suplemento personalizado** → **Adicionar do arquivo**
   → selecione `manifest.xml`.
3. Abra um e-mail; o botão "Emitir NF" aparece na faixa de opções.

### Outlook Desktop (Windows/Mac)
1. Mesma hospedagem HTTPS.
2. **Arquivo** → **Gerenciar Suplementos** (abre o navegador na mesma
   tela de suplementos do Outlook na Web) → repita o passo 2 acima.
   Suplementos sideload na Web sincronizam com o Desktop conectado à
   mesma conta.

## 4. Comportamento esperado

- Usuário abre um e-mail → clica em **Emitir NF** na faixa de opções →
  painel lateral abre.
- Clica em **NF-e**, **NFC-e** ou **NFS-e** → uma janela de diálogo
  (`displayDialogAsync`) abre o portal oficial correspondente.
- Se o tenant/política bloquear `displayDialogAsync` para domínios
  externos, o código cai para `window.open` como fallback.

## Próximos passos possíveis

- Extrair dados do e-mail (remetente, CNPJ no corpo) e copiar para a
  área de transferência antes de abrir o portal, para colar nos campos.
- Trocar os atalhos por chamadas à API Focus NFe (emissão automática
  sem sair do Outlook) — ver projeto `nfe-api`.
