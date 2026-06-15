/*
  Módulo de autenticação: tenta SSO via NAA (MSAL.js PublicClientNext)
  e, se indisponível (host antigo, política bloqueando), cai para o
  fluxo de diálogo do Office (Office.context.ui.displayDialogAsync
  apontando para auth.html com MSAL redirect tradicional).

  Uso no taskpane.js:
    import { obterContaAtiva, login, obterTokenApi } from "./auth.js";
*/

let msalInstance;

async function inicializarMsal() {
  if (msalInstance) return msalInstance;

  // PublicClientNext é a API do MSAL.js que suporta NAA dentro de hosts
  // Office. Requer @azure/msal-browser >= 3.7.0 com suporte experimental NAA.
  msalInstance = await msal.PublicClientApplication.createPublicClientApplication(msalConfig);
  return msalInstance;
}

export async function obterContaAtiva() {
  const instancia = await inicializarMsal();
  const contas = instancia.getAllAccounts();
  return contas.length > 0 ? contas[0] : null;
}

export async function login() {
  const instancia = await inicializarMsal();

  try {
    const resultado = await instancia.acquireTokenSilent({
      scopes: LOGIN_SCOPES,
    });
    return resultado.account;
  } catch (erroSilencioso) {
    // Sem sessão em cache: tenta SSO interativo via NAA.
    try {
      const resultado = await instancia.acquireTokenPopup({
        scopes: LOGIN_SCOPES,
      });
      return resultado.account;
    } catch (erroNaa) {
      console.warn("NAA indisponível, usando fallback de diálogo:", erroNaa);
      return await loginViaDialogoFallback();
    }
  }
}

export async function obterTokenApi() {
  const instancia = await inicializarMsal();
  const conta = await obterContaAtiva();

  if (!conta) {
    throw new Error("Usuário não autenticado. Chame login() primeiro.");
  }

  try {
    const resultado = await instancia.acquireTokenSilent({
      scopes: API_SCOPES,
      account: conta,
    });
    return resultado.accessToken;
  } catch (erro) {
    const resultado = await instancia.acquireTokenPopup({
      scopes: API_SCOPES,
      account: conta,
    });
    return resultado.accessToken;
  }
}

/*
  Fallback para hosts/versões do Outlook que ainda não suportam NAA.
  Abre auth.html em um diálogo do Office, que executa o fluxo MSAL
  tradicional (redirect) e devolve o resultado via messageParent.
*/
function loginViaDialogoFallback() {
  return new Promise((resolve, reject) => {
    Office.context.ui.displayDialogAsync(
      `https://${DOMINIO_ADDIN}/auth.html`,
      { height: 60, width: 40 },
      (resultado) => {
        if (resultado.status === Office.AsyncResultStatus.Failed) {
          reject(new Error("Não foi possível abrir o diálogo de login: " + resultado.error.message));
          return;
        }

        const dialog = resultado.value;

        dialog.addEventHandler(Office.EventType.DialogMessageReceived, (args) => {
          dialog.close();
          try {
            const dados = JSON.parse(args.message);
            if (dados.erro) {
              reject(new Error(dados.erro));
            } else {
              resolve(dados.conta);
            }
          } catch (e) {
            reject(e);
          }
        });

        dialog.addEventHandler(Office.EventType.DialogEventReceived, () => {
          reject(new Error("Diálogo de login fechado pelo usuário."));
        });
      }
    );
  });
}
