/*
  Lógica do painel: cada botão abre o portal oficial correspondente
  em uma nova janela.

  Por que displayDialogAsync e não window.open:
  - No Outlook Desktop (Windows/Mac), o task pane roda em um WebView
    isolado. window.open costuma ser bloqueado ou simplesmente não
    abre uma janela do sistema.
  - Office.context.ui.displayDialogAsync é a API suportada pela
    plataforma Office para abrir uma janela de diálogo a partir de um
    add-in, funcionando de forma consistente no Desktop e na Web.
  - A janela de diálogo carrega a URL informada; o usuário pode
    navegar normalmente dentro dela (login gov.br, certificado, etc.).
    Ela NÃO está dentro de um iframe controlado pelo add-in — é uma
    janela própria do sistema operacional/navegador.
*/

import { obterContaAtiva, login } from "./auth.js";

// Edite estas URLs para apontar para os portais corretos da sua operação.
// NF-e e NFC-e variam por estado; ajuste para a SEFAZ do seu UF.
const PORTAIS = {
  nfe: "https://www.nfe.fazenda.gov.br/portal/principal.aspx",
  nfce: "https://www.nfce.fazenda.gov.br/portal/principal.aspx",
  nfse: "https://www.nfse.gov.br/EmissorNacional/",
};

Office.onReady(() => {
  document.querySelectorAll(".atalho").forEach((botao) => {
    botao.addEventListener("click", () => {
      const tipo = botao.getAttribute("data-portal");
      abrirPortal(tipo);
    });
  });

  const linkConfigurar = document.getElementById("link-configurar");
  linkConfigurar.addEventListener("click", (evento) => {
    evento.preventDefault();
    mostrarAviso(
      "Para alterar os portais, edite o arquivo taskpane.js (objeto PORTAIS)."
    );
  });

  inicializarSessao();
});

/*
  Verifica se há uma conta autenticada via NAA. Se houver, exibe o
  nome/e-mail do usuário no cabeçalho. Caso contrário, exibe o botão
  "Entrar" para que o usuário inicie o SSO manualmente.
*/
async function inicializarSessao() {
  const statusElemento = document.getElementById("usuario-status");
  const botaoEntrar = document.getElementById("botao-entrar");

  try {
    const conta = await obterContaAtiva();

    if (conta) {
      statusElemento.textContent = conta.username || conta.name || "Conectado";
      botaoEntrar.hidden = true;
    } else {
      statusElemento.textContent = "Não conectado";
      botaoEntrar.hidden = false;
      botaoEntrar.addEventListener("click", async () => {
        botaoEntrar.disabled = true;
        botaoEntrar.textContent = "Entrando...";
        try {
          const novaConta = await login();
          statusElemento.textContent = novaConta.username || novaConta.name || "Conectado";
          botaoEntrar.hidden = true;
        } catch (erro) {
          mostrarAviso("Falha no login: " + erro.message);
          botaoEntrar.disabled = false;
          botaoEntrar.textContent = "Entrar";
        }
      });
    }
  } catch (erro) {
    // MSAL/NAA pode não estar disponível em hosts mais antigos.
    // O painel continua funcional sem SSO (apenas os atalhos de portal).
    statusElemento.textContent = "";
    console.warn("SSO indisponível neste host:", erro);
  }
}

function abrirPortal(tipo) {
  const url = PORTAIS[tipo];

  if (!url) {
    mostrarAviso("Portal não configurado para este tipo de nota.");
    return;
  }

  // displayDialogAsync exige HTTPS e que o domínio esteja listado em
  // <AppDomains> no manifest.xml, ou que o usuário aceite a navegação
  // dependendo da política do tenant.
  Office.context.ui.displayDialogAsync(
    url,
    { height: 80, width: 70, displayInIframe: false },
    (resultado) => {
      if (resultado.status === Office.AsyncResultStatus.Failed) {
        // Fallback: caso o diálogo não possa ser aberto (algumas
        // configurações de tenant bloqueiam displayDialogAsync para
        // domínios externos), tenta abrir em nova aba do navegador.
        const novaJanela = window.open(url, "_blank");
        if (!novaJanela) {
          mostrarAviso(
            "Não foi possível abrir o portal. Copie o link: " + url
          );
        } else {
          mostrarAviso("Portal aberto em nova aba.");
        }
        return;
      }

      mostrarAviso("Portal aberto em nova janela.");
    }
  );
}

let timeoutAviso;
function mostrarAviso(mensagem) {
  const elemento = document.getElementById("aviso");
  elemento.textContent = mensagem;

  clearTimeout(timeoutAviso);
  timeoutAviso = setTimeout(() => {
    elemento.textContent = "";
  }, 5000);
}
