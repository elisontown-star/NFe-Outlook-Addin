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

// Edite estas URLs para apontar para os portais corretos da sua operação.
// NF-e e NFC-e variam por estado; ajuste para a SEFAZ do seu UF.
var PORTAIS = {
  nfe: "https://www.nfe.fazenda.gov.br/portal/principal.aspx",
  nfce: "https://www.nfce.fazenda.gov.br/portal/principal.aspx",
  nfse: "https://www.nfse.gov.br/EmissorNacional/",
};

Office.onReady(function () {
  var botoes = document.querySelectorAll(".atalho");
  for (var i = 0; i < botoes.length; i++) {
    botoes[i].addEventListener("click", function () {
      var tipo = this.getAttribute("data-portal");
      abrirPortal(tipo);
    });
  }

  var linkConfigurar = document.getElementById("link-configurar");
  linkConfigurar.addEventListener("click", function (evento) {
    evento.preventDefault();
    mostrarAviso(
      "Para alterar os portais, edite o arquivo taskpane.js (objeto PORTAIS)."
    );
  });
});

function abrirPortal(tipo) {
  var url = PORTAIS[tipo];

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
    function (resultado) {
      if (resultado.status === Office.AsyncResultStatus.Failed) {
        // Fallback: caso o diálogo não possa ser aberto (algumas
        // configurações de tenant bloqueiam displayDialogAsync para
        // domínios externos), tenta abrir em nova aba do navegador.
        var novaJanela = window.open(url, "_blank");
        if (!novaJanela) {
          mostrarAviso("Não foi possível abrir o portal. Copie o link: " + url);
        } else {
          mostrarAviso("Portal aberto em nova aba.");
        }
        return;
      }

      mostrarAviso("Portal aberto em nova janela.");
    }
  );
}

var timeoutAviso;
function mostrarAviso(mensagem) {
  var elemento = document.getElementById("aviso");
  elemento.textContent = mensagem;

  clearTimeout(timeoutAviso);
  timeoutAviso = setTimeout(function () {
    elemento.textContent = "";
  }, 5000);
}
