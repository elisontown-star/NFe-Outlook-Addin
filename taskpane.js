/*
  Add-in Emissor de NF — seleção de estado + Portal do Contribuinte.

  Fluxo:
  1. Usuário escolhe a UF (etapa 1)
  2. Aparece um único botão "Portal do Contribuinte" (etapa 2)
  3. O botão abre o portal do contribuinte da SEFAZ daquele estado

  Abertura das janelas:
  - Usa Office.context.ui.displayDialogAsync (pop-up dentro do Office).
  - Alguns portais bloqueiam ser exibidos em diálogo/iframe
    (X-Frame-Options); nesse caso o código cai para window.open.

  COMO EXPANDIR: adicione novos estados ao objeto PORTAIS abaixo,
  seguindo o formato { nome, url } e inclua a sigla em ORDEM_UF.
*/

// Portal do Contribuinte oficial de cada SEFAZ, por UF.
var PORTAIS = {
  SP: {
    nome: "São Paulo",
    url: "https://www.fazenda.sp.gov.br/PFE",
  },
  RJ: {
    nome: "Rio de Janeiro",
    url: "https://atendimentodigitalrj.fazenda.rj.gov.br/login",
  },
  MG: {
    nome: "Minas Gerais",
    url: "https://www2.fazenda.mg.gov.br/sol/",
  },
  RS: {
    nome: "Rio Grande do Sul",
    url: "https://www.sefaz.rs.gov.br/Receita/Portal",
  },
  PR: {
    nome: "Paraná",
    url: "https://receita.pr.gov.br/",
  },
  SC: {
    nome: "Santa Catarina",
    url: "https://sat.sef.sc.gov.br/",
  },
  BA: {
    nome: "Bahia",
    url: "https://portal.sefaz.ba.gov.br/scripts/sat/login.asp",
  },
};

// Ordem de exibição no dropdown
var ORDEM_UF = ["SP", "RJ", "MG", "RS", "PR", "SC", "BA"];

var ufSelecionada = null;

Office.onReady(function () {
  popularSelectUf();

  var select = document.getElementById("select-uf");
  var btnContinuar = document.getElementById("btn-continuar");
  var btnTrocar = document.getElementById("btn-trocar-uf");
  var btnPortal = document.getElementById("btn-portal");

  select.addEventListener("change", function () {
    btnContinuar.disabled = !select.value;
  });

  btnContinuar.addEventListener("click", function () {
    if (select.value) {
      ufSelecionada = select.value;
      mostrarEtapaPortais();
    }
  });

  btnTrocar.addEventListener("click", function () {
    mostrarEtapaEstado();
  });

  btnPortal.addEventListener("click", function () {
    abrirPortalContribuinte();
  });
});

function popularSelectUf() {
  var select = document.getElementById("select-uf");
  for (var i = 0; i < ORDEM_UF.length; i++) {
    var uf = ORDEM_UF[i];
    var opt = document.createElement("option");
    opt.value = uf;
    opt.textContent = uf + " — " + PORTAIS[uf].nome;
    select.appendChild(opt);
  }
}

function mostrarEtapaPortais() {
  document.getElementById("etapa-estado").hidden = true;
  document.getElementById("etapa-portais").hidden = false;
  document.getElementById("uf-atual").textContent =
    ufSelecionada + " — " + PORTAIS[ufSelecionada].nome;
}

function mostrarEtapaEstado() {
  document.getElementById("etapa-portais").hidden = true;
  document.getElementById("etapa-estado").hidden = false;
}

function abrirPortalContribuinte() {
  var url = ufSelecionada && PORTAIS[ufSelecionada]
    ? PORTAIS[ufSelecionada].url
    : null;

  if (!url) {
    mostrarAviso("Portal não disponível para este estado.");
    return;
  }

  Office.context.ui.displayDialogAsync(
    url,
    { height: 80, width: 70, displayInIframe: false },
    function (resultado) {
      if (resultado.status === Office.AsyncResultStatus.Failed) {
        var nova = window.open(url, "_blank");
        if (!nova) {
          mostrarAviso("Não foi possível abrir. Link: " + url);
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
