/*
  Add-in Emissor de NF — seleção de estado + atalhos para portais oficiais.

  Fluxo:
  1. Usuário escolhe a UF (etapa 1)
  2. Aparecem os botões NF-e / NFC-e / NFS-e (etapa 2)
  3. NF-e e NFC-e abrem o portal da SEFAZ daquele estado
     NFS-e abre o Emissor Nacional (nacional, não varia por estado)

  Abertura das janelas:
  - Usa Office.context.ui.displayDialogAsync (pop-up dentro do Office).
  - Alguns portais bloqueiam ser exibidos em diálogo/iframe
    (X-Frame-Options); nesse caso o código cai para window.open.

  COMO EXPANDIR: adicione novos estados ao objeto PORTAIS abaixo,
  seguindo o mesmo formato { nfe, nfce }. A NFS-e é nacional e fica
  em PORTAL_NFSE.
*/

// Portal nacional da NFS-e (Emissor Nacional) — igual para todos os estados.
var PORTAL_NFSE = "https://www.nfse.gov.br/EmissorNacional/";

// Portais oficiais das SEFAZ, por UF.
// nfe  = portal de NF-e do estado
// nfce = portal de NFC-e do estado
var PORTAIS = {
  SP: {
    nome: "São Paulo",
    nfe: "https://portal.fazenda.sp.gov.br/servicos/nfe",
    nfce: "https://portal.fazenda.sp.gov.br/servicos/nfce",
  },
  RJ: {
    nome: "Rio de Janeiro",
    nfe: "https://portal.fazenda.rj.gov.br/dfe/",
    nfce: "https://portal.fazenda.rj.gov.br/dfe/",
  },
  MG: {
    nome: "Minas Gerais",
    nfe: "https://portalsped.fazenda.mg.gov.br/spedmg/nfe/",
    nfce: "https://portalsped.fazenda.mg.gov.br/spedmg/nfce/",
  },
  RS: {
    nome: "Rio Grande do Sul",
    nfe: "https://dfe-portal.svrs.rs.gov.br/NFE",
    nfce: "https://dfe-portal.svrs.rs.gov.br/NFCE",
  },
  PR: {
    nome: "Paraná",
    nfe: "https://www.fazenda.pr.gov.br/servicos/EMPRESA/Nota-Fiscal-Eletronica-NF-e",
    nfce: "https://www.fazenda.pr.gov.br/servicos/EMPRESA/Nota-Fiscal-de-Consumidor-Eletronica-NFC-e",
  },
  SC: {
    nome: "Santa Catarina",
    nfe: "https://sat.sef.sc.gov.br/tax.NET/Sat.Dfe.Web/Default.aspx",
    nfce: "https://sat.sef.sc.gov.br/nfce/consulta",
  },
  BA: {
    nome: "Bahia",
    nfe: "https://nfe.sefaz.ba.gov.br/servicos/nfe/default.aspx",
    nfce: "https://nfe.sefaz.ba.gov.br/servicos/nfce/Modulos/Geral/NFCEC_consulta_chave_acesso.aspx",
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

  // Liga os botões de portal
  var botoes = document.querySelectorAll(".atalho");
  for (var i = 0; i < botoes.length; i++) {
    botoes[i].addEventListener("click", function () {
      abrirPortal(this.getAttribute("data-portal"));
    });
  }
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

function abrirPortal(tipo) {
  var url;

  if (tipo === "nfse") {
    // NFS-e é nacional
    url = PORTAL_NFSE;
  } else if (ufSelecionada && PORTAIS[ufSelecionada]) {
    url = PORTAIS[ufSelecionada][tipo];
  }

  if (!url) {
    mostrarAviso("Portal não disponível para este estado/tipo.");
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
