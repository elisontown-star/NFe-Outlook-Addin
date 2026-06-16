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

/* ===================== Navegação por abas ===================== */
(function () {
  var abaPortais = document.getElementById("aba-portais");
  var abaEmitir = document.getElementById("aba-emitir");
  var conteudoPortais = document.getElementById("conteudo-portais");
  var conteudoEmitir = document.getElementById("conteudo-emitir");

  function ativar(aba, conteudo, outraAba, outroConteudo) {
    aba.classList.add("aba--ativa");
    aba.setAttribute("aria-selected", "true");
    outraAba.classList.remove("aba--ativa");
    outraAba.setAttribute("aria-selected", "false");
    conteudo.hidden = false;
    outroConteudo.hidden = true;
  }

  if (abaPortais && abaEmitir) {
    abaPortais.addEventListener("click", function () {
      ativar(abaPortais, conteudoPortais, abaEmitir, conteudoEmitir);
    });
    abaEmitir.addEventListener("click", function () {
      ativar(abaEmitir, conteudoEmitir, abaPortais, conteudoPortais);
    });
  }
})();

/* ===================== Formulário NFS-e ===================== */
(function () {
  var form = document.getElementById("form-nfse");
  if (!form) return;

  var btnEmitir = document.getElementById("btn-emitir-nfse");
  var resultadoDps = document.getElementById("resultado-dps");
  var dpsOutput = document.getElementById("dps-output");
  var btnCopiarDps = document.getElementById("btn-copiar-dps");

  function urlApi() {
    return document.getElementById("api-url").value.replace(/\/$/, "");
  }

  function coletarDados() {
    var fd = new FormData(form);
    return {
      prestador_cnpj: fd.get("prestador_cnpj"),
      prestador_inscricao_municipal: fd.get("prestador_inscricao_municipal"),
      prestador_codigo_municipio: fd.get("prestador_codigo_municipio"),
      tomador_cnpj: fd.get("tomador_cnpj"),
      tomador_razao_social: fd.get("tomador_razao_social"),
      tomador_email: fd.get("tomador_email") || null,
      valor_servicos: parseFloat(fd.get("valor_servicos")),
      aliquota_iss: parseFloat(fd.get("aliquota_iss")),
      item_lista_servico: fd.get("item_lista_servico"),
      discriminacao: fd.get("discriminacao"),
      iss_retido: fd.get("iss_retido") === "on",
      codigo_tributario_municipio: fd.get("codigo_tributario_municipio") || null,
    };
  }

  // Gerar DPS (XML) — funciona já, só monta o XML
  form.addEventListener("submit", function (e) {
    e.preventDefault();
    var dados = coletarDados();

    fetch(urlApi() + "/nfse/gerar-dps", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(dados),
    })
      .then(function (resp) {
        if (!resp.ok) return resp.json().then(function (e) { throw e; });
        return resp.json();
      })
      .then(function (data) {
        dpsOutput.textContent = data.dps_xml;
        resultadoDps.hidden = false;
        mostrarAviso("DPS (XML) gerada com sucesso.");
      })
      .catch(function (err) {
        mostrarAviso("Erro ao gerar DPS. Verifique se a API está acessível.");
      });
  });

  // Emitir — preparado, mas avisa que precisa de certificado/hospedagem
  if (btnEmitir) {
    btnEmitir.addEventListener("click", function () {
      if (!form.reportValidity()) return;
      mostrarAviso(
        "A emissão será ativada quando a API estiver hospedada e o " +
        "certificado digital A1 configurado. Por ora, use 'Gerar DPS' para conferir o XML."
      );
    });
  }

  if (btnCopiarDps) {
    btnCopiarDps.addEventListener("click", function () {
      navigator.clipboard.writeText(dpsOutput.textContent).then(function () {
        btnCopiarDps.textContent = "Copiado!";
        setTimeout(function () { btnCopiarDps.textContent = "Copiar"; }, 2000);
      });
    });
  }

  // ===================== Autopreenchimento por CNPJ =====================
  function soDigitos(v) {
    return (v || "").replace(/\D/g, "");
  }

  function autoPreencherCnpj(tipo) {
    var prefixo = tipo === "prestador" ? "prestador" : "tomador";
    var campoCnpj = form.querySelector("[name=" + prefixo + "_cnpj]");
    var cnpj = soDigitos(campoCnpj.value);
    if (cnpj.length !== 14) return;

    var valorOriginal = campoCnpj.value;
    campoCnpj.disabled = true;
    campoCnpj.value = "Consultando...";

    fetch(urlApi() + "/cnpj/" + cnpj)
      .then(function (resp) {
        campoCnpj.value = valorOriginal;
        if (!resp.ok) throw new Error("CNPJ não encontrado");
        return resp.json();
      })
      .then(function (d) {
        if (tipo === "prestador") {
          if (d.codigo_municipio_ibge)
            form.querySelector("[name=prestador_codigo_municipio]").value = d.codigo_municipio_ibge;
        } else {
          if (d.razao_social)
            form.querySelector("[name=tomador_razao_social]").value = d.razao_social;
          if (d.email)
            form.querySelector("[name=tomador_email]").value = d.email;
        }
        mostrarAviso("Dados de " + (d.razao_social || cnpj) + " carregados.");
      })
      .catch(function () {
        campoCnpj.value = valorOriginal;
        mostrarAviso("CNPJ não encontrado ou API indisponível.");
      })
      .finally(function () {
        campoCnpj.disabled = false;
      });
  }

  var cnpjPrest = form.querySelector("[name=prestador_cnpj]");
  var cnpjTom = form.querySelector("[name=tomador_cnpj]");
  if (cnpjPrest) cnpjPrest.addEventListener("blur", function () { autoPreencherCnpj("prestador"); });
  if (cnpjTom) cnpjTom.addEventListener("blur", function () { autoPreencherCnpj("tomador"); });
})();
